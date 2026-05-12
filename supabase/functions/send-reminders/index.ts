// FinFlow – Daily reminder Edge Function
// Triggered by Supabase pg_cron (see README for setup)
// Env vars required: SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, RESEND_API_KEY

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const CARRIER_GATEWAYS: Record<string, string> = {
  att:       "@txt.att.net",
  tmobile:   "@tmomail.net",
  verizon:   "@vtext.com",
  sprint:    "@messaging.sprintpcs.com",
  boost:     "@smsmyboostmobile.com",
  cricket:   "@mms.cricketwireless.net",
  googlefi:  "@msg.fi.google.com",
  metro:     "@mymetropcs.com",
  uscellular:"@email.uscc.net",
};

const supabase = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!
);

async function getSetting(key: string): Promise<string> {
  const { data } = await supabase.from("settings").select("value").eq("key", key).single();
  return data?.value ?? "";
}

async function log(type: string, recipient: string, subject: string, status: string, message: string) {
  await supabase.from("reminder_log").insert({ type, recipient, subject, status, message });
}

async function sendEmail(to: string, subject: string, html: string, text: string) {
  const apiKey = await getSetting("resend_api_key");
  if (!apiKey) { await log("EMAIL", to, subject, "SKIPPED", "No Resend API key configured"); return; }

  const res = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: { Authorization: `Bearer ${apiKey}`, "Content-Type": "application/json" },
    body: JSON.stringify({ from: "FinFlow <reminders@resend.dev>", to, subject, html, text }),
  });

  const status = res.ok ? "Sent" : `FAILED ${res.status}`;
  await log("EMAIL", to, subject, status, await res.text());
}

async function sendSms(message: string, phoneNumber: string, carrier: string) {
  const gateway = CARRIER_GATEWAYS[carrier];
  if (!gateway) { await log("SMS", phoneNumber, "SMS", "SKIPPED", `Unknown carrier: ${carrier}`); return; }
  const to = phoneNumber.replace(/\D/g, "") + gateway;
  // Send via Resend to the carrier email gateway address
  await sendEmail(to, "FinFlow", `<p>${message}</p>`, message);
}

function urgencyColor(days: number): string {
  if (days === 0) return "#ff5b6b";
  if (days <= 3)  return "#ff9f43";
  return "#5c7cff";
}

function billEmailHtml(bill: Record<string, unknown>, daysUntil: number): string {
  const label = daysUntil === 0 ? "🚨 DUE TODAY" : daysUntil === 1 ? "⚠️ Due Tomorrow" : `📅 Due in ${daysUntil} days`;
  const color = urgencyColor(daysUntil);
  const util = (bill.credit_limit as number) > 0
    ? `<tr><td style="padding:9px 12px;font-weight:bold;background:#f8f9fa;">Credit Utilization</td><td style="padding:9px 12px;">${((bill.balance as number) / (bill.credit_limit as number) * 100).toFixed(1)}%</td></tr>`
    : "";

  return `<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0f2f5;">
<div style="font-family:'Helvetica Neue',Arial,sans-serif;max-width:600px;margin:30px auto;background:#f0f2f5;padding:20px;">
<div style="background:linear-gradient(135deg,#1a1a2e,#16213e);padding:28px;border-radius:12px 12px 0 0;">
  <div style="color:#5c7cff;font-size:13px;font-weight:600;letter-spacing:1px;margin-bottom:6px;">FINFLOW</div>
  <h1 style="color:white;margin:0;font-size:22px;">💳 Bill Payment Reminder</h1>
</div>
<div style="background:white;padding:28px;border-radius:0 0 12px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08);">
  <div style="background:${color}18;border-left:4px solid ${color};padding:14px 18px;border-radius:6px;margin-bottom:24px;">
    <span style="color:${color};font-size:18px;font-weight:bold;">${label}</span>
  </div>
  <table style="width:100%;border-collapse:collapse;font-size:14px;">
    <tr><td style="padding:9px 12px;font-weight:bold;background:#f8f9fa;">Account</td><td style="padding:9px 12px;">${bill.name}</td></tr>
    <tr><td style="padding:9px 12px;font-weight:bold;">Type</td><td style="padding:9px 12px;">${bill.type}</td></tr>
    <tr style="background:#f8f9fa;"><td style="padding:9px 12px;font-weight:bold;">Amount Due</td><td style="padding:9px 12px;color:#ff5b6b;font-size:20px;font-weight:bold;">$${(bill.min_payment as number).toFixed(2)}</td></tr>
    <tr><td style="padding:9px 12px;font-weight:bold;">Due Day</td><td style="padding:9px 12px;">Day ${bill.due_day} of each month</td></tr>
    <tr style="background:#f8f9fa;"><td style="padding:9px 12px;font-weight:bold;">Balance</td><td style="padding:9px 12px;">$${(bill.balance as number).toFixed(2)}</td></tr>
    <tr><td style="padding:9px 12px;font-weight:bold;">APR</td><td style="padding:9px 12px;">${bill.apr}%</td></tr>
    ${util}
  </table>
  <p style="color:#aaa;font-size:11px;margin-top:24px;border-top:1px solid #eee;padding-top:16px;text-align:center;">
    Sent by FinFlow · manage reminders at your FinFlow app
  </p>
</div></div></body></html>`;
}

Deno.serve(async () => {
  try {
    const today = new Date();
    const currentDay = today.getDate();
    const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();

    const enableEmail  = await getSetting("enable_email") !== "false";
    const enableSms    = await getSetting("enable_sms") === "true";
    const email        = await getSetting("email");
    const phoneNumber  = await getSetting("phone_number");
    const carrier      = await getSetting("phone_carrier");
    const graceDays    = parseInt(await getSetting("overdue_grace_days")) || 3;
    const dueTodayAlert= await getSetting("send_due_today") !== "false";
    const overdueEnabled = await getSetting("overdue_alert_enabled") !== "false";
    const reminderDays = (await getSetting("reminder_days") || "7,3,1")
      .split(",").map((x) => parseInt(x.trim()));

    const { data: debts } = await supabase
      .from("debts")
      .select("*")
      .eq("included", true);

    const upcoming: Array<Record<string, unknown>> = [];
    const overdue:  Array<Record<string, unknown>> = [];

    for (const debt of debts ?? []) {
      const due = parseInt(debt.due_day);
      if (!due) continue;

      // Days until next occurrence with month wraparound
      const daysUntil = due >= currentDay
        ? due - currentDay
        : (daysInMonth - currentDay) + due;

      if (daysUntil === 0 && dueTodayAlert) {
        upcoming.push({ ...debt, daysUntilDue: 0 });
      } else if (daysUntil > 0 && reminderDays.includes(daysUntil)) {
        upcoming.push({ ...debt, daysUntilDue: daysUntil });
      }

      if (overdueEnabled && due < currentDay && (currentDay - due) > graceDays) {
        overdue.push({ ...debt, daysOverdue: currentDay - due });
      }
    }

    // Send upcoming reminders
    for (const bill of upcoming) {
      const days = bill.daysUntilDue as number;
      const label = days === 0 ? "🚨 DUE TODAY" : days === 1 ? "⚠️ Due Tomorrow" : `📅 Due in ${days} days`;
      const subject = `${label}: ${bill.name} – $${(bill.min_payment as number).toFixed(2)}`;

      if (enableEmail && email) await sendEmail(email, subject, billEmailHtml(bill, days), `${label}: ${bill.name} $${(bill.min_payment as number).toFixed(2)} due day ${bill.due_day}`);
      if (enableSms && phoneNumber) await sendSms(`FinFlow: ${bill.name} ${days === 0 ? "DUE TODAY" : `due in ${days}d`} $${(bill.min_payment as number).toFixed(2)}`, phoneNumber, carrier);
    }

    // Send overdue alert (batched)
    if (overdue.length > 0) {
      const names = overdue.map((b) => b.name).join(", ");
      const subject = `🚨 FinFlow: ${overdue.length} bill(s) may be overdue`;

      if (enableEmail && email) {
        const rows = overdue.map((b) =>
          `<tr><td style="padding:9px 12px;border-bottom:1px solid #eee;">${b.name}</td>
           <td style="padding:9px 12px;border-bottom:1px solid #eee;color:#ff5b6b;font-weight:bold;">$${(b.min_payment as number).toFixed(2)}</td>
           <td style="padding:9px 12px;border-bottom:1px solid #eee;color:#ff5b6b;">${b.daysOverdue} days ago</td></tr>`
        ).join("");
        const html = `<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0f2f5;">
<div style="font-family:Arial,sans-serif;max-width:600px;margin:30px auto;">
<div style="background:#ff5b6b;padding:24px;border-radius:12px 12px 0 0;"><h1 style="color:white;margin:0;">🚨 Overdue Bills</h1></div>
<div style="background:white;padding:24px;border-radius:0 0 12px 12px;">
<table style="width:100%;border-collapse:collapse;font-size:13px;">
<thead><tr style="background:#1a1a2e;color:white;"><th style="padding:10px 12px;text-align:left;">Account</th><th style="padding:10px 12px;text-align:left;">Amount</th><th style="padding:10px 12px;text-align:left;">Overdue</th></tr></thead>
<tbody>${rows}</tbody></table></div></div></body></html>`;
        await sendEmail(email, subject, html, `Overdue: ${names}`);
      }
      if (enableSms && phoneNumber) await sendSms(`FinFlow OVERDUE: ${names}`, phoneNumber, carrier);
    }

    // Weekly digest — send on Mondays (weekday 1)
    if (today.getDay() === 1) {
      const digestEnabled = await getSetting("weekly_digest_enabled") !== "false";
      if (digestEnabled && enableEmail && email) {
        const { data: inc } = await supabase.from("income").select("*");
        const { data: exp } = await supabase.from("expenses").select("*");
        const monthlyIncome = (inc ?? []).reduce((s: number, r: {amount: number; frequency: string}) => {
          const f = r.frequency;
          if (f === "bi-weekly") return s + r.amount * 26 / 12;
          if (f === "weekly")    return s + r.amount * 52 / 12;
          if (f === "annual")    return s + r.amount / 12;
          return s + r.amount;
        }, 0);
        const monthlyExp = (exp ?? []).reduce((s: number, r: {amount: number}) => s + r.amount, 0);

        const billsThisWeek = (debts ?? []).filter((d) => {
          const until = d.due_day >= currentDay ? d.due_day - currentDay : (daysInMonth - currentDay) + d.due_day;
          return until <= 7;
        });

        const subject = `📋 FinFlow Weekly Digest – ${today.toLocaleDateString("en-US", { month: "long", day: "numeric" })}`;
        const text = `Income: $${monthlyIncome.toFixed(2)}\nExpenses: $${monthlyExp.toFixed(2)}\nBills this week: ${billsThisWeek.length}`;
        const html = `<!DOCTYPE html><html><body style="background:#f0f2f5;margin:0;padding:20px;font-family:Arial,sans-serif;">
<div style="max-width:560px;margin:0 auto;">
<div style="background:linear-gradient(135deg,#1a1a2e,#16213e);padding:24px;border-radius:12px 12px 0 0;">
  <h1 style="color:white;margin:0;font-size:20px;">📋 Weekly Digest</h1>
  <p style="color:#a0a0b0;margin:6px 0 0;">${today.toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric" })}</p>
</div>
<div style="background:white;padding:24px;border-radius:0 0 12px 12px;">
  <table style="width:100%;margin-bottom:20px;">
    <tr>
      <td style="padding:12px;background:#f0fff4;border-radius:8px;text-align:center;"><div style="font-size:11px;color:#888;">Income</div><div style="font-size:20px;font-weight:bold;color:#22d47e;">$${monthlyIncome.toFixed(0)}</div></td>
      <td style="width:12px;"></td>
      <td style="padding:12px;background:#fff0f1;border-radius:8px;text-align:center;"><div style="font-size:11px;color:#888;">Expenses</div><div style="font-size:20px;font-weight:bold;color:#ff5b6b;">$${monthlyExp.toFixed(0)}</div></td>
    </tr>
  </table>
  <h3 style="font-size:14px;border-bottom:2px solid #a78bfa;padding-bottom:8px;">Bills Due This Week (${billsThisWeek.length})</h3>
  ${billsThisWeek.length === 0 ? '<p style="color:#aaa;text-align:center;">🎉 No bills due this week!</p>' :
    billsThisWeek.map((b) => {
      const until = b.due_day >= currentDay ? b.due_day - currentDay : (daysInMonth - currentDay) + b.due_day;
      return `<div style="display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid #f0f0f0;"><span>${b.name}</span><span style="font-weight:bold;">$${(b.min_payment as number).toFixed(2)} · ${until === 0 ? "TODAY" : `in ${until}d`}</span></div>`;
    }).join("")}
  <p style="color:#aaa;font-size:11px;margin-top:20px;text-align:center;">FinFlow Weekly Digest</p>
</div></div></body></html>`;
        await sendEmail(email, subject, html, text);
      }
    }

    await log("BATCH_CHECK", email, `${upcoming.length} reminders, ${overdue.length} overdue`, "OK", "Daily check completed");
    return new Response(JSON.stringify({ ok: true, upcoming: upcoming.length, overdue: overdue.length }), {
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: String(err) }), { status: 500, headers: { "Content-Type": "application/json" } });
  }
});
