# FinFlow – Google Apps Script Bill Reminder

Automated bill reminders, budget alerts, and financial reports — backed by Google Sheets.

---

## Features

| Feature | Details |
|---|---|
| **Bill Reminders** | Email + SMS alerts 7/3/1 days before due date, and on the due day |
| **Overdue Alerts** | Flags unpaid bills past a configurable grace period |
| **Budget Alerts** | Warns when obligations exceed X% of income or credit utilization is high |
| **Weekly Digest** | Monday morning snapshot of upcoming bills and cash flow |
| **Monthly Report** | Full financial report with debt table, goal progress, and budget breakdown |
| **Payoff Planner** | Avalanche (highest APR first) or Snowball (lowest balance first) with month-by-month schedule |
| **Savings Goals** | Track multiple goals with progress bars and time-to-completion |
| **12/24/36-Month Forecast** | Month-by-month income, expenses, debt paydown, and surplus projection |
| **Phone Notifications** | SMS via carrier email gateway, Pushover push notifications, or any webhook (IFTTT / Zapier / Make) |
| **Dashboard Sidebar** | Live KPIs, upcoming bills, and goal progress inside Google Sheets |
| **Settings Sidebar** | Full settings UI without touching the Settings sheet directly |
| **Reminder Log** | Every notification is logged to the "Reminder Log" sheet with status |

---

## Setup (5 minutes)

### Step 1 – Create a new Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a blank spreadsheet
2. Name it something like **FinFlow**

### Step 2 – Open Apps Script

1. In your sheet, click **Extensions → Apps Script**
2. Delete all existing code in `Code.gs`

### Step 3 – Paste the script files

Copy each file from this folder into Apps Script:

| This file → | Create this file in Apps Script |
|---|---|
| `Code.gs` | Paste into the default `Code.gs` |
| `Dashboard.html` | Click **+** → **HTML** → name it `Dashboard` |
| `Settings.html` | Click **+** → **HTML** → name it `Settings` |

Also update `appsscript.json`:
1. Click **Project Settings** (gear icon) → check **Show "appsscript.json"**
2. Click `appsscript.json` in the file list and replace its contents with the provided `appsscript.json`
3. Change `"timeZone"` to your timezone (e.g. `"America/Los_Angeles"`)

### Step 4 – Save and authorize

1. Press **Ctrl+S** (or Cmd+S) to save
2. Click **Run → onOpen** (or any function) to trigger the authorization prompt
3. Click **Review permissions → Advanced → Go to FinFlow (unsafe)** → **Allow**

   The script needs permission to: send Gmail, access your spreadsheet, and make external HTTP requests (for Pushover/webhook).

### Step 5 – Initialize the spreadsheet

1. Go back to your Google Sheet
2. Refresh the page — you'll see a **💳 FinFlow** menu appear
3. Click **💳 FinFlow → 🔧 Setup Sheets**
4. This creates all the sheets and fills in sample data

### Step 6 – Configure your settings

1. Click **💳 FinFlow → ⚙️ Settings** to open the settings sidebar, OR
2. Edit the **Settings** sheet directly

**Essential settings:**
- `email` — your Gmail address (pre-filled with your Google account)
- `phoneEmail` / `phoneNumber` + `phoneCarrier` — for SMS reminders (see below)
- `reminderDays` — e.g. `7,3,1` (7 days, 3 days, 1 day before due)

### Step 7 – Enter your real data

Replace the sample data in each sheet:

| Sheet | What to enter |
|---|---|
| **Debts** | Your credit cards, loans, etc. — balance, APR, min payment, due day |
| **Income** | All income sources and their frequency |
| **Expenses** | Monthly recurring expenses by category |
| **Goals** | Savings targets with current amount and monthly contribution |

### Step 8 – Enable automated triggers

Click **💳 FinFlow → ⏰ Setup Automated Triggers**

This creates four time-based triggers:
- **9:00 AM daily** → Bill reminder check
- **8:00 AM daily** → Budget alert check
- **Monday 8:00 AM** → Weekly digest
- **1st of month 7:00 AM** → Monthly report

---

## Phone Notification Setup

### Option A: SMS via carrier email gateway (free, no extra accounts)

Every US carrier provides an email address that delivers as a text message.

| Carrier | Gateway format |
|---|---|
| AT&T | `1234567890@txt.att.net` |
| T-Mobile | `1234567890@tmomail.net` |
| Verizon | `1234567890@vtext.com` |
| Sprint | `1234567890@messaging.sprintpcs.com` |
| Boost | `1234567890@smsmyboostmobile.com` |
| Cricket | `1234567890@mms.cricketwireless.net` |
| Google Fi | `1234567890@msg.fi.google.com` |
| Metro PCS | `1234567890@mymetropcs.com` |

In Settings: set `phoneEmail` to your full gateway address, or set `phoneNumber` + `phoneCarrier` separately.

### Option B: Pushover (recommended for rich notifications with sound)

1. Create a free account at [pushover.net](https://pushover.net)
2. Install the Pushover app on your phone (~$5 one-time)
3. Create an Application at pushover.net → copy the **API Token**
4. Copy your **User Key** from your account page
5. In Settings: set `phoneNotificationMethod` = `pushover`, fill in `pushoverToken` and `pushoverUser`

Bill reminders play the "cash register" sound. Overdue alerts send with high priority.

### Option C: IFTTT / Zapier / Make webhook

1. Create a webhook applet in IFTTT that triggers on a web request
2. Copy the webhook URL
3. In Settings: set `phoneNotificationMethod` = `webhook`, set `webhookUrl`

The script POSTs JSON: `{ "value1": "FinFlow", "value2": "<message>", "value3": "<date>" }`

---

## Using the Dashboard

Click **💳 FinFlow → 📊 Dashboard** to open the sidebar with:
- Live KPIs (income, expenses, surplus, total debt)
- Upcoming bills for the next 30 days
- Savings goal progress bars
- Quick action buttons (send reminders, monthly report, test notification)

---

## Deploying with clasp (optional, for developers)

If you want to manage the script from the command line:

```bash
npm install -g @google/clasp
clasp login
clasp create --type sheets --title "FinFlow"
# Copy the scriptId from .clasp.json into .clasp.json here
clasp push
```

---

## Troubleshooting

| Problem | Solution |
|---|---|
| Menu doesn't appear | Refresh the sheet; re-run authorization |
| "You do not have permission" error | Re-authorize: Apps Script → Run → Authorize |
| Emails not sending | Check Gmail daily send quota (500/day for regular Gmail) |
| SMS not arriving | Verify carrier gateway; some carriers block email-to-SMS on newer plans |
| Triggers not firing | Check Apps Script → Triggers panel; time is in your account's timezone |
| "No debts found" in payoff plan | Make sure the Debts sheet has data and "Included" column is `TRUE` |
