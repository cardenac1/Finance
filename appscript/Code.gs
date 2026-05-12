// FinFlow Bill Reminder - Google Apps Script
// Backend: Google Sheets | Notifications: Email + SMS + Pushover + Webhook

// ==================== CONSTANTS ====================

var SHEETS = {
  DEBTS: 'Debts',
  INCOME: 'Income',
  EXPENSES: 'Expenses',
  GOALS: 'Goals',
  SETTINGS: 'Settings',
  LOG: 'Reminder Log'
};

var DEBT_HEADERS    = ['Name', 'Type', 'Balance ($)', 'APR (%)', 'Min Payment ($)', 'Due Day', 'Credit Limit ($)', 'Notes', 'Color', 'Included'];
var INCOME_HEADERS  = ['Name', 'Amount ($)', 'Frequency'];
var EXPENSE_HEADERS = ['Name', 'Amount ($)', 'Category'];
var GOAL_HEADERS    = ['Name', 'Target ($)', 'Saved ($)', 'Monthly Contribution ($)', 'Emoji', 'Notes'];
var LOG_HEADERS     = ['Timestamp', 'Type', 'Recipient', 'Bill / Subject', 'Status', 'Message'];

// ==================== MENU ====================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💳 FinFlow')
    .addItem('🔧 Setup Sheets', 'setupSpreadsheet')
    .addSeparator()
    .addItem('📊 Dashboard', 'showDashboard')
    .addItem('⚙️ Settings', 'showSettings')
    .addSeparator()
    .addItem('🔔 Send Bill Reminders Now', 'sendBillReminders')
    .addItem('📋 Send Weekly Digest Now', 'sendWeeklyDigest')
    .addItem('📈 Send Monthly Report Now', 'generateMonthlyReport')
    .addItem('⚠️ Check Budget Alerts Now', 'checkBudgetAlerts')
    .addSeparator()
    .addItem('🎯 View Payoff Plan', 'showPayoffPlan')
    .addItem('📅 View 12-Month Forecast', 'showForecast')
    .addSeparator()
    .addItem('⏰ Setup Automated Triggers', 'setupTriggers')
    .addItem('🗑️ Remove All Triggers', 'removeTriggers')
    .addToUi();
}

// ==================== SHEET SETUP ====================

function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  setupSheet(ss, SHEETS.DEBTS,    DEBT_HEADERS);
  setupSheet(ss, SHEETS.INCOME,   INCOME_HEADERS);
  setupSheet(ss, SHEETS.EXPENSES, EXPENSE_HEADERS);
  setupSheet(ss, SHEETS.GOALS,    GOAL_HEADERS);
  setupSheet(ss, SHEETS.SETTINGS, ['Setting', 'Value', 'Description']);
  setupSheet(ss, SHEETS.LOG,      LOG_HEADERS);

  setupDefaultSettings(ss);
  addSampleData(ss);
  applyFormatting(ss);

  SpreadsheetApp.getUi().alert(
    '✅ FinFlow setup complete!\n\n' +
    'Next steps:\n' +
    '1. Review the Settings sheet and fill in your email & phone info\n' +
    '2. Replace the sample data with your real bills/income/expenses\n' +
    '3. Go to FinFlow → Setup Automated Triggers to enable reminders'
  );
}

function setupSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  var hr = sheet.getRange(1, 1, 1, headers.length);
  hr.setValues([headers]);
  hr.setBackground('#1a1a2e');
  hr.setFontColor('#5c7cff');
  hr.setFontWeight('bold');
  hr.setFontSize(11);
  hr.setFrozenRows(1);
  sheet.setFrozenRows(1);
  return sheet;
}

function setupDefaultSettings(ss) {
  var sheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (sheet.getLastRow() > 1) return;

  var defaults = [
    ['email',                   Session.getActiveUser().getEmail(), 'Your Gmail address for email notifications'],
    ['phoneEmail',              '',          'Phone-to-email gateway (e.g. 6505551234@txt.att.net). Leave blank to use carrier + number below.'],
    ['phoneNumber',             '',          '10-digit mobile number (digits only, no dashes)'],
    ['phoneCarrier',            'att',       'Carrier for SMS gateway: att | tmobile | verizon | sprint | boost | cricket | googlefi | metro'],
    ['enableEmailNotifications','true',      'Send email reminders (true/false)'],
    ['enablePhoneNotifications','false',     'Send SMS/push to phone (true/false)'],
    ['phoneNotificationMethod', 'sms',       'Phone method: sms | pushover | webhook'],
    ['pushoverToken',           '',          'Pushover app API token (https://pushover.net)'],
    ['pushoverUser',            '',          'Pushover user key'],
    ['webhookUrl',              '',          'Generic webhook URL (IFTTT, Zapier, Make, etc.)'],
    ['reminderDays',            '7,3,1',     'Days before due date to send reminders (comma-separated)'],
    ['sendDueTodayAlert',       'true',      'Send alert when bill is due today (true/false)'],
    ['overdueAlertEnabled',     'true',      'Alert for bills that appear overdue (true/false)'],
    ['overdueGraceDays',        '3',         'Grace period in days before flagging as overdue'],
    ['extraPayment',            '0',         'Extra monthly amount applied toward debt payoff ($)'],
    ['payoffStrategy',          'avalanche', 'Payoff strategy: avalanche (highest APR first) | snowball (lowest balance first)'],
    ['budgetAlertThreshold',    '90',        'Alert when total obligations exceed X% of income'],
    ['weeklyDigestEnabled',     'true',      'Send weekly financial summary (true/false)'],
    ['monthlyReportEnabled',    'true',      'Send monthly report (true/false)'],
    ['monthlyReportDay',        '1',         'Day of month to send monthly report (1-28)'],
    ['forecastMonths',          '12',        'Default forecast horizon in months (12, 24, or 36)'],
  ];

  sheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
}

function addSampleData(ss) {
  var debts = ss.getSheetByName(SHEETS.DEBTS);
  if (debts.getLastRow() <= 1) {
    debts.getRange(2, 1, 5, 10).setValues([
      ['AMEX Gold',       'Credit Card',   3200,  24.99,  96,  15, 10000, 'Travel rewards card',   '#5c7cff', true],
      ['Chase Sapphire',  'Credit Card',   1850,  19.99,  55,  22,  8000, 'Cash back card',        '#a78bfa', true],
      ['Personal Loan',   'Personal Loan', 8500,  12.50, 280,   5,     0, 'Home improvement',      '#22d47e', true],
      ['Student Loans',   'Student Loan', 15000,   5.80, 165,   1,     0, 'Federal student loans', '#ff9f43', true],
      ['Tesla Auto Loan', 'Auto Loan',    22000,   6.90, 445,  18,     0, 'Tesla Model 3',         '#ff5b6b', true],
    ]);
  }

  var income = ss.getSheetByName(SHEETS.INCOME);
  if (income.getLastRow() <= 1) {
    income.getRange(2, 1, 2, 3).setValues([
      ['Primary Salary', 7500, 'monthly'],
      ['Side Income',     500, 'monthly'],
    ]);
  }

  var expenses = ss.getSheetByName(SHEETS.EXPENSES);
  if (expenses.getLastRow() <= 1) {
    expenses.getRange(2, 1, 4, 3).setValues([
      ['Rent',          1800, 'housing'],
      ['Utilities',      150, 'utilities'],
      ['Subscriptions',   45, 'subscriptions'],
      ['Insurance',      120, 'health'],
    ]);
  }

  var goals = ss.getSheetByName(SHEETS.GOALS);
  if (goals.getLastRow() <= 1) {
    goals.getRange(2, 1, 2, 6).setValues([
      ['Hawaii Vacation', 5000,  800, 200, '✈️', 'Trip planned for December'],
      ['Emergency Fund', 10000, 3200, 300, '🛡️', '3-6 months of expenses'],
    ]);
  }
}

function applyFormatting(ss) {
  var currency = '$#,##0.00';
  var pct      = '0.00"%"';

  function fmtCols(sheetName, currCols, pctCols) {
    var s = ss.getSheetByName(sheetName);
    var rows = Math.max(s.getLastRow() - 1, 1);
    (currCols || []).forEach(function(c) {
      s.getRange(2, c, rows).setNumberFormat(currency);
    });
    (pctCols || []).forEach(function(c) {
      s.getRange(2, c, rows).setNumberFormat(pct);
    });
  }

  fmtCols(SHEETS.DEBTS,    [3, 5, 7], [4]);
  fmtCols(SHEETS.INCOME,   [2]);
  fmtCols(SHEETS.EXPENSES, [2]);
  fmtCols(SHEETS.GOALS,    [2, 3, 4]);

  // Column widths – Debts
  var d = ss.getSheetByName(SHEETS.DEBTS);
  [180, 120, 100, 80, 110, 80, 110, 200, 80, 80].forEach(function(w, i) {
    d.setColumnWidth(i + 1, w);
  });
}

// ==================== SETTINGS ACCESS ====================

function getSetting(key) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) return (data[i][1] || '').toString();
  }
  return '';
}

function setSetting(key, value) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value, '']);
}

function getAllSettings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var out  = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) out[data[i][0]] = data[i][1];
  }
  return out;
}

function saveSettingFromUI(key, value) {
  setSetting(key, value);
  return { success: true };
}

function saveAllSettingsFromUI(obj) {
  Object.keys(obj).forEach(function(k) { setSetting(k, obj[k]); });
  return { success: true };
}

// ==================== DATA ACCESS ====================

function getDebts() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.DEBTS);
  if (!s || s.getLastRow() <= 1) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, DEBT_HEADERS.length).getValues()
    .filter(function(r) { return r[0]; })
    .map(function(r) {
      return {
        name:       r[0],
        type:       r[1],
        balance:    parseFloat(r[2]) || 0,
        apr:        parseFloat(r[3]) || 0,
        minPayment: parseFloat(r[4]) || 0,
        dueDay:     parseInt(r[5])   || 1,
        creditLimit:parseFloat(r[6]) || 0,
        notes:      r[7] || '',
        color:      r[8] || '#5c7cff',
        included:   r[9] !== false && r[9] !== 'false' && r[9] !== 0
      };
    });
}

function getIncome() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.INCOME);
  if (!s || s.getLastRow() <= 1) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, INCOME_HEADERS.length).getValues()
    .filter(function(r) { return r[0]; })
    .map(function(r) {
      return { name: r[0], amount: parseFloat(r[1]) || 0, frequency: r[2] || 'monthly' };
    });
}

function getExpenses() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EXPENSES);
  if (!s || s.getLastRow() <= 1) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, EXPENSE_HEADERS.length).getValues()
    .filter(function(r) { return r[0]; })
    .map(function(r) {
      return { name: r[0], amount: parseFloat(r[1]) || 0, category: r[2] || 'other' };
    });
}

function getGoals() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.GOALS);
  if (!s || s.getLastRow() <= 1) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, GOAL_HEADERS.length).getValues()
    .filter(function(r) { return r[0]; })
    .map(function(r) {
      return {
        name:               r[0],
        target:             parseFloat(r[1]) || 0,
        current:            parseFloat(r[2]) || 0,
        monthlyContribution:parseFloat(r[3]) || 0,
        emoji:              r[4] || '🎯',
        notes:              r[5] || ''
      };
    });
}

// ==================== FINANCIAL CALCULATIONS ====================

function getMonthlyIncome() {
  return getIncome().reduce(function(t, s) {
    var f = (s.frequency || '').toLowerCase();
    if (f === 'bi-weekly') return t + s.amount * 26 / 12;
    if (f === 'weekly')    return t + s.amount * 52 / 12;
    if (f === 'annual')    return t + s.amount / 12;
    return t + s.amount; // monthly default
  }, 0);
}

function getMonthlyExpenses()     { return getExpenses().reduce(function(t, e) { return t + e.amount; }, 0); }
function getTotalDebt()           { return getDebts().filter(function(d){return d.included;}).reduce(function(t,d){return t+d.balance;},0); }
function getMonthlyDebtPayments() { return getDebts().filter(function(d){return d.included;}).reduce(function(t,d){return t+d.minPayment;},0); }
function getMonthlyInterestCost() {
  return getDebts().filter(function(d){return d.included;}).reduce(function(t,d){
    return t + d.balance * (d.apr / 100) / 12;
  }, 0);
}
function getNetSurplus() { return getMonthlyIncome() - getMonthlyExpenses() - getMonthlyDebtPayments(); }

// ==================== PAYOFF CALCULATIONS ====================

function calculateAvalanche(extra) {
  var debts = getDebts().filter(function(d){return d.included && d.balance > 0;});
  debts.sort(function(a,b){return b.apr - a.apr;});
  return _runPayoff(debts, parseFloat(extra) || parseFloat(getSetting('extraPayment')) || 0);
}

function calculateSnowball(extra) {
  var debts = getDebts().filter(function(d){return d.included && d.balance > 0;});
  debts.sort(function(a,b){return a.balance - b.balance;});
  return _runPayoff(debts, parseFloat(extra) || parseFloat(getSetting('extraPayment')) || 0);
}

function _runPayoff(sorted, extraPayment) {
  if (!sorted.length) return null;
  var balances = sorted.map(function(d){return d.balance;});
  var results  = sorted.map(function(d){return {name:d.name,totalInterest:0,payoffMonth:0};});
  var month = 0;

  while (balances.some(function(b){return b > 0.01;}) && month < 360) {
    month++;
    var extra = extraPayment;
    var firstActive = balances.findIndex(function(b){return b > 0;});

    for (var i = 0; i < sorted.length; i++) {
      if (balances[i] <= 0) continue;
      var interest = balances[i] * (sorted[i].apr / 100) / 12;
      results[i].totalInterest += interest;
      balances[i] += interest;
      var payment = sorted[i].minPayment + (i === firstActive ? extra : 0);
      payment = Math.min(payment, balances[i]);
      balances[i] = Math.max(0, balances[i] - payment);
      if (balances[i] < 0.01 && !results[i].payoffMonth) results[i].payoffMonth = month;
    }
  }

  var payoffDate = new Date();
  payoffDate.setMonth(payoffDate.getMonth() + month);

  return {
    results:       results,
    totalMonths:   month,
    totalInterest: results.reduce(function(s,r){return s+r.totalInterest;}, 0),
    payoffDate:    payoffDate.toLocaleDateString('en-US', {month:'long', year:'numeric'})
  };
}

// ==================== FORECAST ====================

function generateForecast(months) {
  months = parseInt(months) || parseInt(getSetting('forecastMonths')) || 12;
  var debts    = getDebts().filter(function(d){return d.included && d.balance > 0;});
  var income   = getMonthlyIncome();
  var expenses = getMonthlyExpenses();
  var extra    = parseFloat(getSetting('extraPayment')) || 0;
  var strategy = getSetting('payoffStrategy') || 'avalanche';
  var goals    = getGoals();
  var goalContribs = goals.reduce(function(s,g){return s+g.monthlyContribution;}, 0);

  var sorted = debts.slice().sort(function(a,b){
    return strategy === 'snowball' ? a.balance - b.balance : b.apr - a.apr;
  });
  var balances = sorted.map(function(d){return d.balance;});
  var cumSavings = goals.reduce(function(s,g){return s+g.current;}, 0);
  var rows = [];

  for (var m = 0; m < months; m++) {
    var d = new Date();
    d.setMonth(d.getMonth() + m);
    var label = d.toLocaleDateString('en-US', {month:'short', year:'numeric'});

    var totalDebtPmt = 0;
    var ex = extra;
    var firstActive = balances.findIndex(function(b){return b > 0;});

    for (var i = 0; i < sorted.length; i++) {
      if (balances[i] <= 0) continue;
      var int = balances[i] * (sorted[i].apr / 100) / 12;
      balances[i] += int;
      var pmt = sorted[i].minPayment + (i === firstActive ? ex : 0);
      pmt = Math.min(pmt, balances[i]);
      balances[i] = Math.max(0, balances[i] - pmt);
      totalDebtPmt += pmt;
    }

    var surplus = income - expenses - totalDebtPmt - goalContribs;
    cumSavings += goalContribs + Math.max(0, surplus);

    rows.push({
      month:           label,
      income:          income,
      expenses:        expenses,
      debtPayments:    totalDebtPmt,
      goalContribs:    goalContribs,
      surplus:         surplus,
      cumSavings:      cumSavings,
      totalDebt:       balances.reduce(function(s,b){return s+b;}, 0),
      debtByAccount:   sorted.map(function(d,i){return {name:d.name,balance:balances[i]};})
    });
  }

  return rows;
}

// ==================== BILL REMINDERS ====================

function sendBillReminders() {
  var debts          = getDebts();
  var reminderDays   = (getSetting('reminderDays') || '7,3,1').split(',').map(function(x){return parseInt(x.trim());});
  var enableEmail    = getSetting('enableEmailNotifications') !== 'false';
  var enablePhone    = getSetting('enablePhoneNotifications') === 'true';
  var overdueEnabled = getSetting('overdueAlertEnabled') !== 'false';
  var graceDays      = parseInt(getSetting('overdueGraceDays')) || 3;
  var dueTodayAlert  = getSetting('sendDueTodayAlert') !== 'false';

  var today       = new Date();
  var currentDay  = today.getDate();
  var daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();

  var upcoming = [];
  var overdue  = [];

  debts.forEach(function(debt) {
    if (!debt.dueDay || !debt.included) return;
    var due = parseInt(debt.dueDay);

    if (due > currentDay) {
      var daysUntil = due - currentDay;
      if (daysUntil === 0 && dueTodayAlert) {
        upcoming.push(Object.assign({}, debt, {daysUntilDue: 0}));
      } else if (reminderDays.indexOf(daysUntil) !== -1) {
        upcoming.push(Object.assign({}, debt, {daysUntilDue: daysUntil}));
      }
    } else if (due === currentDay && dueTodayAlert) {
      upcoming.push(Object.assign({}, debt, {daysUntilDue: 0}));
    } else if (overdueEnabled && (currentDay - due) > graceDays) {
      overdue.push(Object.assign({}, debt, {daysOverdue: currentDay - due}));
    }
  });

  upcoming.forEach(function(bill) {
    if (enableEmail) sendEmailReminder(bill, bill.daysUntilDue);
    if (enablePhone) sendPhoneNotification(_phoneMsg(bill, bill.daysUntilDue));
  });

  if (overdue.length > 0) {
    if (enableEmail) sendOverdueAlert(overdue);
    if (enablePhone) sendPhoneNotification('⚠️ OVERDUE: ' + overdue.map(function(b){return b.name;}).join(', '));
  }

  logEntry('BATCH_CHECK', getSetting('email'),
    upcoming.length + ' reminders, ' + overdue.length + ' overdue',
    'OK',
    'Bill reminder check completed');
}

function _phoneMsg(bill, daysUntilDue) {
  var when = daysUntilDue === 0 ? 'DUE TODAY' : 'Due in ' + daysUntilDue + 'd';
  return 'FinFlow: ' + bill.name + ' ' + when + ' $' + bill.minPayment.toFixed(2);
}

// ==================== EMAIL REMINDER ====================

function sendEmailReminder(bill, daysUntilDue) {
  var email = getSetting('email');
  if (!email) return;

  var urgencyLabel = daysUntilDue === 0 ? '🚨 DUE TODAY'
                   : daysUntilDue === 1 ? '⚠️ Due Tomorrow'
                   : '📅 Due in ' + daysUntilDue + ' days';
  var urgencyColor = daysUntilDue === 0 ? '#ff5b6b'
                   : daysUntilDue <= 3  ? '#ff9f43'
                   : '#5c7cff';
  var subject = urgencyLabel + ': ' + bill.name + ' – $' + bill.minPayment.toFixed(2);

  var utilizationRow = bill.creditLimit > 0
    ? '<tr><td style="padding:9px 12px;font-weight:bold;background:#f8f9fa;">Credit Utilization</td>' +
      '<td style="padding:9px 12px;">' +
      (bill.balance / bill.creditLimit * 100).toFixed(1) + '% ' +
      '($' + bill.balance.toFixed(2) + ' / $' + bill.creditLimit.toFixed(2) + ')</td></tr>'
    : '';

  var notesRow = bill.notes
    ? '<tr style="background:#f8f9fa;"><td style="padding:9px 12px;font-weight:bold;">Notes</td>' +
      '<td style="padding:9px 12px;">' + bill.notes + '</td></tr>'
    : '';

  var html = '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0f2f5;">' +
    '<div style="font-family:\'Helvetica Neue\',Arial,sans-serif;max-width:600px;margin:30px auto;background:#f0f2f5;padding:20px;">' +
    '<div style="background:linear-gradient(135deg,#1a1a2e,#16213e);padding:28px;border-radius:12px 12px 0 0;">' +
    '<div style="color:#5c7cff;font-size:13px;font-weight:600;letter-spacing:1px;margin-bottom:6px;">FINFLOW</div>' +
    '<h1 style="color:white;margin:0;font-size:22px;">💳 Bill Payment Reminder</h1>' +
    '</div>' +
    '<div style="background:white;padding:28px;border-radius:0 0 12px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08);">' +
    '<div style="background:' + urgencyColor + '15;border-left:4px solid ' + urgencyColor + ';padding:14px 18px;border-radius:6px;margin-bottom:24px;">' +
    '<span style="color:' + urgencyColor + ';font-size:18px;font-weight:bold;">' + urgencyLabel + '</span>' +
    '</div>' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px;">' +
    '<tr><td style="padding:9px 12px;font-weight:bold;background:#f8f9fa;border-radius:6px 0 0 0;">Account</td>' +
    '<td style="padding:9px 12px;">' + bill.name + '</td></tr>' +
    '<tr><td style="padding:9px 12px;font-weight:bold;">Type</td>' +
    '<td style="padding:9px 12px;">' + bill.type + '</td></tr>' +
    '<tr style="background:#f8f9fa;"><td style="padding:9px 12px;font-weight:bold;">Amount Due</td>' +
    '<td style="padding:9px 12px;color:#ff5b6b;font-size:20px;font-weight:bold;">$' + bill.minPayment.toFixed(2) + '</td></tr>' +
    '<tr><td style="padding:9px 12px;font-weight:bold;">Due Day</td>' +
    '<td style="padding:9px 12px;">Day ' + bill.dueDay + ' of each month</td></tr>' +
    '<tr style="background:#f8f9fa;"><td style="padding:9px 12px;font-weight:bold;">Current Balance</td>' +
    '<td style="padding:9px 12px;">$' + bill.balance.toFixed(2) + '</td></tr>' +
    '<tr><td style="padding:9px 12px;font-weight:bold;">APR</td>' +
    '<td style="padding:9px 12px;">' + bill.apr + '%</td></tr>' +
    utilizationRow + notesRow +
    '</table>' +
    '<p style="color:#aaa;font-size:11px;margin-top:24px;border-top:1px solid #eee;padding-top:16px;text-align:center;">' +
    'Sent by FinFlow Bill Reminder · Powered by Google Apps Script<br>' +
    'To manage reminders: open your FinFlow Google Sheet → FinFlow menu' +
    '</p></div></div></body></html>';

  try {
    GmailApp.sendEmail(email, subject,
      urgencyLabel + ': ' + bill.name + ' – $' + bill.minPayment.toFixed(2) + ' due on day ' + bill.dueDay,
      { htmlBody: html, name: 'FinFlow Bill Reminder' });
    logEntry('EMAIL', email, bill.name, 'Sent', subject);
  } catch (e) {
    logEntry('EMAIL', email, bill.name, 'FAILED', e.toString());
  }
}

// ==================== OVERDUE ALERT ====================

function sendOverdueAlert(overdueBills) {
  var email = getSetting('email');
  if (!email) return;

  var subject = '🚨 FinFlow: ' + overdueBills.length + ' bill(s) may be overdue!';

  var rows = overdueBills.map(function(b) {
    return '<tr>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #eee;">' + b.name + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #eee;">' + b.type + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #eee;color:#ff5b6b;font-weight:bold;">$' + b.minPayment.toFixed(2) + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #eee;">Day ' + b.dueDay + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #eee;color:#ff5b6b;">' + b.daysOverdue + ' days ago</td>' +
      '</tr>';
  }).join('');

  var html = '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0f2f5;">' +
    '<div style="font-family:Arial,sans-serif;max-width:620px;margin:30px auto;">' +
    '<div style="background:#ff5b6b;padding:24px;border-radius:12px 12px 0 0;">' +
    '<h1 style="color:white;margin:0;font-size:20px;">🚨 Overdue Bills Alert</h1>' +
    '</div>' +
    '<div style="background:white;padding:24px;border-radius:0 0 12px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08);">' +
    '<p style="color:#555;">The following bills may be overdue. Please pay immediately to avoid late fees and credit score impact.</p>' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
    '<thead><tr style="background:#1a1a2e;color:white;">' +
    '<th style="padding:10px 12px;text-align:left;">Account</th>' +
    '<th style="padding:10px 12px;text-align:left;">Type</th>' +
    '<th style="padding:10px 12px;text-align:left;">Amount</th>' +
    '<th style="padding:10px 12px;text-align:left;">Due Day</th>' +
    '<th style="padding:10px 12px;text-align:left;">Overdue</th>' +
    '</tr></thead><tbody>' + rows + '</tbody></table>' +
    '<p style="color:#aaa;font-size:11px;margin-top:20px;text-align:center;">FinFlow Alert · ' + new Date().toLocaleString() + '</p>' +
    '</div></div></body></html>';

  try {
    GmailApp.sendEmail(email, subject, 'Overdue: ' + overdueBills.map(function(b){return b.name;}).join(', '),
      { htmlBody: html, name: 'FinFlow Alert' });
    logEntry('EMAIL', email, 'Overdue Alert', 'Sent', subject);
  } catch (e) {
    logEntry('EMAIL', email, 'Overdue Alert', 'FAILED', e.toString());
  }
}

// ==================== PHONE NOTIFICATIONS ====================

function sendPhoneNotification(message) {
  var method = getSetting('phoneNotificationMethod') || 'sms';
  if (method === 'pushover') {
    _sendPushover(message);
  } else if (method === 'webhook') {
    _sendWebhook(message);
  } else {
    _sendSmsEmail(message);
  }
}

var CARRIER_GATEWAYS = {
  att:       '@txt.att.net',
  tmobile:   '@tmomail.net',
  'T-Mobile':'@tmomail.net',
  verizon:   '@vtext.com',
  sprint:    '@messaging.sprintpcs.com',
  boost:     '@smsmyboostmobile.com',
  cricket:   '@mms.cricketwireless.net',
  uscellular:'@email.uscc.net',
  googlefi:  '@msg.fi.google.com',
  metro:     '@mymetropcs.com'
};

function _sendSmsEmail(message) {
  var to = getSetting('phoneEmail');
  if (!to) {
    var num     = getSetting('phoneNumber').replace(/\D/g, '');
    var carrier = (getSetting('phoneCarrier') || '').toLowerCase();
    var gw      = CARRIER_GATEWAYS[carrier];
    if (!num || !gw) { logEntry('SMS', '—', '—', 'SKIPPED', 'No phone email or carrier configured'); return; }
    to = num + gw;
  }
  try {
    GmailApp.sendEmail(to, 'FinFlow', message, { name: 'FinFlow' });
    logEntry('SMS', to, 'Phone', 'Sent', message);
  } catch (e) {
    logEntry('SMS', to, 'Phone', 'FAILED', e.toString());
  }
}

function _sendPushover(message) {
  var token = getSetting('pushoverToken');
  var user  = getSetting('pushoverUser');
  if (!token || !user) { logEntry('PUSH', 'Pushover', '—', 'SKIPPED', 'Pushover credentials missing'); return; }

  try {
    UrlFetchApp.fetch('https://api.pushover.net/1/messages.json', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        token:    token,
        user:     user,
        title:    'FinFlow',
        message:  message,
        sound:    'cashregister',
        priority: (message.indexOf('OVERDUE') !== -1 || message.indexOf('TODAY') !== -1) ? 1 : 0
      })
    });
    logEntry('PUSH', 'Pushover', 'Phone', 'Sent', message);
  } catch (e) {
    logEntry('PUSH', 'Pushover', 'Phone', 'FAILED', e.toString());
  }
}

function _sendWebhook(message) {
  var url = getSetting('webhookUrl');
  if (!url) { logEntry('WEBHOOK', '—', '—', 'SKIPPED', 'No webhook URL configured'); return; }
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ value1: 'FinFlow', value2: message, value3: new Date().toLocaleDateString() })
    });
    logEntry('WEBHOOK', url, 'Phone', 'Sent', message);
  } catch (e) {
    logEntry('WEBHOOK', url, 'Phone', 'FAILED', e.toString());
  }
}

// ==================== WEEKLY DIGEST ====================

function sendWeeklyDigest() {
  if (getSetting('weeklyDigestEnabled') === 'false') return;
  var email = getSetting('email');
  if (!email) return;

  var debts    = getDebts().filter(function(d){return d.included;});
  var income   = getMonthlyIncome();
  var expenses = getMonthlyExpenses();
  var surplus  = getNetSurplus();
  var debt     = getTotalDebt();
  var today    = new Date();
  var curDay   = today.getDate();
  var dim      = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();

  var billsThisWeek = debts.filter(function(d) {
    var until = d.dueDay >= curDay ? d.dueDay - curDay : (dim - curDay) + d.dueDay;
    return until <= 7;
  }).map(function(d) {
    var until = d.dueDay >= curDay ? d.dueDay - curDay : (dim - curDay) + d.dueDay;
    return { debt: d, until: until };
  }).sort(function(a,b){return a.until - b.until;});

  var billRows = billsThisWeek.length > 0
    ? billsThisWeek.map(function(x) {
        var c = x.until === 0 ? '#ff5b6b' : x.until <= 2 ? '#ff9f43' : '#333';
        var w = x.until === 0 ? 'TODAY' : x.until === 1 ? 'Tomorrow' : 'In ' + x.until + ' days';
        return '<tr>' +
          '<td style="padding:8px 12px;border-bottom:1px solid #f0f0f0;">' + x.debt.name + '</td>' +
          '<td style="padding:8px 12px;border-bottom:1px solid #f0f0f0;color:' + c + ';font-weight:bold;">' + w + '</td>' +
          '<td style="padding:8px 12px;border-bottom:1px solid #f0f0f0;text-align:right;font-weight:bold;">$' + x.debt.minPayment.toFixed(2) + '</td>' +
          '</tr>';
      }).join('')
    : '<tr><td colspan="3" style="padding:20px;text-align:center;color:#aaa;">🎉 No bills due this week!</td></tr>';

  var surplusColor = surplus >= 0 ? '#22d47e' : '#ff5b6b';
  var subject = '📋 FinFlow Weekly Digest – ' + today.toLocaleDateString('en-US', {month:'long', day:'numeric'});

  var html = '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0f2f5;">' +
    '<div style="font-family:Arial,sans-serif;max-width:580px;margin:30px auto;">' +
    '<div style="background:linear-gradient(135deg,#1a1a2e,#16213e);padding:24px;border-radius:12px 12px 0 0;">' +
    '<div style="color:#a78bfa;font-size:12px;font-weight:600;letter-spacing:1px;margin-bottom:6px;">FINFLOW</div>' +
    '<h1 style="color:white;margin:0;font-size:20px;">📋 Weekly Digest</h1>' +
    '<p style="color:#a0a0b0;margin:6px 0 0;">' + today.toLocaleDateString('en-US',{weekday:'long',month:'long',day:'numeric',year:'numeric'}) + '</p>' +
    '</div>' +
    '<div style="background:white;padding:24px;border-radius:0 0 12px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08);">' +
    // KPI row
    '<div style="display:flex;gap:12px;margin-bottom:22px;flex-wrap:wrap;">' +
    _kpiBox('Net Surplus', '$' + Math.abs(surplus).toFixed(0), surplusColor, '#f0f4ff') +
    _kpiBox('Total Debt',  '$' + debt.toFixed(0),    '#ff9f43', '#fff7f0') +
    _kpiBox('Income',      '$' + income.toFixed(0),  '#22d47e', '#f0fff4') +
    _kpiBox('Expenses',    '$' + expenses.toFixed(0),'#ff5b6b', '#fff0f1') +
    '</div>' +
    '<h3 style="color:#1a1a2e;font-size:15px;margin:0 0 12px;border-bottom:2px solid #a78bfa;padding-bottom:8px;">Bills Due This Week</h3>' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
    '<thead><tr style="background:#f8f9fa;">' +
    '<th style="padding:8px 12px;text-align:left;font-weight:600;">Bill</th>' +
    '<th style="padding:8px 12px;text-align:left;font-weight:600;">When</th>' +
    '<th style="padding:8px 12px;text-align:right;font-weight:600;">Amount</th>' +
    '</tr></thead><tbody>' + billRows + '</tbody></table>' +
    '<p style="color:#aaa;font-size:11px;margin-top:20px;text-align:center;">FinFlow Weekly Digest · ' + new Date().toLocaleString() + '</p>' +
    '</div></div></body></html>';

  try {
    GmailApp.sendEmail(email, subject,
      'FinFlow Weekly Digest\nSurplus: $' + surplus.toFixed(2) + '\nDebt: $' + debt.toFixed(2) + '\nBills this week: ' + billsThisWeek.length,
      { htmlBody: html, name: 'FinFlow' });
    logEntry('DIGEST', email, 'Weekly', 'Sent', subject);
  } catch (e) {
    logEntry('DIGEST', email, 'Weekly', 'FAILED', e.toString());
  }
}

// ==================== MONTHLY REPORT ====================

function generateMonthlyReport() {
  if (getSetting('monthlyReportEnabled') === 'false') return;
  var email = getSetting('email');
  if (!email) return;

  var debts        = getDebts().filter(function(d){return d.included;});
  var income       = getMonthlyIncome();
  var expenses     = getMonthlyExpenses();
  var debtPayments = getMonthlyDebtPayments();
  var interest     = getMonthlyInterestCost();
  var totalDebt    = getTotalDebt();
  var surplus      = getNetSurplus();
  var goals        = getGoals();
  var strategy     = getSetting('payoffStrategy') || 'avalanche';
  var plan         = strategy === 'avalanche' ? calculateAvalanche() : calculateSnowball();

  var month   = new Date().toLocaleDateString('en-US', {month:'long', year:'numeric'});
  var subject = '📊 FinFlow Monthly Report – ' + month;

  var debtRows = debts.map(function(d) {
    var util = d.creditLimit > 0 ? (d.balance / d.creditLimit * 100).toFixed(1) + '%' : '—';
    var uc   = d.creditLimit > 0 && (d.balance / d.creditLimit) > 0.8 ? '#ff5b6b' : '#22d47e';
    return '<tr>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;">' + d.name + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;">' + d.type + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:right;">$' + d.balance.toFixed(2) + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:right;">' + d.apr + '%</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:right;color:#ff5b6b;font-weight:bold;">$' + d.minPayment.toFixed(2) + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:right;color:' + uc + ';">' + util + '</td>' +
      '</tr>';
  }).join('');

  var goalRows = goals.map(function(g) {
    var pct = g.target > 0 ? (g.current / g.target * 100).toFixed(1) : 0;
    var ml  = g.monthlyContribution > 0 ? Math.ceil((g.target - g.current) / g.monthlyContribution) : '∞';
    return '<tr>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;">' + g.emoji + ' ' + g.name + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:right;">$' + g.current.toFixed(2) + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:right;">$' + g.target.toFixed(2) + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:right;color:#22d47e;font-weight:bold;">' + pct + '%</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:right;">' + ml + ' mo</td>' +
      '</tr>';
  }).join('');

  var planBox = plan
    ? '<div style="background:#e8f5e9;border-left:4px solid #22d47e;padding:14px 18px;border-radius:6px;margin:16px 0;font-size:13px;">' +
      '🎯 <strong>Debt-Free Estimate:</strong> <span style="color:#22d47e;font-weight:bold;">' + plan.payoffDate + '</span>' +
      ' &nbsp;|&nbsp; <strong>Total Interest:</strong> <span style="color:#ff9f43;">$' + plan.totalInterest.toFixed(2) + '</span>' +
      ' &nbsp;|&nbsp; <strong>Strategy:</strong> ' + strategy.charAt(0).toUpperCase() + strategy.slice(1) +
      '</div>'
    : '';

  var surplusColor = surplus >= 0 ? '#22d47e' : '#ff5b6b';
  var goalContribs = goals.reduce(function(s,g){return s+g.monthlyContribution;}, 0);

  var html = '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0f2f5;">' +
    '<div style="font-family:Arial,sans-serif;max-width:700px;margin:30px auto;">' +
    '<div style="background:linear-gradient(135deg,#1a1a2e,#16213e);padding:28px;border-radius:12px 12px 0 0;">' +
    '<div style="color:#5c7cff;font-size:12px;font-weight:600;letter-spacing:1px;margin-bottom:6px;">FINFLOW</div>' +
    '<h1 style="color:white;margin:0;font-size:22px;">📊 Monthly Financial Report</h1>' +
    '<p style="color:#a0a0b0;margin:6px 0 0;">' + month + '</p>' +
    '</div>' +
    '<div style="background:white;padding:28px;border-radius:0 0 12px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08);">' +
    // KPI grid
    '<h3 style="color:#1a1a2e;border-bottom:2px solid #5c7cff;padding-bottom:8px;margin-top:0;">Financial Summary</h3>' +
    '<div style="display:grid;grid-template-columns:repeat(2,1fr);gap:14px;margin-bottom:22px;">' +
    _kpiBox('Monthly Income',     '$' + income.toFixed(2),       '#22d47e', '#f0fff4') +
    _kpiBox('Monthly Expenses',   '$' + expenses.toFixed(2),     '#ff5b6b', '#fff0f1') +
    _kpiBox('Net Surplus',        (surplus >= 0 ? '+' : '') + '$' + surplus.toFixed(2), surplusColor, '#f0f4ff') +
    _kpiBox('Total Debt',         '$' + totalDebt.toFixed(2),    '#ff9f43', '#fff7f0') +
    '</div>' +
    '<div style="background:#fff8e1;padding:12px 16px;border-radius:8px;font-size:13px;margin-bottom:18px;">' +
    '<strong>Monthly Interest Cost:</strong> <span style="color:#ff9f43;">$' + interest.toFixed(2) + '</span>' +
    ' &nbsp;·&nbsp; <strong>Min Debt Payments:</strong> <span style="color:#ff5b6b;">$' + debtPayments.toFixed(2) + '</span>' +
    ' &nbsp;·&nbsp; <strong>Goal Contributions:</strong> <span style="color:#22d47e;">$' + goalContribs.toFixed(2) + '</span>' +
    '</div>' +
    planBox +
    // Budget breakdown
    '<h3 style="color:#1a1a2e;border-bottom:2px solid #a78bfa;padding-bottom:8px;">Budget Breakdown</h3>' +
    _budgetBar('Living Expenses', expenses, income, '#ff5b6b') +
    _budgetBar('Debt Payments',   debtPayments, income, '#ff9f43') +
    _budgetBar('Savings / Goals', goalContribs, income, '#22d47e') +
    _budgetBar('Surplus',         Math.max(0, surplus - goalContribs), income, '#5c7cff') +
    // Accounts table
    '<h3 style="color:#1a1a2e;border-bottom:2px solid #5c7cff;padding-bottom:8px;margin-top:24px;">Account Details</h3>' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
    '<thead><tr style="background:#1a1a2e;color:white;">' +
    '<th style="padding:10px 12px;text-align:left;">Account</th>' +
    '<th style="padding:10px 12px;text-align:left;">Type</th>' +
    '<th style="padding:10px 12px;text-align:right;">Balance</th>' +
    '<th style="padding:10px 12px;text-align:right;">APR</th>' +
    '<th style="padding:10px 12px;text-align:right;">Min Payment</th>' +
    '<th style="padding:10px 12px;text-align:right;">Utilization</th>' +
    '</tr></thead><tbody>' + debtRows + '</tbody>' +
    '<tfoot><tr style="background:#f0f4ff;font-weight:bold;">' +
    '<td colspan="2" style="padding:10px 12px;">TOTALS</td>' +
    '<td style="padding:10px 12px;text-align:right;">$' + totalDebt.toFixed(2) + '</td>' +
    '<td style="padding:10px 12px;text-align:right;">—</td>' +
    '<td style="padding:10px 12px;text-align:right;">$' + debtPayments.toFixed(2) + '</td>' +
    '<td style="padding:10px 12px;text-align:right;">—</td>' +
    '</tr></tfoot></table>' +
    // Goals table
    (goals.length > 0
      ? '<h3 style="color:#1a1a2e;border-bottom:2px solid #22d47e;padding-bottom:8px;margin-top:24px;">Savings Goals</h3>' +
        '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
        '<thead><tr style="background:#1a1a2e;color:white;">' +
        '<th style="padding:10px 12px;text-align:left;">Goal</th>' +
        '<th style="padding:10px 12px;text-align:right;">Saved</th>' +
        '<th style="padding:10px 12px;text-align:right;">Target</th>' +
        '<th style="padding:10px 12px;text-align:right;">Progress</th>' +
        '<th style="padding:10px 12px;text-align:right;">Time Left</th>' +
        '</tr></thead><tbody>' + goalRows + '</tbody></table>'
      : '') +
    '<p style="color:#aaa;font-size:11px;margin-top:28px;border-top:1px solid #eee;padding-top:16px;text-align:center;">' +
    'Automatically generated by FinFlow · ' + new Date().toLocaleString() + '</p>' +
    '</div></div></body></html>';

  try {
    GmailApp.sendEmail(email, subject,
      'FinFlow Monthly Report\nIncome: $' + income.toFixed(2) +
      '\nExpenses: $' + expenses.toFixed(2) +
      '\nSurplus: $' + surplus.toFixed(2) +
      '\nTotal Debt: $' + totalDebt.toFixed(2),
      { htmlBody: html, name: 'FinFlow Reports' });
    logEntry('REPORT', email, 'Monthly', 'Sent', subject);
  } catch (e) {
    logEntry('REPORT', email, 'Monthly', 'FAILED', e.toString());
  }
}

// ==================== BUDGET ALERTS ====================

function checkBudgetAlerts() {
  var income       = getMonthlyIncome();
  var expenses     = getMonthlyExpenses();
  var debtPayments = getMonthlyDebtPayments();
  var threshold    = parseFloat(getSetting('budgetAlertThreshold')) || 90;
  var email        = getSetting('email');
  if (!email) return;

  var alerts = [];

  var obligationsRatio = income > 0 ? (expenses + debtPayments) / income * 100 : 0;
  if (obligationsRatio > threshold) {
    alerts.push({
      icon: '💸',
      title: 'Budget Threshold Exceeded',
      body: 'Total obligations ($' + (expenses + debtPayments).toFixed(2) + ') are ' +
        obligationsRatio.toFixed(1) + '% of income. Threshold: ' + threshold + '%.',
      color: '#ff9f43'
    });
  }

  getDebts().filter(function(d){return d.included && d.creditLimit > 0;}).forEach(function(d) {
    var util = d.balance / d.creditLimit * 100;
    if (util > 80) {
      alerts.push({
        icon: '💳',
        title: 'High Credit Utilization: ' + d.name,
        body: util.toFixed(1) + '% utilized ($' + d.balance.toFixed(2) + ' / $' + d.creditLimit.toFixed(2) + '). ' +
          'High utilization hurts your credit score.',
        color: '#ff5b6b'
      });
    }
  });

  var surplus = getNetSurplus();
  if (surplus < 0) {
    alerts.push({
      icon: '⚠️',
      title: 'Negative Monthly Surplus',
      body: 'You are spending $' + Math.abs(surplus).toFixed(2) + ' more than you earn each month. Review your expenses.',
      color: '#ff5b6b'
    });
  }

  if (alerts.length === 0) {
    logEntry('ALERT', email, 'Budget Check', 'OK', 'No alerts triggered');
    return;
  }

  var alertsHtml = alerts.map(function(a) {
    return '<div style="background:' + a.color + '18;border-left:4px solid ' + a.color + ';padding:14px 18px;border-radius:6px;margin:10px 0;">' +
      '<div style="font-weight:bold;color:' + a.color + ';margin-bottom:4px;">' + a.icon + ' ' + a.title + '</div>' +
      '<div style="color:#444;font-size:13px;">' + a.body + '</div>' +
      '</div>';
  }).join('');

  var subject = '⚠️ FinFlow: ' + alerts.length + ' financial alert(s) need attention';

  var html = '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0f2f5;">' +
    '<div style="font-family:Arial,sans-serif;max-width:580px;margin:30px auto;">' +
    '<div style="background:#ff9f43;padding:24px;border-radius:12px 12px 0 0;">' +
    '<h1 style="color:white;margin:0;font-size:20px;">⚠️ FinFlow Financial Alerts</h1>' +
    '</div>' +
    '<div style="background:white;padding:24px;border-radius:0 0 12px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08);">' +
    alertsHtml +
    '<p style="color:#aaa;font-size:11px;margin-top:20px;text-align:center;">FinFlow Alerts · ' + new Date().toLocaleString() + '</p>' +
    '</div></div></body></html>';

  try {
    GmailApp.sendEmail(email, subject, alerts.map(function(a){return a.title + ': ' + a.body;}).join('\n'),
      { htmlBody: html, name: 'FinFlow Alerts' });
    logEntry('ALERT', email, 'Budget', 'Sent', alerts.length + ' alert(s)');
  } catch (e) {
    logEntry('ALERT', email, 'Budget', 'FAILED', e.toString());
  }
}

// ==================== HTML HELPERS ====================

function _kpiBox(label, value, color, bg) {
  return '<div style="flex:1;min-width:120px;background:' + bg + ';padding:14px;border-radius:10px;border-left:4px solid ' + color + ';text-align:center;">' +
    '<div style="font-size:11px;color:#888;margin-bottom:4px;">' + label + '</div>' +
    '<div style="font-size:20px;font-weight:bold;color:' + color + ';">' + value + '</div>' +
    '</div>';
}

function _budgetBar(label, amount, total, color) {
  var pct = total > 0 ? Math.min(100, amount / total * 100).toFixed(1) : 0;
  return '<div style="margin-bottom:12px;">' +
    '<div style="display:flex;justify-content:space-between;font-size:13px;margin-bottom:4px;">' +
    '<span>' + label + '</span>' +
    '<span style="color:' + color + ';font-weight:bold;">$' + amount.toFixed(2) + ' (' + pct + '%)</span>' +
    '</div>' +
    '<div style="background:#f0f0f0;border-radius:4px;height:10px;overflow:hidden;">' +
    '<div style="background:' + color + ';width:' + pct + '%;height:100%;border-radius:4px;"></div>' +
    '</div></div>';
}

// ==================== TRIGGERS ====================

function setupTriggers() {
  removeTriggers();

  ScriptApp.newTrigger('sendBillReminders')
    .timeBased().atHour(9).everyDays(1).create();

  ScriptApp.newTrigger('checkBudgetAlerts')
    .timeBased().atHour(8).everyDays(1).create();

  ScriptApp.newTrigger('sendWeeklyDigest')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();

  ScriptApp.newTrigger('_monthlyReportCheck')
    .timeBased().onMonthDay(1).atHour(7).create();

  SpreadsheetApp.getUi().alert(
    '✅ Automated triggers activated!\n\n' +
    '• Bill reminders:    Daily at 9:00 AM\n' +
    '• Budget alerts:     Daily at 8:00 AM\n' +
    '• Weekly digest:     Every Monday at 8:00 AM\n' +
    '• Monthly report:    1st of each month at 7:00 AM\n\n' +
    'All times are in your Google account timezone.\n' +
    'You can change the monthly report day in Settings.'
  );
}

function removeTriggers() {
  var names = ['sendBillReminders','checkBudgetAlerts','sendWeeklyDigest','_monthlyReportCheck'];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (names.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
}

function _monthlyReportCheck() {
  var reportDay = parseInt(getSetting('monthlyReportDay')) || 1;
  if (new Date().getDate() === reportDay) generateMonthlyReport();
}

// ==================== LOGGING ====================

function logEntry(type, recipient, subject, status, message) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
  if (!s) return;
  s.appendRow([new Date().toLocaleString(), type, recipient, subject, status, message]);
  // Keep log to last 500 rows
  if (s.getLastRow() > 502) s.deleteRows(2, s.getLastRow() - 502);
}

// ==================== SIDEBAR UI ====================

function showDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('FinFlow Dashboard').setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showSettings() {
  var html = HtmlService.createHtmlOutputFromFile('Settings')
    .setTitle('FinFlow Settings').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showPayoffPlan() {
  var strategy = getSetting('payoffStrategy') || 'avalanche';
  var plan = strategy === 'avalanche' ? calculateAvalanche() : calculateSnowball();
  if (!plan) {
    SpreadsheetApp.getUi().alert('No active debts found. Add debts to the Debts sheet first.');
    return;
  }
  var html = HtmlService.createHtmlOutput(_payoffPlanHtml(plan, strategy))
    .setTitle('Payoff Plan').setWidth(440);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showForecast() {
  var months   = parseInt(getSetting('forecastMonths')) || 12;
  var forecast = generateForecast(months);
  var html = HtmlService.createHtmlOutput(_forecastHtml(forecast, months))
    .setTitle(months + '-Month Forecast').setWidth(560);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ==================== DATA FOR SIDEBAR ====================

function getDataForSidebar() {
  var debts    = getDebts();
  var income   = getMonthlyIncome();
  var expenses = getMonthlyExpenses();
  var surplus  = getNetSurplus();
  var debt     = getTotalDebt();
  var interest = getMonthlyInterestCost();
  var goals    = getGoals();

  var today      = new Date();
  var curDay     = today.getDate();
  var dim        = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();

  var upcoming = debts.filter(function(d){return d.included;}).map(function(d) {
    var until = d.dueDay > curDay ? d.dueDay - curDay : d.dueDay < curDay ? (dim - curDay) + d.dueDay : 0;
    return { name: d.name, minPayment: d.minPayment, dueDay: d.dueDay, daysUntil: until };
  }).sort(function(a,b){return a.daysUntil - b.daysUntil;}).slice(0, 5);

  return {
    income:          income,
    expenses:        expenses,
    surplus:         surplus,
    debt:            debt,
    interest:        interest,
    debtPayments:    getMonthlyDebtPayments(),
    debts:           debts,
    goals:           goals,
    upcoming:        upcoming,
    settings:        getAllSettings()
  };
}

function testNotification() {
  var enableEmail = getSetting('enableEmailNotifications') !== 'false';
  var enablePhone = getSetting('enablePhoneNotifications') === 'true';
  var results = [];

  if (enableEmail) {
    var email = getSetting('email');
    try {
      GmailApp.sendEmail(email,
        '✅ FinFlow Test Notification',
        'This is a test from FinFlow. Your email notifications are working!',
        { name: 'FinFlow' });
      results.push('Email ✅ sent to ' + email);
    } catch (e) {
      results.push('Email ❌ ' + e.message);
    }
  }

  if (enablePhone) {
    try {
      sendPhoneNotification('FinFlow test: Phone notifications are working! ✅');
      results.push('Phone ✅ notification sent');
    } catch (e) {
      results.push('Phone ❌ ' + e.message);
    }
  }

  if (!enableEmail && !enablePhone) results.push('No notification methods enabled. Check Settings.');
  return { success: true, message: results.join('\n') };
}

// ==================== INLINE SIDEBAR HTML ====================

function _payoffPlanHtml(plan, strategy) {
  var rows = plan.results.map(function(r) {
    return '<div style="background:#f8f9fa;padding:12px 14px;border-radius:8px;margin:8px 0;border-left:3px solid #5c7cff;">' +
      '<div style="font-weight:bold;margin-bottom:4px;">' + r.name + '</div>' +
      '<div style="font-size:12px;color:#666;">Payoff: ' +
        (r.payoffMonth ? 'Month ' + r.payoffMonth : 'Not paid off in 30 years') +
        ' &nbsp;·&nbsp; Interest: <span style="color:#ff9f43;">$' + r.totalInterest.toFixed(2) + '</span></div>' +
      '</div>';
  }).join('');

  return '<html><head><style>body{font-family:Arial,sans-serif;padding:16px;background:#f5f5f5;margin:0;}' +
    'h2{color:#1a1a2e;margin-top:0;}.card{background:white;padding:16px;border-radius:10px;margin-bottom:14px;box-shadow:0 2px 8px rgba(0,0,0,.08);}</style></head>' +
    '<body>' +
    '<h2>📊 ' + strategy.charAt(0).toUpperCase()+strategy.slice(1) + ' Payoff Plan</h2>' +
    '<div class="card">' +
    '<div style="display:flex;justify-content:space-between;margin-bottom:8px;"><span>Debt-Free Month</span><strong style="color:#22d47e;">Month ' + plan.totalMonths + '</strong></div>' +
    '<div style="display:flex;justify-content:space-between;margin-bottom:8px;"><span>Estimated Date</span><strong>' + plan.payoffDate + '</strong></div>' +
    '<div style="display:flex;justify-content:space-between;"><span>Total Interest</span><strong style="color:#ff9f43;">$' + plan.totalInterest.toFixed(2) + '</strong></div>' +
    '</div>' +
    '<h3 style="color:#1a1a2e;margin-bottom:8px;">Per Account</h3>' + rows +
    '</body></html>';
}

function _forecastHtml(forecast, months) {
  var rows = forecast.map(function(r) {
    var sc = r.surplus >= 0 ? '#22d47e' : '#ff5b6b';
    return '<tr>' +
      '<td style="padding:6px 8px;white-space:nowrap;font-size:11px;">' + r.month + '</td>' +
      '<td style="padding:6px 8px;text-align:right;color:#22d47e;font-size:11px;">$' + r.income.toFixed(0) + '</td>' +
      '<td style="padding:6px 8px;text-align:right;color:#ff5b6b;font-size:11px;">$' + r.expenses.toFixed(0) + '</td>' +
      '<td style="padding:6px 8px;text-align:right;color:#ff9f43;font-size:11px;">$' + r.debtPayments.toFixed(0) + '</td>' +
      '<td style="padding:6px 8px;text-align:right;color:' + sc + ';font-size:11px;font-weight:bold;">$' + r.surplus.toFixed(0) + '</td>' +
      '<td style="padding:6px 8px;text-align:right;font-size:11px;">$' + r.totalDebt.toFixed(0) + '</td>' +
      '</tr>';
  }).join('');

  return '<html><head><style>body{font-family:Arial,sans-serif;padding:14px;margin:0;overflow-x:auto;}' +
    'table{border-collapse:collapse;min-width:100%;background:white;border-radius:8px;overflow:hidden;}' +
    'h2{color:#1a1a2e;margin-top:0;font-size:16px;}</style></head><body>' +
    '<h2>📅 ' + months + '-Month Forecast</h2>' +
    '<table><thead><tr style="background:#1a1a2e;color:white;">' +
    '<th style="padding:8px;text-align:left;font-size:11px;">Month</th>' +
    '<th style="padding:8px;text-align:right;font-size:11px;">Income</th>' +
    '<th style="padding:8px;text-align:right;font-size:11px;">Expenses</th>' +
    '<th style="padding:8px;text-align:right;font-size:11px;">Debt</th>' +
    '<th style="padding:8px;text-align:right;font-size:11px;">Surplus</th>' +
    '<th style="padding:8px;text-align:right;font-size:11px;">Total Debt</th>' +
    '</tr></thead><tbody>' + rows + '</tbody></table></body></html>';
}
