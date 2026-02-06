/**
 * Triggers.gs - Time-based Automation
 * Acquisition Workspace v4.0
 */

// ============================================================================
// TRIGGER SETUP
// ============================================================================

/**
 * Setup all recommended triggers
 */
function setupAllTriggers() {
  removeAllTriggers();

  // Daily morning report at 8 AM
  ScriptApp.newTrigger('dailyMorningReport')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  // Weekly summary on Monday at 9 AM
  ScriptApp.newTrigger('weeklySummary')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  // Check contract deadlines daily at 7 AM
  ScriptApp.newTrigger('checkContractDeadlines')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();

  // Auto-archive monthly on 1st at 6 AM
  ScriptApp.newTrigger('monthlyAutoArchive')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();

  SpreadsheetApp.getUi().alert('All triggers configured!');
}

/**
 * Remove all existing triggers
 */
function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * List current triggers
 */
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let info = 'Current Triggers:\n\n';

  triggers.forEach(trigger => {
    info += `- ${trigger.getHandlerFunction()}\n`;
  });

  if (triggers.length === 0) {
    info = 'No triggers set up. Run setupAllTriggers() to configure.';
  }

  SpreadsheetApp.getUi().alert(info);
}

// ============================================================================
// DAILY AUTOMATION
// ============================================================================

/**
 * Daily morning report
 */
function dailyMorningReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const onMarket = ss.getSheetByName(CONFIG.SHEETS.ON_MARKET);
  const offMarket = ss.getSheetByName(CONFIG.SHEETS.OFF_MARKET);
  const pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);

  // Count metrics
  const totalOnMarket = onMarket ? Math.max(0, onMarket.getLastRow() - 1) : 0;
  const totalOffMarket = offMarket ? Math.max(0, offMarket.getLastRow() - 1) : 0;

  let underContract = 0;
  let pendingClose = 0;

  if (pipeline) {
    const statusData = pipeline.getRange('Y2:Y' + pipeline.getLastRow()).getValues();
    statusData.forEach(row => {
      if (row[0] === 'Under Contract') underContract++;
      if (row[0] === 'Pending Close') pendingClose++;
    });
  }

  Logger.log('Daily Report:');
  Logger.log(`- On Market Leads: ${totalOnMarket}`);
  Logger.log(`- Off Market Leads: ${totalOffMarket}`);
  Logger.log(`- Under Contract: ${underContract}`);
  Logger.log(`- Pending Close: ${pendingClose}`);

  // Update dashboard timestamp
  const dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (dashboard) {
    dashboard.getRange('A2').setValue('Last Updated: ' + new Date().toLocaleString());
  }

  // Uncomment to send email:
  /*
  const body = `
Good morning! Here's your daily acquisition update:

Active Leads:
- On Market: ${totalOnMarket}
- Off Market: ${totalOffMarket}

Pipeline:
- Under Contract: ${underContract}
- Pending Close: ${pendingClose}

View Dashboard: ${ss.getUrl()}
  `;

  MailApp.sendEmail({
    to: 'your-email@example.com',
    subject: 'Daily Acquisition Report - ' + new Date().toLocaleDateString(),
    body: body
  });
  */
}

/**
 * Check for upcoming contract deadlines
 */
function checkContractDeadlines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);

  if (!pipeline || pipeline.getLastRow() < 2) return;

  const today = new Date();
  const threeDays = new Date(today.getTime() + 3 * 24 * 60 * 60 * 1000);
  const sevenDays = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);

  const data = pipeline.getDataRange().getValues();
  const headers = data[0];

  const addressCol = headers.indexOf('Property Address');
  const inspectionCol = headers.indexOf('Inspection Period End');
  const closeCol = headers.indexOf('Close Date');
  const statusCol = headers.indexOf('Status');

  const urgent = [];
  const upcoming = [];

  for (let i = 1; i < data.length; i++) {
    const status = data[i][statusCol];
    if (status === 'Closed' || status === 'Cancelled') continue;

    const address = data[i][addressCol];
    const inspection = data[i][inspectionCol];
    const close = data[i][closeCol];

    // Check inspection deadline
    if (inspection instanceof Date) {
      if (inspection <= threeDays && inspection >= today) {
        urgent.push({ address, type: 'Inspection', date: inspection });
      } else if (inspection <= sevenDays && inspection > threeDays) {
        upcoming.push({ address, type: 'Inspection', date: inspection });
      }
    }

    // Check close deadline
    if (close instanceof Date) {
      if (close <= threeDays && close >= today) {
        urgent.push({ address, type: 'Close', date: close });
      } else if (close <= sevenDays && close > threeDays) {
        upcoming.push({ address, type: 'Close', date: close });
      }
    }
  }

  if (urgent.length > 0) {
    Logger.log('URGENT DEADLINES (within 3 days):');
    urgent.forEach(d => Logger.log(`- ${d.address}: ${d.type} on ${d.date.toLocaleDateString()}`));
  }

  if (upcoming.length > 0) {
    Logger.log('Upcoming deadlines (within 7 days):');
    upcoming.forEach(d => Logger.log(`- ${d.address}: ${d.type} on ${d.date.toLocaleDateString()}`));
  }

  // Uncomment to send email alert for urgent deadlines:
  /*
  if (urgent.length > 0) {
    let body = 'URGENT - The following deadlines are within 3 days:\n\n';
    urgent.forEach(d => {
      body += `${d.address}: ${d.type} deadline on ${d.date.toLocaleDateString()}\n`;
    });

    MailApp.sendEmail({
      to: 'your-email@example.com',
      subject: 'URGENT: Contract Deadlines Approaching',
      body: body
    });
  }
  */

  return { urgent, upcoming };
}

// ============================================================================
// WEEKLY AUTOMATION
// ============================================================================

/**
 * Weekly summary report
 */
function weeklySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kpiSheet = ss.getSheetByName(CONFIG.SHEETS.LEAD_KPIS);

  if (!kpiSheet) return;

  // Force recalculate
  SpreadsheetApp.flush();

  // Get KPI values
  const totalLeads = kpiSheet.getRange('C6').getValue();
  const offersMade = kpiSheet.getRange('C9').getValue();
  const contractsSigned = kpiSheet.getRange('C10').getValue();
  const dealsClosed = kpiSheet.getRange('C12').getValue();
  const revenue = kpiSheet.getRange('C13').getValue();

  Logger.log('Weekly Summary:');
  Logger.log(`- Total Leads: ${totalLeads}`);
  Logger.log(`- Offers Made: ${offersMade}`);
  Logger.log(`- Contracts Signed: ${contractsSigned}`);
  Logger.log(`- Deals Closed: ${dealsClosed}`);
  Logger.log(`- Revenue: $${revenue}`);

  // Uncomment to send email:
  /*
  const body = `
Weekly Acquisition Summary

Metrics This Month:
- Total Leads: ${totalLeads}
- Offers Made: ${offersMade}
- Contracts Signed: ${contractsSigned}
- Deals Closed: ${dealsClosed}
- Revenue: $${revenue.toLocaleString()}

View Full KPIs: ${ss.getUrl()}
  `;

  MailApp.sendEmail({
    to: 'your-email@example.com',
    subject: 'Weekly Acquisition Summary',
    body: body
  });
  */
}

// ============================================================================
// MONTHLY AUTOMATION
// ============================================================================

/**
 * Auto-archive KPIs on 1st of month
 */
function monthlyAutoArchive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kpiSheet = ss.getSheetByName(CONFIG.SHEETS.LEAD_KPIS);
  const yearlySheet = ss.getSheetByName(CONFIG.SHEETS.YEARLY_KPIS);

  if (!kpiSheet || !yearlySheet) return;

  // Get previous month (since this runs on 1st of new month)
  const now = new Date();
  const prevMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const monthIndex = prevMonth.getMonth(); // 0-11
  const monthCol = monthIndex + 2; // Column B=Jan(2), C=Feb(3), etc.

  // Archive previous month's data
  const totalLeads = kpiSheet.getRange('C6').getValue();
  const onMarketLeads = kpiSheet.getRange('C7').getValue();
  const offMarketLeads = kpiSheet.getRange('C8').getValue();
  const offersMade = kpiSheet.getRange('C9').getValue();
  const contractsSigned = kpiSheet.getRange('C10').getValue();
  const dealsClosed = kpiSheet.getRange('C12').getValue();
  const revenue = kpiSheet.getRange('C13').getValue();

  yearlySheet.getRange(4, monthCol).setValue(totalLeads);
  yearlySheet.getRange(5, monthCol).setValue(onMarketLeads);
  yearlySheet.getRange(6, monthCol).setValue(offMarketLeads);
  yearlySheet.getRange(7, monthCol).setValue(offersMade);
  yearlySheet.getRange(8, monthCol).setValue(contractsSigned);
  yearlySheet.getRange(9, monthCol).setValue(dealsClosed);
  yearlySheet.getRange(10, monthCol).setValue(revenue);

  if (dealsClosed > 0) {
    yearlySheet.getRange(11, monthCol).setValue(revenue / dealsClosed);
  }
  if (totalLeads > 0) {
    yearlySheet.getRange(12, monthCol).setValue(dealsClosed / totalLeads);
  }

  Logger.log(`Monthly KPIs archived for ${prevMonth.toLocaleDateString('en-US', { month: 'long', year: 'numeric' })}`);
}

// ============================================================================
// TEAM PERFORMANCE TRACKING
// ============================================================================

/**
 * Get team member performance
 */
function getTeamPerformance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const onMarket = ss.getSheetByName(CONFIG.SHEETS.ON_MARKET);
  const offMarket = ss.getSheetByName(CONFIG.SHEETS.OFF_MARKET);

  const performance = {};

  CONFIG.TEAM_MEMBERS.forEach(member => {
    performance[member] = { leads: 0, offers: 0, contracts: 0 };
  });

  // Count On Market
  if (onMarket && onMarket.getLastRow() > 1) {
    const data = onMarket.getRange('C2:W' + onMarket.getLastRow()).getValues();
    data.forEach(row => {
      const member = row[0]; // Team Member column
      if (performance[member]) {
        performance[member].leads++;
        if (row[19] === 'Yes') performance[member].offers++; // Offer Made
        if (row[20] === 'Yes') performance[member].contracts++; // Contract Signed
      }
    });
  }

  // Count Off Market
  if (offMarket && offMarket.getLastRow() > 1) {
    const data = offMarket.getRange('C2:V' + offMarket.getLastRow()).getValues();
    data.forEach(row => {
      const member = row[0]; // Team Member column
      if (performance[member]) {
        performance[member].leads++;
        if (row[18] === 'Yes') performance[member].offers++; // Offer Made
        if (row[19] === 'Yes') performance[member].contracts++; // Contract Signed
      }
    });
  }

  Logger.log('Team Performance:');
  Object.keys(performance).forEach(member => {
    const p = performance[member];
    const rate = p.leads > 0 ? (p.contracts / p.leads * 100).toFixed(1) : 0;
    Logger.log(`${member}: ${p.leads} leads, ${p.offers} offers, ${p.contracts} contracts (${rate}% conversion)`);
  });

  return performance;
}

// ============================================================================
// DATA MAINTENANCE
// ============================================================================

/**
 * Find duplicate leads by address
 */
function findDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const duplicates = [];

  // Check On Market
  const onMarket = ss.getSheetByName(CONFIG.SHEETS.ON_MARKET);
  if (onMarket && onMarket.getLastRow() > 1) {
    const addresses = onMarket.getRange('F2:F' + onMarket.getLastRow()).getValues().flat();
    const seen = {};
    addresses.forEach((addr, idx) => {
      const normalized = addr.toString().toLowerCase().trim();
      if (normalized && seen[normalized]) {
        duplicates.push({ sheet: 'On Market', address: addr, rows: [seen[normalized], idx + 2] });
      } else if (normalized) {
        seen[normalized] = idx + 2;
      }
    });
  }

  // Check Off Market
  const offMarket = ss.getSheetByName(CONFIG.SHEETS.OFF_MARKET);
  if (offMarket && offMarket.getLastRow() > 1) {
    const addresses = offMarket.getRange('E2:E' + offMarket.getLastRow()).getValues().flat();
    const seen = {};
    addresses.forEach((addr, idx) => {
      const normalized = addr.toString().toLowerCase().trim();
      if (normalized && seen[normalized]) {
        duplicates.push({ sheet: 'Off Market', address: addr, rows: [seen[normalized], idx + 2] });
      } else if (normalized) {
        seen[normalized] = idx + 2;
      }
    });
  }

  if (duplicates.length > 0) {
    Logger.log('Found duplicates:');
    duplicates.forEach(d => {
      Logger.log(`- ${d.sheet}: "${d.address}" in rows ${d.rows.join(', ')}`);
    });
  } else {
    Logger.log('No duplicates found.');
  }

  return duplicates;
}

/**
 * Validate data integrity
 */
function validateData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const issues = [];

  // Validate On Market
  const onMarket = ss.getSheetByName(CONFIG.SHEETS.ON_MARKET);
  if (onMarket && onMarket.getLastRow() > 1) {
    const data = onMarket.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) issues.push(`On Market row ${i + 1}: Missing Lead ID`);
      if (!data[i][5]) issues.push(`On Market row ${i + 1}: Missing Address`);
      const zip = data[i][9];
      if (zip && !/^\d{5}$/.test(zip.toString())) {
        issues.push(`On Market row ${i + 1}: Invalid Zip Code "${zip}"`);
      }
    }
  }

  // Validate Off Market
  const offMarket = ss.getSheetByName(CONFIG.SHEETS.OFF_MARKET);
  if (offMarket && offMarket.getLastRow() > 1) {
    const data = offMarket.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) issues.push(`Off Market row ${i + 1}: Missing Lead ID`);
      if (!data[i][4]) issues.push(`Off Market row ${i + 1}: Missing Address`);
      const zip = data[i][8];
      if (zip && !/^\d{5}$/.test(zip.toString())) {
        issues.push(`Off Market row ${i + 1}: Invalid Zip Code "${zip}"`);
      }
    }
  }

  Logger.log(`Data validation found ${issues.length} issues`);
  issues.forEach(issue => Logger.log(`- ${issue}`));

  return issues;
}

// ============================================================================
// FUTURE: GMAIL INTEGRATION PLACEHOLDER
// ============================================================================

/**
 * Placeholder for Gmail integration
 * When configured, this could:
 * - Auto-create leads from emails matching certain patterns
 * - Send follow-up reminders
 * - Track email communications with sellers
 */
function gmailIntegration() {
  // Future implementation
  // Will require Gmail API authorization
  Logger.log('Gmail integration not yet configured.');
}

// ============================================================================
// FUTURE: DEALMACHINE INTEGRATION PLACEHOLDER
// ============================================================================

/**
 * Placeholder for DealMachine integration
 * When configured, this could:
 * - Import leads from DealMachine
 * - Sync property data
 * - Track driving for dollars routes
 */
function dealMachineIntegration() {
  // Future implementation
  // Will require DealMachine API access
  Logger.log('DealMachine integration not yet configured.');
}
