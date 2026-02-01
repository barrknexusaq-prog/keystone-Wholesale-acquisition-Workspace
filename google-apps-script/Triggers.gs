/**
 * Triggers.gs - Time-based and Event Triggers
 *
 * This file contains functions for automated triggers
 * Set these up via Apps Script > Triggers menu
 */

// ============================================================================
// TRIGGER SETUP FUNCTIONS
// ============================================================================

/**
 * Set up all recommended triggers
 * Run this function once to configure automation
 */
function setupAllTriggers() {
  // Remove existing triggers first
  removeAllTriggers();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Daily summary at 8 AM
  ScriptApp.newTrigger('dailyMorningReport')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  // Weekly KPI report on Monday at 9 AM
  ScriptApp.newTrigger('weeklyKPIReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  // Monthly reset on 1st of month
  ScriptApp.newTrigger('monthlyReset')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();

  // Auto-archive old leads weekly
  ScriptApp.newTrigger('archiveOldLeads')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(2)
    .create();

  SpreadsheetApp.getUi().alert('All triggers have been set up!');
}

/**
 * Remove all existing triggers
 */
function removeAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * List all current triggers
 */
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let info = 'Current Triggers:\n\n';

  triggers.forEach(trigger => {
    info += `- ${trigger.getHandlerFunction()} (${trigger.getEventType()})\n`;
  });

  if (triggers.length === 0) {
    info = 'No triggers currently set up.';
  }

  SpreadsheetApp.getUi().alert(info);
}

// ============================================================================
// DAILY AUTOMATION
// ============================================================================

/**
 * Daily morning report - runs at 8 AM
 */
function dailyMorningReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timezone = Session.getScriptTimeZone();
  const today = new Date();
  const yesterday = new Date(today.getTime() - 24 * 60 * 60 * 1000);

  // Get yesterday's date string for comparison
  const yesterdayStr = Utilities.formatDate(yesterday, timezone, 'yyyy-MM-dd');

  // Count metrics
  const mainPage = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);
  const pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);

  let newLeadsYesterday = 0;
  let totalActiveLeads = 0;
  let offersOutstanding = 0;
  let underContract = 0;

  if (mainPage) {
    const data = mainPage.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const dateAdded = data[i][1];
      const status = data[i][19]; // Status column

      if (dateAdded instanceof Date) {
        const dateStr = Utilities.formatDate(dateAdded, timezone, 'yyyy-MM-dd');
        if (dateStr === yesterdayStr) {
          newLeadsYesterday++;
        }
      }

      if (status && status !== 'Closed' && status !== 'Dead') {
        totalActiveLeads++;
      }
      if (status === 'Offer Pending') {
        offersOutstanding++;
      }
    }
  }

  if (pipeline) {
    const pipelineData = pipeline.getDataRange().getValues();
    for (let i = 1; i < pipelineData.length; i++) {
      if (pipelineData[i][22] === 'Under Contract') { // Status column
        underContract++;
      }
    }
  }

  // Update dashboard with timestamp
  const dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (dashboard) {
    dashboard.getRange('A2').setValue('Last Updated: ' + today.toLocaleString());
  }

  // Log the report (or send email if configured)
  Logger.log(`Daily Report for ${yesterdayStr}:`);
  Logger.log(`- New Leads: ${newLeadsYesterday}`);
  Logger.log(`- Total Active: ${totalActiveLeads}`);
  Logger.log(`- Offers Outstanding: ${offersOutstanding}`);
  Logger.log(`- Under Contract: ${underContract}`);

  // Uncomment to send email:
  /*
  const emailBody = `
  Good morning! Here's your daily acquisition update:

  Yesterday's Activity:
  - New Leads Added: ${newLeadsYesterday}

  Current Pipeline:
  - Total Active Leads: ${totalActiveLeads}
  - Offers Outstanding: ${offersOutstanding}
  - Properties Under Contract: ${underContract}

  View Dashboard: ${ss.getUrl()}
  `;

  MailApp.sendEmail({
    to: 'your-email@example.com',
    subject: 'Daily Acquisition Report - ' + Utilities.formatDate(today, timezone, 'MMM d, yyyy'),
    body: emailBody
  });
  */
}

/**
 * Check for leads needing follow-up (no activity in 3+ days)
 */
function checkStaleLeads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainPage = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);

  if (!mainPage) return;

  const timezone = Session.getScriptTimeZone();
  const today = new Date();
  const threeDaysAgo = new Date(today.getTime() - 3 * 24 * 60 * 60 * 1000);

  const data = mainPage.getDataRange().getValues();
  const staleLeads = [];

  for (let i = 1; i < data.length; i++) {
    const dateAdded = data[i][1];
    const status = data[i][19];
    const address = data[i][4];

    // Check if lead is stale (active but no recent activity)
    if (dateAdded instanceof Date && dateAdded < threeDaysAgo) {
      if (status === 'New' || status === 'Contacted') {
        staleLeads.push({
          row: i + 1,
          address: address,
          status: status,
          daysOld: Math.floor((today - dateAdded) / (24 * 60 * 60 * 1000))
        });
      }
    }
  }

  if (staleLeads.length > 0) {
    Logger.log(`Found ${staleLeads.length} stale leads needing follow-up`);
    staleLeads.forEach(lead => {
      Logger.log(`- Row ${lead.row}: ${lead.address} (${lead.status}, ${lead.daysOld} days old)`);
    });
  }

  return staleLeads;
}

// ============================================================================
// WEEKLY AUTOMATION
// ============================================================================

/**
 * Weekly KPI report - runs Monday at 9 AM
 */
function weeklyKPIReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kpiSheet = ss.getSheetByName(CONFIG.SHEETS.LEAD_KPIS);

  if (!kpiSheet) return;

  // Force recalculation
  SpreadsheetApp.flush();

  // Get KPI values
  const leadsGenerated = kpiSheet.getRange('C6').getValue();
  const leadsQualified = kpiSheet.getRange('C7').getValue();
  const offersMade = kpiSheet.getRange('C9').getValue();
  const contractsSigned = kpiSheet.getRange('C10').getValue();
  const dealsClosed = kpiSheet.getRange('C11').getValue();
  const totalRevenue = kpiSheet.getRange('C12').getValue();

  // Get conversion rates
  const leadToQualified = kpiSheet.getRange('C18').getValue();
  const qualifiedToOffer = kpiSheet.getRange('C19').getValue();
  const offerToContract = kpiSheet.getRange('C20').getValue();

  Logger.log('Weekly KPI Report:');
  Logger.log(`- Leads Generated: ${leadsGenerated}`);
  Logger.log(`- Leads Qualified: ${leadsQualified}`);
  Logger.log(`- Offers Made: ${offersMade}`);
  Logger.log(`- Contracts Signed: ${contractsSigned}`);
  Logger.log(`- Deals Closed: ${dealsClosed}`);
  Logger.log(`- Total Revenue: $${totalRevenue}`);
  Logger.log('Conversion Rates:');
  Logger.log(`- Lead to Qualified: ${(leadToQualified * 100).toFixed(1)}%`);
  Logger.log(`- Qualified to Offer: ${(qualifiedToOffer * 100).toFixed(1)}%`);
  Logger.log(`- Offer to Contract: ${(offerToContract * 100).toFixed(1)}%`);

  // Uncomment to send email report
  /*
  const emailBody = `
  Weekly KPI Report
  ==================

  Key Metrics This Month:
  - Leads Generated: ${leadsGenerated}
  - Leads Qualified: ${leadsQualified}
  - Offers Made: ${offersMade}
  - Contracts Signed: ${contractsSigned}
  - Deals Closed: ${dealsClosed}
  - Total Revenue: $${totalRevenue.toLocaleString()}

  Conversion Rates:
  - Lead to Qualified: ${(leadToQualified * 100).toFixed(1)}%
  - Qualified to Offer: ${(qualifiedToOffer * 100).toFixed(1)}%
  - Offer to Contract: ${(offerToContract * 100).toFixed(1)}%

  View Full KPIs: ${ss.getUrl()}
  `;

  MailApp.sendEmail({
    to: 'your-email@example.com',
    subject: 'Weekly Acquisition KPI Report',
    body: emailBody
  });
  */
}

/**
 * Archive old dead leads (older than 90 days)
 */
function archiveOldLeads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainPage = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);

  if (!mainPage) return;

  // Create or get archive sheet
  let archive = ss.getSheetByName('Lead Archive');
  if (!archive) {
    archive = ss.insertSheet('Lead Archive');
    // Copy headers
    const headers = mainPage.getRange(1, 1, 1, mainPage.getLastColumn()).getValues();
    archive.getRange(1, 1, 1, headers[0].length).setValues(headers);
    archive.getRange(1, 1, 1, headers[0].length).setBackground('#666666').setFontColor('white');
  }

  const timezone = Session.getScriptTimeZone();
  const today = new Date();
  const ninetyDaysAgo = new Date(today.getTime() - 90 * 24 * 60 * 60 * 1000);

  const data = mainPage.getDataRange().getValues();
  const rowsToArchive = [];

  // Find rows to archive (from bottom to top to avoid index shifting)
  for (let i = data.length - 1; i >= 1; i--) {
    const dateAdded = data[i][1];
    const status = data[i][19];

    if (dateAdded instanceof Date && dateAdded < ninetyDaysAgo && status === 'Dead') {
      rowsToArchive.push({ row: i + 1, data: data[i] });
    }
  }

  // Archive and delete
  rowsToArchive.forEach(item => {
    archive.appendRow(item.data);
    mainPage.deleteRow(item.row);
  });

  Logger.log(`Archived ${rowsToArchive.length} old dead leads`);
}

// ============================================================================
// MONTHLY AUTOMATION
// ============================================================================

/**
 * Monthly reset and archival
 */
function monthlyReset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timezone = Session.getScriptTimeZone();
  const today = new Date();
  const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const monthName = Utilities.formatDate(lastMonth, timezone, 'MMMM yyyy');

  // Archive KPIs to history
  archiveMonthlyKPIs(monthName);

  // Update KPI sheet with new month
  createKPISheet(ss);

  Logger.log(`Monthly reset completed for ${monthName}`);
}

/**
 * Archive monthly KPIs to history sheet
 */
function archiveMonthlyKPIs(monthName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create or get KPI history sheet
  let history = ss.getSheetByName('KPI History');
  if (!history) {
    history = ss.insertSheet('KPI History');
    history.getRange('A1:H1').setValues([[
      'Month', 'Leads Generated', 'Leads Qualified', 'Offers Made',
      'Contracts Signed', 'Deals Closed', 'Revenue', 'Lead to Close %'
    ]]);
    history.getRange('A1:H1').setBackground('#1E3A5F').setFontColor('white').setFontWeight('bold');
  }

  // Get current KPI values
  const kpiSheet = ss.getSheetByName(CONFIG.SHEETS.LEAD_KPIS);
  if (!kpiSheet) return;

  const kpiData = [
    monthName,
    kpiSheet.getRange('C6').getValue(),  // Leads Generated
    kpiSheet.getRange('C7').getValue(),  // Leads Qualified
    kpiSheet.getRange('C9').getValue(),  // Offers Made
    kpiSheet.getRange('C10').getValue(), // Contracts Signed
    kpiSheet.getRange('C11').getValue(), // Deals Closed
    kpiSheet.getRange('C12').getValue(), // Revenue
    kpiSheet.getRange('C22').getValue()  // Lead to Close %
  ];

  history.appendRow(kpiData);
}

// ============================================================================
// CONTRACT DEADLINE MONITORING
// ============================================================================

/**
 * Check for upcoming contract deadlines
 * Run daily to alert about inspection periods and close dates
 */
function checkContractDeadlines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);

  if (!pipeline) return;

  const today = new Date();
  const threeDaysFromNow = new Date(today.getTime() + 3 * 24 * 60 * 60 * 1000);
  const sevenDaysFromNow = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);

  const data = pipeline.getDataRange().getValues();
  const headers = data[0];

  // Find column indices
  const addressCol = headers.indexOf('Property Address');
  const inspectionCol = headers.indexOf('Inspection Period End');
  const closeCol = headers.indexOf('Close Date');
  const statusCol = headers.indexOf('Status');

  const urgentDeadlines = [];
  const upcomingDeadlines = [];

  for (let i = 1; i < data.length; i++) {
    const status = data[i][statusCol];
    if (status === 'Closed' || status === 'Cancelled') continue;

    const address = data[i][addressCol];
    const inspectionDate = data[i][inspectionCol];
    const closeDate = data[i][closeCol];

    // Check inspection deadline
    if (inspectionDate instanceof Date) {
      if (inspectionDate <= threeDaysFromNow && inspectionDate >= today) {
        urgentDeadlines.push({
          address: address,
          type: 'Inspection Period',
          date: inspectionDate
        });
      } else if (inspectionDate <= sevenDaysFromNow && inspectionDate > threeDaysFromNow) {
        upcomingDeadlines.push({
          address: address,
          type: 'Inspection Period',
          date: inspectionDate
        });
      }
    }

    // Check close deadline
    if (closeDate instanceof Date) {
      if (closeDate <= threeDaysFromNow && closeDate >= today) {
        urgentDeadlines.push({
          address: address,
          type: 'Close Date',
          date: closeDate
        });
      } else if (closeDate <= sevenDaysFromNow && closeDate > threeDaysFromNow) {
        upcomingDeadlines.push({
          address: address,
          type: 'Close Date',
          date: closeDate
        });
      }
    }
  }

  // Log results
  if (urgentDeadlines.length > 0) {
    Logger.log('URGENT DEADLINES (within 3 days):');
    urgentDeadlines.forEach(d => {
      Logger.log(`- ${d.address}: ${d.type} on ${d.date.toLocaleDateString()}`);
    });
  }

  if (upcomingDeadlines.length > 0) {
    Logger.log('Upcoming Deadlines (within 7 days):');
    upcomingDeadlines.forEach(d => {
      Logger.log(`- ${d.address}: ${d.type} on ${d.date.toLocaleDateString()}`);
    });
  }

  return { urgent: urgentDeadlines, upcoming: upcomingDeadlines };
}

// ============================================================================
// DATA CLEANUP & MAINTENANCE
// ============================================================================

/**
 * Remove duplicate leads based on address
 */
function findDuplicateLeads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainPage = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);

  if (!mainPage) return;

  const data = mainPage.getDataRange().getValues();
  const addresses = {};
  const duplicates = [];

  for (let i = 1; i < data.length; i++) {
    const address = data[i][4]?.toString().toLowerCase().trim();
    if (!address) continue;

    if (addresses[address]) {
      duplicates.push({
        address: data[i][4],
        rows: [addresses[address], i + 1]
      });
    } else {
      addresses[address] = i + 1;
    }
  }

  if (duplicates.length > 0) {
    Logger.log('Found duplicate leads:');
    duplicates.forEach(d => {
      Logger.log(`- "${d.address}" in rows ${d.rows.join(', ')}`);
    });
  } else {
    Logger.log('No duplicate leads found');
  }

  return duplicates;
}

/**
 * Validate data integrity
 */
function validateData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const issues = [];

  // Check Main Page
  const mainPage = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);
  if (mainPage) {
    const data = mainPage.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = i + 1;
      const leadId = data[i][0];
      const address = data[i][4];
      const zip = data[i][7];

      if (!leadId) {
        issues.push(`Row ${row}: Missing Lead ID`);
      }
      if (!address) {
        issues.push(`Row ${row}: Missing Address`);
      }
      if (zip && !/^\d{5}$/.test(zip.toString())) {
        issues.push(`Row ${row}: Invalid Zip Code "${zip}"`);
      }
    }
  }

  Logger.log(`Data validation found ${issues.length} issues`);
  issues.forEach(issue => Logger.log(`- ${issue}`));

  return issues;
}
