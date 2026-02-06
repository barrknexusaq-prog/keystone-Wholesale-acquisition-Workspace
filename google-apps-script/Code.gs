/**
 * Keystone Acquisition Workspace v4.0 - Google Apps Script
 * Real Estate Lead Tracking & Automation System
 *
 * Simplified structure with auto-tracking to Contract Pipeline
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const CONFIG = {
  // Sheet Names
  SHEETS: {
    ON_MARKET: 'On Market Leads',
    OFF_MARKET: 'Off Market Leads',
    CONTRACT_PIPELINE: 'Contract Pipeline',
    LEAD_KPIS: 'Lead KPIs',
    YEARLY_KPIS: 'Yearly KPIs',
    ZIP_CODES: 'Zip Codes',
    DASHBOARD: 'Dashboard',
    BUYERS_LIST: 'Buyers List'
  },

  // Lead Sources - On Market
  ON_MARKET_SOURCES: ['Zillow', 'Redfin', 'Realtor.com', 'MLS', 'Crexi', 'LoopNet'],

  // Lead Issues - Off Market (renamed from Source)
  OFF_MARKET_ISSUES: [
    'Wholesaler',
    'Agent Referral',
    'Direct Mail',
    'Cold Call',
    'Driving for Dollars',
    'Probate',
    'Tax Lien',
    'Expired',
    'Foreclosure',
    'Tired Landlord',
    'Other'
  ],

  // Acquisition Team Members
  TEAM_MEMBERS: ['Tess Walter', 'Nick Barr'],

  // Disposition Options
  DISPOSITION: ['Assigned', 'Double Close', 'Wholetail', 'Fix & Flip', 'Buy & Hold', 'Dead Deal'],

  // Monthly Goals
  MONTHLY_GOALS: {
    LEADS_GENERATED: 100,
    OFFERS_MADE: 25,
    CONTRACTS_SIGNED: 5,
    DEALS_CLOSED: 3,
    REVENUE_TARGET: 50000
  },

  // External Sheet for Buyers Import (update with your actual sheet ID)
  BUYERS_IMPORT: {
    SPREADSHEET_ID: '', // Add your dispo_workspace 3.0 spreadsheet ID here
    SHEET_NAME: 'Buyer'
  }
};

// ============================================================================
// MENU & INITIALIZATION
// ============================================================================

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Acquisition Tools')
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('Update KPIs', 'updateKPIs')
    .addSeparator()
    .addSubMenu(ui.createMenu('Add New Lead')
      .addItem('On Market Lead', 'showOnMarketForm')
      .addItem('Off Market Lead', 'showOffMarketForm'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Setup')
      .addItem('Initialize All Sheets', 'initializeWorkspace')
      .addItem('Import Buyers from Dispo Workspace', 'importBuyersFromDispo')
      .addItem('Start New Month', 'startNewMonth')
      .addItem('Import Zip Codes', 'showZipCodeImport'))
    .addSeparator()
    .addItem('Help', 'showHelp')
    .addToUi();
}

/**
 * Initialize the entire workspace
 */
function initializeWorkspace() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Initialize Workspace',
    'This will create/update all sheets. Existing data will be preserved. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  createOnMarketSheet(ss);
  createOffMarketSheet(ss);
  createContractPipelineSheet(ss);
  createKPISheet(ss);
  createYearlyKPISheet(ss);
  createZipCodesSheet(ss);
  createDashboardSheet(ss);
  createBuyersListSheet(ss);

  setupDataValidation();
  setupConditionalFormatting();

  ui.alert('Workspace initialized successfully!');
}

// ============================================================================
// SHEET CREATION FUNCTIONS
// ============================================================================

/**
 * Create On Market Leads sheet (Zillow, Redfin, MLS, Crexi, LoopNet)
 */
function createOnMarketSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ON_MARKET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.ON_MARKET);
  }

  const headers = [
    'Lead ID',              // A - Auto-generated
    'Date Added',           // B - Auto-generated
    'Acquisition Team Member', // C - Dropdown
    'Source',               // D - Dropdown
    'MLS #',                // E
    'Property Address',     // F
    'City',                 // G
    'County',               // H
    'State',                // I
    'Zip Code',             // J
    'Property Type',        // K
    'Beds',                 // L
    'Baths',                // M
    'Sqft',                 // N
    'Year Built',           // O
    'List Price',           // P
    'Days on Market',       // Q
    'ARV Estimate',         // R
    'Repair Estimate',      // S
    'MAO',                  // T - Auto-calculated
    'Offer Amount',         // U
    'Offer Made (Yes/No)',  // V
    'Contract Signed (Yes/No)', // W - Triggers pipeline move
    'Notes',                // X
    'Agent Name',           // Y
    'Agent Phone',          // Z
    'Listing URL'           // AA
  ];

  setupSheetHeaders(sheet, headers, '#1E3A5F');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(6, 200);  // Address
  sheet.setColumnWidth(27, 250); // URL
}

/**
 * Create Off Market Leads sheet (Wholesalers, Direct Mail, etc.)
 */
function createOffMarketSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.OFF_MARKET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.OFF_MARKET);
  }

  const headers = [
    'Lead ID',              // A - Auto-generated
    'Date Added',           // B - Auto-generated
    'Acquisition Team Member', // C - Dropdown
    'Issue',                // D - Dropdown (renamed from Source)
    'Property Address',     // E
    'City',                 // F
    'County',               // G
    'State',                // H
    'Zip Code',             // I
    'Property Type',        // J
    'Beds',                 // K
    'Baths',                // L
    'Sqft',                 // M
    'Year Built',           // N
    'Asking Price',         // O
    'ARV Estimate',         // P
    'Repair Estimate',      // Q
    'MAO',                  // R - Auto-calculated
    'Offer Amount',         // S
    'Assignment Fee',       // T
    'Offer Made (Yes/No)',  // U
    'Contract Signed (Yes/No)', // V - Triggers pipeline move
    'Motivation Level',     // W
    'Seller Name',          // X
    'Seller Phone',         // Y
    'Notes'                 // Z
  ];

  setupSheetHeaders(sheet, headers, '#2D5016');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(5, 200); // Address
}

/**
 * Create Contract Pipeline sheet
 */
function createContractPipelineSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.CONTRACT_PIPELINE);
  }

  const headers = [
    'Deal ID',              // A - Auto-generated
    'Original Lead ID',     // B - From source sheet
    'Lead Type',            // C - On Market / Off Market
    'Acquisition Team Member', // D - From source
    'Date Under Contract',  // E - Auto-set
    'Property Address',     // F
    'City',                 // G
    'County',               // H
    'State',                // I
    'Zip Code',             // J
    'Purchase Price',       // K
    'ARV',                  // L
    'Repair Estimate',      // M
    'Expected Profit',      // N - Formula
    'Earnest Money',        // O
    'Inspection Period End', // P
    'Close Date',           // Q
    'Disposition Strategy', // R
    'Disposition Outcome',  // S
    'Buyer Name',           // T
    'Buyer Phone',          // U
    'Assignment Fee',       // V
    'Final Sale Price',     // W
    'Actual Profit',        // X - Formula
    'Status',               // Y
    'Title Company',        // Z
    'Title Contact',        // AA
    'Attorney',             // AB
    'Notes',                // AC
    'Days to Close'         // AD - Formula
  ];

  setupSheetHeaders(sheet, headers, '#5C1F1F');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(6, 200); // Address
}

/**
 * Create Monthly KPI sheet
 */
function createKPISheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.LEAD_KPIS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.LEAD_KPIS);
  }

  sheet.clear();

  // Get current month/year
  const now = new Date();
  const currentMonth = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM yyyy');

  // Title
  sheet.getRange('A1').setValue('MONTHLY KPI DASHBOARD').setFontSize(18).setFontWeight('bold');
  sheet.getRange('A1:F1').merge().setBackground('#1E3A5F').setFontColor('white').setHorizontalAlignment('center');

  // Current Period
  sheet.getRange('A3').setValue('Current Period:').setFontWeight('bold');
  sheet.getRange('B3').setValue(currentMonth).setFontSize(12);

  // KPI Headers
  sheet.getRange('A5:F5').setValues([['Metric', 'Goal', 'Actual', 'Variance', '% of Goal', 'Status']]);
  sheet.getRange('A5:F5').setFontWeight('bold').setBackground('#E0E0E0');

  // KPI Metrics (simplified - removed Qualified and Pass to Team Leader)
  const onMarket = CONFIG.SHEETS.ON_MARKET;
  const offMarket = CONFIG.SHEETS.OFF_MARKET;
  const pipeline = CONFIG.SHEETS.CONTRACT_PIPELINE;

  const kpiData = [
    ['Total Leads Generated', CONFIG.MONTHLY_GOALS.LEADS_GENERATED,
     `=COUNTIF('${onMarket}'!A:A,"<>")-1+COUNTIF('${offMarket}'!A:A,"<>")-1`, '', '', ''],
    ['On Market Leads', '', `=COUNTIF('${onMarket}'!A:A,"<>")-1`, '', '', ''],
    ['Off Market Leads', '', `=COUNTIF('${offMarket}'!A:A,"<>")-1`, '', '', ''],
    ['Offers Made', CONFIG.MONTHLY_GOALS.OFFERS_MADE,
     `=COUNTIF('${onMarket}'!V:V,"Yes")+COUNTIF('${offMarket}'!U:U,"Yes")`, '', '', ''],
    ['Contracts Signed', CONFIG.MONTHLY_GOALS.CONTRACTS_SIGNED,
     `=COUNTIF('${onMarket}'!W:W,"Yes")+COUNTIF('${offMarket}'!V:V,"Yes")`, '', '', ''],
    ['Deals in Pipeline', '', `=COUNTIF('${pipeline}'!Y:Y,"Under Contract")+COUNTIF('${pipeline}'!Y:Y,"Pending Close")`, '', '', ''],
    ['Deals Closed', CONFIG.MONTHLY_GOALS.DEALS_CLOSED, `=COUNTIF('${pipeline}'!Y:Y,"Closed")`, '', '', ''],
    ['Total Revenue', CONFIG.MONTHLY_GOALS.REVENUE_TARGET, `=SUMIF('${pipeline}'!Y:Y,"Closed",'${pipeline}'!X:X)`, '', '', '']
  ];

  sheet.getRange('A6:F13').setValues(kpiData);

  // Add formulas for variance, percentage, status
  for (let i = 6; i <= 13; i++) {
    sheet.getRange('D' + i).setFormula(`=IF(B${i}<>"",C${i}-B${i},"")`);
    sheet.getRange('E' + i).setFormula(`=IF(B${i}<>"",C${i}/B${i},"")`).setNumberFormat('0%');
    sheet.getRange('F' + i).setFormula(`=IF(E${i}="","",IF(E${i}>=1,"On Track",IF(E${i}>=0.75,"Behind","At Risk")))`);
  }

  // Format currency row
  sheet.getRange('C13').setNumberFormat('$#,##0');
  sheet.getRange('D13').setNumberFormat('$#,##0');

  // Conversion Rates Section
  sheet.getRange('A16').setValue('CONVERSION RATES').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A16:D16').merge().setBackground('#2D5016').setFontColor('white');

  sheet.getRange('A18:C18').setValues([['Conversion', 'Rate', 'Industry Avg']]);
  sheet.getRange('A18:C18').setFontWeight('bold').setBackground('#E0E0E0');

  const conversionData = [
    ['Lead to Offer', '=IF(C6>0,C9/C6,0)', '25%'],
    ['Offer to Contract', '=IF(C9>0,C10/C9,0)', '20%'],
    ['Contract to Close', '=IF(C10>0,C12/C10,0)', '80%'],
    ['Overall Lead to Close', '=IF(C6>0,C12/C6,0)', '3%']
  ];

  sheet.getRange('A19:C22').setValues(conversionData);
  sheet.getRange('B19:C22').setNumberFormat('0%');

  // Team Performance Section
  sheet.getRange('A25').setValue('TEAM PERFORMANCE').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A25:E25').merge().setBackground('#4A1F5C').setFontColor('white');

  sheet.getRange('A27:E27').setValues([['Team Member', 'Leads', 'Offers', 'Contracts', 'Conversion']]);
  sheet.getRange('A27:E27').setFontWeight('bold').setBackground('#E0E0E0');

  // Add team member rows
  CONFIG.TEAM_MEMBERS.forEach((member, idx) => {
    const row = 28 + idx;
    sheet.getRange('A' + row).setValue(member);
    sheet.getRange('B' + row).setFormula(
      `=COUNTIF('${onMarket}'!C:C,"${member}")+COUNTIF('${offMarket}'!C:C,"${member}")`
    );
    sheet.getRange('C' + row).setFormula(
      `=COUNTIFS('${onMarket}'!C:C,"${member}",'${onMarket}'!V:V,"Yes")+COUNTIFS('${offMarket}'!C:C,"${member}",'${offMarket}'!U:U,"Yes")`
    );
    sheet.getRange('D' + row).setFormula(
      `=COUNTIFS('${onMarket}'!C:C,"${member}",'${onMarket}'!W:W,"Yes")+COUNTIFS('${offMarket}'!C:C,"${member}",'${offMarket}'!V:V,"Yes")`
    );
    sheet.getRange('E' + row).setFormula(`=IF(B${row}>0,D${row}/B${row},0)`).setNumberFormat('0%');
  });

  // Source/Issue Performance
  sheet.getRange('A32').setValue('SOURCE PERFORMANCE').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A32:D32').merge().setBackground('#1F4A5C').setFontColor('white');

  sheet.getRange('A34:D34').setValues([['Source/Issue', 'Leads', 'Contracts', 'Rate']]);
  sheet.getRange('A34:D34').setFontWeight('bold').setBackground('#E0E0E0');

  let row = 35;
  // On Market Sources
  CONFIG.ON_MARKET_SOURCES.forEach(source => {
    sheet.getRange('A' + row).setValue(source + ' (On)');
    sheet.getRange('B' + row).setFormula(`=COUNTIF('${onMarket}'!D:D,"${source}")`);
    sheet.getRange('C' + row).setFormula(`=COUNTIFS('${onMarket}'!D:D,"${source}",'${onMarket}'!W:W,"Yes")`);
    sheet.getRange('D' + row).setFormula(`=IF(B${row}>0,C${row}/B${row},0)`).setNumberFormat('0%');
    row++;
  });

  // Off Market Issues
  CONFIG.OFF_MARKET_ISSUES.forEach(issue => {
    sheet.getRange('A' + row).setValue(issue + ' (Off)');
    sheet.getRange('B' + row).setFormula(`=COUNTIF('${offMarket}'!D:D,"${issue}")`);
    sheet.getRange('C' + row).setFormula(`=COUNTIFS('${offMarket}'!D:D,"${issue}",'${offMarket}'!V:V,"Yes")`);
    sheet.getRange('D' + row).setFormula(`=IF(B${row}>0,C${row}/B${row},0)`).setNumberFormat('0%');
    row++;
  });

  // Set column widths
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
}

/**
 * Create Yearly KPI Tracking sheet
 */
function createYearlyKPISheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.YEARLY_KPIS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.YEARLY_KPIS);
  }

  sheet.clear();

  const currentYear = new Date().getFullYear();

  // Title
  sheet.getRange('A1').setValue(`YEARLY KPI TRACKING - ${currentYear}`).setFontSize(18).setFontWeight('bold');
  sheet.getRange('A1:N1').merge().setBackground('#1E3A5F').setFontColor('white').setHorizontalAlignment('center');

  // Headers
  const headers = [
    'Metric', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'YTD Total'
  ];
  sheet.getRange('A3:N3').setValues([headers]);
  sheet.getRange('A3:N3').setFontWeight('bold').setBackground('#E0E0E0');

  // Metrics rows
  const metrics = [
    'Total Leads',
    'On Market Leads',
    'Off Market Leads',
    'Offers Made',
    'Contracts Signed',
    'Deals Closed',
    'Revenue',
    'Avg Deal Size',
    'Lead to Close %'
  ];

  metrics.forEach((metric, idx) => {
    const row = 4 + idx;
    sheet.getRange('A' + row).setValue(metric);

    // YTD Total formula (sum of months B through M)
    if (metric === 'Avg Deal Size') {
      sheet.getRange('N' + row).setFormula(`=IF(N9>0,N10/N9,0)`);
    } else if (metric === 'Lead to Close %') {
      sheet.getRange('N' + row).setFormula(`=IF(N4>0,N9/N4,0)`).setNumberFormat('0%');
    } else {
      sheet.getRange('N' + row).setFormula(`=SUM(B${row}:M${row})`);
    }
  });

  // Format revenue row
  sheet.getRange('B10:N10').setNumberFormat('$#,##0');
  sheet.getRange('B11:N11').setNumberFormat('$#,##0');

  // Goals section
  sheet.getRange('A15').setValue('MONTHLY GOALS').setFontWeight('bold');
  sheet.getRange('A15:B15').merge().setBackground('#2D5016').setFontColor('white');

  const goals = [
    ['Leads Generated', CONFIG.MONTHLY_GOALS.LEADS_GENERATED],
    ['Offers Made', CONFIG.MONTHLY_GOALS.OFFERS_MADE],
    ['Contracts Signed', CONFIG.MONTHLY_GOALS.CONTRACTS_SIGNED],
    ['Deals Closed', CONFIG.MONTHLY_GOALS.DEALS_CLOSED],
    ['Revenue Target', CONFIG.MONTHLY_GOALS.REVENUE_TARGET]
  ];

  sheet.getRange('A16:B20').setValues(goals);
  sheet.getRange('B20').setNumberFormat('$#,##0');

  // Instructions
  sheet.getRange('A23').setValue('Note: Use "Start New Month" from the menu to archive current month KPIs here.');
  sheet.getRange('A23').setFontStyle('italic').setFontColor('#666666');

  sheet.setColumnWidth(1, 150);
  for (let i = 2; i <= 14; i++) {
    sheet.setColumnWidth(i, 80);
  }
}

/**
 * Create Zip Codes sheet
 */
function createZipCodesSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ZIP_CODES);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.ZIP_CODES);
  }

  const headers = [
    'Zip Code', 'City', 'County', 'State', 'Market Status', 'Priority Level',
    'Avg Home Value', 'Avg Days on Market', 'Total Leads', 'Contracts Won',
    'Success Rate', 'Notes', 'Last Updated'
  ];

  setupSheetHeaders(sheet, headers, '#4A1F5C');
  sheet.setFrozenRows(1);
}

/**
 * Create Dashboard sheet
 */
function createDashboardSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DASHBOARD);
  }

  sheet.clear();

  const onMarket = CONFIG.SHEETS.ON_MARKET;
  const offMarket = CONFIG.SHEETS.OFF_MARKET;
  const pipeline = CONFIG.SHEETS.CONTRACT_PIPELINE;

  // Title
  sheet.getRange('A1').setValue('ACQUISITION DASHBOARD').setFontSize(24).setFontWeight('bold');
  sheet.getRange('A1:J1').merge().setBackground('#1E3A5F').setFontColor('white').setHorizontalAlignment('center');

  // Last Updated
  sheet.getRange('A2').setValue('Last Updated: ' + new Date().toLocaleString());
  sheet.getRange('A2:J2').merge().setHorizontalAlignment('center').setFontStyle('italic');

  // Quick Stats
  sheet.getRange('A4').setValue('QUICK STATS').setFontWeight('bold').setFontSize(14);

  const stats = [
    ['Total Active Leads', `=COUNTIF('${onMarket}'!A:A,"<>")-1+COUNTIF('${offMarket}'!A:A,"<>")-1`],
    ['Offers Pending', `=COUNTIF('${onMarket}'!V:V,"Yes")-COUNTIF('${onMarket}'!W:W,"Yes")+COUNTIF('${offMarket}'!U:U,"Yes")-COUNTIF('${offMarket}'!V:V,"Yes")`],
    ['Under Contract', `=COUNTIF('${pipeline}'!Y:Y,"Under Contract")`],
    ['Pending Close', `=COUNTIF('${pipeline}'!Y:Y,"Pending Close")`],
    ['Closed This Month', `=COUNTIF('${pipeline}'!Y:Y,"Closed")`],
    ['Revenue This Month', `=SUMIF('${pipeline}'!Y:Y,"Closed",'${pipeline}'!X:X)`]
  ];

  stats.forEach((stat, idx) => {
    const col = String.fromCharCode(65 + (idx % 3) * 3); // A, D, G
    const row = idx < 3 ? 5 : 8;
    sheet.getRange(col + row).setValue(stat[0]).setFontWeight('bold').setBackground('#E8E8E8');
    sheet.getRange(col + (row + 1)).setFormula(stat[1]).setFontSize(18).setHorizontalAlignment('center');
  });

  // Format revenue
  sheet.getRange('G9').setNumberFormat('$#,##0');

  // Team Leaderboard
  sheet.getRange('A12').setValue('TEAM LEADERBOARD').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A12:D12').merge().setBackground('#2D5016').setFontColor('white');

  sheet.getRange('A13:D13').setValues([['Team Member', 'Leads', 'Contracts', 'Conversion']]);
  sheet.getRange('A13:D13').setFontWeight('bold').setBackground('#E0E0E0');

  CONFIG.TEAM_MEMBERS.forEach((member, idx) => {
    const row = 14 + idx;
    sheet.getRange('A' + row).setValue(member);
    sheet.getRange('B' + row).setFormula(
      `=COUNTIF('${onMarket}'!C:C,"${member}")+COUNTIF('${offMarket}'!C:C,"${member}")`
    );
    sheet.getRange('C' + row).setFormula(
      `=COUNTIFS('${onMarket}'!C:C,"${member}",'${onMarket}'!W:W,"Yes")+COUNTIFS('${offMarket}'!C:C,"${member}",'${offMarket}'!V:V,"Yes")`
    );
    sheet.getRange('D' + row).setFormula(`=IF(B${row}>0,C${row}/B${row},0)`).setNumberFormat('0%');
  });

  // Monthly Goals Progress
  sheet.getRange('A18').setValue('MONTHLY GOALS PROGRESS').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A18:E18').merge().setBackground('#5C1F1F').setFontColor('white');

  sheet.getRange('A19:E19').setValues([['Metric', 'Goal', 'Actual', 'Progress', 'Status']]);
  sheet.getRange('A19:E19').setFontWeight('bold').setBackground('#E0E0E0');

  const goalData = [
    ['Leads Generated', CONFIG.MONTHLY_GOALS.LEADS_GENERATED,
     `=COUNTIF('${onMarket}'!A:A,"<>")-1+COUNTIF('${offMarket}'!A:A,"<>")-1`],
    ['Offers Made', CONFIG.MONTHLY_GOALS.OFFERS_MADE,
     `=COUNTIF('${onMarket}'!V:V,"Yes")+COUNTIF('${offMarket}'!U:U,"Yes")`],
    ['Contracts Signed', CONFIG.MONTHLY_GOALS.CONTRACTS_SIGNED,
     `=COUNTIF('${onMarket}'!W:W,"Yes")+COUNTIF('${offMarket}'!V:V,"Yes")`],
    ['Deals Closed', CONFIG.MONTHLY_GOALS.DEALS_CLOSED,
     `=COUNTIF('${pipeline}'!Y:Y,"Closed")`],
    ['Revenue', CONFIG.MONTHLY_GOALS.REVENUE_TARGET,
     `=SUMIF('${pipeline}'!Y:Y,"Closed",'${pipeline}'!X:X)`]
  ];

  goalData.forEach((data, idx) => {
    const row = 20 + idx;
    sheet.getRange('A' + row).setValue(data[0]);
    sheet.getRange('B' + row).setValue(data[1]);
    sheet.getRange('C' + row).setFormula(data[2]);
    sheet.getRange('D' + row).setFormula(`=IF(B${row}>0,C${row}/B${row},0)`).setNumberFormat('0%');
    sheet.getRange('E' + row).setFormula(
      `=IF(D${row}>=1,"Complete",IF(D${row}>=0.75,"On Track",IF(D${row}>=0.5,"Behind","At Risk")))`
    );
  });

  // Format revenue rows
  sheet.getRange('B24').setNumberFormat('$#,##0');
  sheet.getRange('C24').setNumberFormat('$#,##0');

  // Pipeline Summary
  sheet.getRange('G12').setValue('PIPELINE SUMMARY').setFontWeight('bold').setFontSize(14);
  sheet.getRange('G12:J12').merge().setBackground('#4A1F5C').setFontColor('white');

  sheet.getRange('G13:J13').setValues([['Status', 'Count', 'Total Value', 'Avg Value']]);
  sheet.getRange('G13:J13').setFontWeight('bold').setBackground('#E0E0E0');

  const pipelineStatuses = ['Under Contract', 'Pending Close', 'Closed'];
  pipelineStatuses.forEach((status, idx) => {
    const row = 14 + idx;
    sheet.getRange('G' + row).setValue(status);
    sheet.getRange('H' + row).setFormula(`=COUNTIF('${pipeline}'!Y:Y,"${status}")`);
    sheet.getRange('I' + row).setFormula(`=SUMIF('${pipeline}'!Y:Y,"${status}",'${pipeline}'!K:K)`).setNumberFormat('$#,##0');
    sheet.getRange('J' + row).setFormula(`=IF(H${row}>0,I${row}/H${row},0)`).setNumberFormat('$#,##0');
  });

  sheet.setColumnWidth(1, 150);
}

/**
 * Create Buyers List sheet
 */
function createBuyersListSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.BUYERS_LIST);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.BUYERS_LIST);
  }

  const headers = [
    'Buyer ID', 'Company Name', 'Contact Name', 'Phone', 'Email', 'Buyer Type',
    'Preferred Property Types', 'Preferred Zip Codes', 'Min Purchase Price', 'Max Purchase Price',
    'Proof of Funds', 'Deals Closed', 'Last Deal Date', 'Rating', 'Notes', 'Active (Yes/No)'
  ];

  setupSheetHeaders(sheet, headers, '#1F4A5C');
  sheet.setFrozenRows(1);
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Setup sheet headers with formatting
 */
function setupSheetHeaders(sheet, headers, bgColor) {
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground(bgColor);
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);
}

/**
 * Setup data validation dropdowns
 */
function setupDataValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // On Market validations
  const onMarket = ss.getSheetByName(CONFIG.SHEETS.ON_MARKET);
  if (onMarket) {
    // Team Member dropdown (Column C)
    const teamRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.TEAM_MEMBERS, true).build();
    onMarket.getRange('C2:C1000').setDataValidation(teamRule);

    // Source dropdown (Column D)
    const sourceRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.ON_MARKET_SOURCES, true).build();
    onMarket.getRange('D2:D1000').setDataValidation(sourceRule);

    // Property Type
    const propRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Single Family', 'Multi Family', 'Townhouse', 'Condo', 'Land', 'Commercial', 'Other'], true).build();
    onMarket.getRange('K2:K1000').setDataValidation(propRule);

    // Yes/No dropdowns
    const yesNoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true).build();
    onMarket.getRange('V2:W1000').setDataValidation(yesNoRule);
  }

  // Off Market validations
  const offMarket = ss.getSheetByName(CONFIG.SHEETS.OFF_MARKET);
  if (offMarket) {
    // Team Member dropdown (Column C)
    const teamRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.TEAM_MEMBERS, true).build();
    offMarket.getRange('C2:C1000').setDataValidation(teamRule);

    // Issue dropdown (Column D) - renamed from Source
    const issueRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.OFF_MARKET_ISSUES, true).build();
    offMarket.getRange('D2:D1000').setDataValidation(issueRule);

    // Property Type
    const propRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Single Family', 'Multi Family', 'Townhouse', 'Condo', 'Land', 'Commercial', 'Other'], true).build();
    offMarket.getRange('J2:J1000').setDataValidation(propRule);

    // Motivation Level
    const motivationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Hot', 'Warm', 'Cold'], true).build();
    offMarket.getRange('W2:W1000').setDataValidation(motivationRule);

    // Yes/No dropdowns
    const yesNoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true).build();
    offMarket.getRange('U2:V1000').setDataValidation(yesNoRule);
  }

  // Contract Pipeline validations
  const pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  if (pipeline) {
    // Disposition
    const dispRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.DISPOSITION, true).build();
    pipeline.getRange('R2:S1000').setDataValidation(dispRule);

    // Status
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Under Contract', 'Pending Close', 'Closed', 'Cancelled'], true).build();
    pipeline.getRange('Y2:Y1000').setDataValidation(statusRule);
  }

  // Zip Codes validations
  const zipCodes = ss.getSheetByName(CONFIG.SHEETS.ZIP_CODES);
  if (zipCodes) {
    const marketRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active', 'Paused', 'Retired'], true).build();
    zipCodes.getRange('E2:E1000').setDataValidation(marketRule);

    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['High', 'Medium', 'Low'], true).build();
    zipCodes.getRange('F2:F1000').setDataValidation(priorityRule);
  }

  // Buyers List validations
  const buyers = ss.getSheetByName(CONFIG.SHEETS.BUYERS_LIST);
  if (buyers) {
    const buyerTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Investor', 'Flipper', 'Landlord', 'Wholesaler', 'Other'], true).build();
    buyers.getRange('F2:F1000').setDataValidation(buyerTypeRule);

    const yesNoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true).build();
    buyers.getRange('K2:K1000').setDataValidation(yesNoRule);
    buyers.getRange('P2:P1000').setDataValidation(yesNoRule);

    const ratingRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['A', 'B', 'C', 'D'], true).build();
    buyers.getRange('N2:N1000').setDataValidation(ratingRule);
  }
}

/**
 * Setup conditional formatting
 */
function setupConditionalFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Contract Pipeline status colors
  const pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  if (pipeline) {
    const statusRange = pipeline.getRange('Y2:Y1000');
    const rules = [];

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Under Contract')
      .setBackground('#FFF3CD')
      .setRanges([statusRange]).build());

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Pending Close')
      .setBackground('#D1ECF1')
      .setRanges([statusRange]).build());

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Closed')
      .setBackground('#D4EDDA')
      .setRanges([statusRange]).build());

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Cancelled')
      .setBackground('#F8D7DA')
      .setRanges([statusRange]).build());

    pipeline.setConditionalFormatRules(rules);
  }
}

// ============================================================================
// AUTOMATION - ON EDIT TRIGGER
// ============================================================================

/**
 * Trigger on cell edit
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  if (row === 1) return; // Skip header

  // Auto-generate Lead ID and Date
  if (col > 2 && sheet.getRange(row, 1).getValue() === '') {
    if (sheetName === CONFIG.SHEETS.ON_MARKET) {
      const id = 'OM-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd') + '-' + row;
      sheet.getRange(row, 1).setValue(id);
      sheet.getRange(row, 2).setValue(new Date());
    } else if (sheetName === CONFIG.SHEETS.OFF_MARKET) {
      const id = 'OF-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd') + '-' + row;
      sheet.getRange(row, 1).setValue(id);
      sheet.getRange(row, 2).setValue(new Date());
    }
  }

  // Calculate MAO when ARV or Repairs change
  if (sheetName === CONFIG.SHEETS.ON_MARKET && (col === 18 || col === 19)) {
    // On Market: ARV=R(18), Repairs=S(19), MAO=T(20)
    const arv = sheet.getRange(row, 18).getValue();
    const repairs = sheet.getRange(row, 19).getValue();
    if (arv && repairs) {
      const mao = (arv * 0.7) - repairs;
      sheet.getRange(row, 20).setValue(Math.max(0, mao)).setNumberFormat('$#,##0');
    }
  } else if (sheetName === CONFIG.SHEETS.OFF_MARKET && (col === 16 || col === 17)) {
    // Off Market: ARV=P(16), Repairs=Q(17), MAO=R(18)
    const arv = sheet.getRange(row, 16).getValue();
    const repairs = sheet.getRange(row, 17).getValue();
    if (arv && repairs) {
      const mao = (arv * 0.7) - repairs;
      sheet.getRange(row, 18).setValue(Math.max(0, mao)).setNumberFormat('$#,##0');
    }
  }

  // Auto-move to Contract Pipeline when Contract Signed = Yes
  // On Market: Contract Signed = Column W (23)
  // Off Market: Contract Signed = Column V (22)
  if (sheetName === CONFIG.SHEETS.ON_MARKET && col === 23 && e.value === 'Yes') {
    moveToContractPipeline(sheet, row, 'On Market');
  } else if (sheetName === CONFIG.SHEETS.OFF_MARKET && col === 22 && e.value === 'Yes') {
    moveToContractPipeline(sheet, row, 'Off Market');
  }
}

/**
 * Move lead to Contract Pipeline
 */
function moveToContractPipeline(sourceSheet, row, leadType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  if (!pipeline) return;

  const rowData = sourceSheet.getRange(row, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const dealId = 'CP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmm');

  let pipelineRow;

  if (leadType === 'On Market') {
    // On Market columns: LeadID(0), Date(1), TeamMember(2), Source(3), MLS(4), Address(5), City(6),
    // County(7), State(8), Zip(9), PropType(10), Beds(11), Baths(12), Sqft(13), YearBuilt(14),
    // ListPrice(15), DOM(16), ARV(17), Repairs(18), MAO(19), OfferAmt(20), OfferMade(21),
    // ContractSigned(22), Notes(23), AgentName(24), AgentPhone(25), URL(26)
    pipelineRow = [
      dealId,           // Deal ID
      rowData[0],       // Original Lead ID
      leadType,         // Lead Type
      rowData[2],       // Acquisition Team Member
      new Date(),       // Date Under Contract
      rowData[5],       // Property Address
      rowData[6],       // City
      rowData[7],       // County
      rowData[8],       // State
      rowData[9],       // Zip Code
      rowData[20] || rowData[15], // Purchase Price (Offer or List)
      rowData[17],      // ARV
      rowData[18],      // Repair Estimate
      '',               // Expected Profit
      '',               // Earnest Money
      '',               // Inspection Period End
      '',               // Close Date
      '',               // Disposition Strategy
      '',               // Disposition Outcome
      '',               // Buyer Name
      '',               // Buyer Phone
      '',               // Assignment Fee
      '',               // Final Sale Price
      '',               // Actual Profit
      'Under Contract', // Status
      '',               // Title Company
      '',               // Title Contact
      '',               // Attorney
      rowData[23],      // Notes
      ''                // Days to Close
    ];
  } else {
    // Off Market columns: LeadID(0), Date(1), TeamMember(2), Issue(3), Address(4), City(5),
    // County(6), State(7), Zip(8), PropType(9), Beds(10), Baths(11), Sqft(12), YearBuilt(13),
    // AskingPrice(14), ARV(15), Repairs(16), MAO(17), OfferAmt(18), AssignmentFee(19),
    // OfferMade(20), ContractSigned(21), Motivation(22), SellerName(23), SellerPhone(24), Notes(25)
    pipelineRow = [
      dealId,           // Deal ID
      rowData[0],       // Original Lead ID
      leadType,         // Lead Type
      rowData[2],       // Acquisition Team Member
      new Date(),       // Date Under Contract
      rowData[4],       // Property Address
      rowData[5],       // City
      rowData[6],       // County
      rowData[7],       // State
      rowData[8],       // Zip Code
      rowData[18] || rowData[14], // Purchase Price (Offer or Asking)
      rowData[15],      // ARV
      rowData[16],      // Repair Estimate
      '',               // Expected Profit
      '',               // Earnest Money
      '',               // Inspection Period End
      '',               // Close Date
      '',               // Disposition Strategy
      '',               // Disposition Outcome
      '',               // Buyer Name
      '',               // Buyer Phone
      rowData[19],      // Assignment Fee
      '',               // Final Sale Price
      '',               // Actual Profit
      'Under Contract', // Status
      '',               // Title Company
      '',               // Title Contact
      '',               // Attorney
      rowData[25],      // Notes
      ''                // Days to Close
    ];
  }

  pipeline.appendRow(pipelineRow);
}

// ============================================================================
// MONTHLY KPI FUNCTIONS
// ============================================================================

/**
 * Start a new month - archives current KPIs to Yearly sheet
 */
function startNewMonth() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Start New Month',
    'This will archive current month KPIs to the Yearly sheet and reset for a new month. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  archiveMonthlyKPIs();
  updateKPIs();

  ui.alert('New month started! Previous KPIs archived to Yearly KPIs sheet.');
}

/**
 * Archive current month's KPIs to yearly tracking
 */
function archiveMonthlyKPIs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kpiSheet = ss.getSheetByName(CONFIG.SHEETS.LEAD_KPIS);
  const yearlySheet = ss.getSheetByName(CONFIG.SHEETS.YEARLY_KPIS);

  if (!kpiSheet || !yearlySheet) return;

  // Get current month
  const now = new Date();
  const monthIndex = now.getMonth(); // 0-11
  const monthCol = monthIndex + 2;   // Column B=Jan(2), C=Feb(3), etc.

  // Get KPI values from the current month KPI sheet
  const totalLeads = kpiSheet.getRange('C6').getValue();
  const onMarketLeads = kpiSheet.getRange('C7').getValue();
  const offMarketLeads = kpiSheet.getRange('C8').getValue();
  const offersMade = kpiSheet.getRange('C9').getValue();
  const contractsSigned = kpiSheet.getRange('C10').getValue();
  const dealsClosed = kpiSheet.getRange('C12').getValue();
  const revenue = kpiSheet.getRange('C13').getValue();

  // Write to yearly sheet
  yearlySheet.getRange(4, monthCol).setValue(totalLeads);      // Total Leads
  yearlySheet.getRange(5, monthCol).setValue(onMarketLeads);   // On Market
  yearlySheet.getRange(6, monthCol).setValue(offMarketLeads);  // Off Market
  yearlySheet.getRange(7, monthCol).setValue(offersMade);      // Offers Made
  yearlySheet.getRange(8, monthCol).setValue(contractsSigned); // Contracts Signed
  yearlySheet.getRange(9, monthCol).setValue(dealsClosed);     // Deals Closed
  yearlySheet.getRange(10, monthCol).setValue(revenue);        // Revenue

  // Calculate and set Avg Deal Size
  if (dealsClosed > 0) {
    yearlySheet.getRange(11, monthCol).setValue(revenue / dealsClosed);
  }

  // Calculate Lead to Close %
  if (totalLeads > 0) {
    yearlySheet.getRange(12, monthCol).setValue(dealsClosed / totalLeads);
  }
}

/**
 * Update KPI sheet with current month
 */
function updateKPIs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createKPISheet(ss);
  SpreadsheetApp.getUi().alert('KPIs updated!');
}

/**
 * Refresh dashboard
 */
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);

  if (dashboard) {
    dashboard.getRange('A2').setValue('Last Updated: ' + new Date().toLocaleString());
    SpreadsheetApp.flush();
  }

  SpreadsheetApp.getUi().alert('Dashboard refreshed!');
}

// ============================================================================
// BUYERS IMPORT FROM DISPO WORKSPACE
// ============================================================================

/**
 * Import buyers from external Dispo Workspace sheet
 */
function importBuyersFromDispo() {
  const ui = SpreadsheetApp.getUi();

  // Check if spreadsheet ID is configured
  if (!CONFIG.BUYERS_IMPORT.SPREADSHEET_ID) {
    const response = ui.prompt(
      'Import Buyers',
      'Enter the Spreadsheet ID of your Dispo Workspace 3.0 (found in the URL):',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) return;

    const spreadsheetId = response.getResponseText().trim();
    if (!spreadsheetId) {
      ui.alert('No spreadsheet ID provided.');
      return;
    }

    importBuyersFromSheet(spreadsheetId);
  } else {
    importBuyersFromSheet(CONFIG.BUYERS_IMPORT.SPREADSHEET_ID);
  }
}

/**
 * Import buyers from specified spreadsheet
 */
function importBuyersFromSheet(spreadsheetId) {
  const ui = SpreadsheetApp.getUi();

  try {
    const sourceSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sourceSheet = sourceSpreadsheet.getSheetByName(CONFIG.BUYERS_IMPORT.SHEET_NAME);

    if (!sourceSheet) {
      ui.alert('Could not find "Buyer" sheet in the specified spreadsheet.');
      return;
    }

    const sourceData = sourceSheet.getDataRange().getValues();
    if (sourceData.length < 2) {
      ui.alert('No buyer data found in source sheet.');
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const buyerSheet = ss.getSheetByName(CONFIG.SHEETS.BUYERS_LIST);

    if (!buyerSheet) {
      ui.alert('Buyers List sheet not found. Please initialize workspace first.');
      return;
    }

    // Clear existing data (keep headers)
    const lastRow = buyerSheet.getLastRow();
    if (lastRow > 1) {
      buyerSheet.getRange(2, 1, lastRow - 1, buyerSheet.getLastColumn()).clear();
    }

    // Copy data (skip header row from source)
    const dataToImport = sourceData.slice(1);
    if (dataToImport.length > 0) {
      // Map source columns to destination columns (adjust based on your dispo workspace structure)
      // This assumes similar column structure - adjust mapping as needed
      buyerSheet.getRange(2, 1, dataToImport.length, dataToImport[0].length).setValues(dataToImport);
    }

    ui.alert(`Successfully imported ${dataToImport.length} buyers from Dispo Workspace!`);

  } catch (error) {
    ui.alert('Error importing buyers: ' + error.message +
             '\n\nMake sure you have access to the source spreadsheet.');
  }
}

// ============================================================================
// FORM DIALOGS
// ============================================================================

/**
 * Show On Market Lead form
 */
function showOnMarketForm() {
  const teamOptions = CONFIG.TEAM_MEMBERS.map(m => `<option value="${m}">${m}</option>`).join('');
  const sourceOptions = CONFIG.ON_MARKET_SOURCES.map(s => `<option value="${s}">${s}</option>`).join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      .form-group { margin-bottom: 12px; }
      label { display: block; font-weight: bold; margin-bottom: 4px; }
      input, select { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
      .btn { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px; }
      .btn-primary { background: #1E3A5F; color: white; }
      .btn-secondary { background: #ccc; }
      .row { display: flex; gap: 10px; }
      .col { flex: 1; }
    </style>
    <h3>Add On Market Lead</h3>
    <form id="leadForm">
      <div class="row">
        <div class="col form-group">
          <label>Team Member *</label>
          <select id="teamMember" required>${teamOptions}</select>
        </div>
        <div class="col form-group">
          <label>Source *</label>
          <select id="source" required>${sourceOptions}</select>
        </div>
      </div>
      <div class="form-group">
        <label>MLS #</label>
        <input type="text" id="mls">
      </div>
      <div class="form-group">
        <label>Property Address *</label>
        <input type="text" id="address" required>
      </div>
      <div class="row">
        <div class="col form-group">
          <label>City *</label>
          <input type="text" id="city" required>
        </div>
        <div class="col form-group">
          <label>State</label>
          <input type="text" id="state" value="TX">
        </div>
        <div class="col form-group">
          <label>Zip Code *</label>
          <input type="text" id="zip" required>
        </div>
      </div>
      <div class="row">
        <div class="col form-group">
          <label>List Price *</label>
          <input type="number" id="price" required>
        </div>
        <div class="col form-group">
          <label>ARV Estimate</label>
          <input type="number" id="arv">
        </div>
      </div>
      <div class="form-group">
        <label>Listing URL</label>
        <input type="url" id="url">
      </div>
      <div class="form-group">
        <label>Notes</label>
        <textarea id="notes" style="width:100%;height:60px;"></textarea>
      </div>
      <button type="button" class="btn btn-primary" onclick="submitForm()">Add Lead</button>
      <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
    </form>
    <script>
      function submitForm() {
        const data = {
          teamMember: document.getElementById('teamMember').value,
          source: document.getElementById('source').value,
          mls: document.getElementById('mls').value,
          address: document.getElementById('address').value,
          city: document.getElementById('city').value,
          state: document.getElementById('state').value,
          zip: document.getElementById('zip').value,
          price: document.getElementById('price').value,
          arv: document.getElementById('arv').value,
          url: document.getElementById('url').value,
          notes: document.getElementById('notes').value
        };
        google.script.run
          .withSuccessHandler(() => { alert('Lead added!'); google.script.host.close(); })
          .withFailureHandler(err => alert('Error: ' + err.message))
          .addOnMarketLead(data);
      }
    </script>
  `).setWidth(450).setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, 'Add On Market Lead');
}

/**
 * Show Off Market Lead form
 */
function showOffMarketForm() {
  const teamOptions = CONFIG.TEAM_MEMBERS.map(m => `<option value="${m}">${m}</option>`).join('');
  const issueOptions = CONFIG.OFF_MARKET_ISSUES.map(i => `<option value="${i}">${i}</option>`).join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      .form-group { margin-bottom: 12px; }
      label { display: block; font-weight: bold; margin-bottom: 4px; }
      input, select { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
      .btn { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px; }
      .btn-primary { background: #2D5016; color: white; }
      .btn-secondary { background: #ccc; }
      .row { display: flex; gap: 10px; }
      .col { flex: 1; }
    </style>
    <h3>Add Off Market Lead</h3>
    <form id="leadForm">
      <div class="row">
        <div class="col form-group">
          <label>Team Member *</label>
          <select id="teamMember" required>${teamOptions}</select>
        </div>
        <div class="col form-group">
          <label>Issue *</label>
          <select id="issue" required>${issueOptions}</select>
        </div>
      </div>
      <div class="row">
        <div class="col form-group">
          <label>Motivation Level</label>
          <select id="motivation">
            <option value="">Select</option>
            <option value="Hot">Hot</option>
            <option value="Warm">Warm</option>
            <option value="Cold">Cold</option>
          </select>
        </div>
      </div>
      <div class="form-group">
        <label>Property Address *</label>
        <input type="text" id="address" required>
      </div>
      <div class="row">
        <div class="col form-group">
          <label>City *</label>
          <input type="text" id="city" required>
        </div>
        <div class="col form-group">
          <label>State</label>
          <input type="text" id="state" value="TX">
        </div>
        <div class="col form-group">
          <label>Zip Code *</label>
          <input type="text" id="zip" required>
        </div>
      </div>
      <div class="row">
        <div class="col form-group">
          <label>Asking Price *</label>
          <input type="number" id="price" required>
        </div>
        <div class="col form-group">
          <label>ARV Estimate</label>
          <input type="number" id="arv">
        </div>
      </div>
      <div class="row">
        <div class="col form-group">
          <label>Seller Name</label>
          <input type="text" id="sellerName">
        </div>
        <div class="col form-group">
          <label>Seller Phone</label>
          <input type="tel" id="sellerPhone">
        </div>
      </div>
      <div class="form-group">
        <label>Notes</label>
        <textarea id="notes" style="width:100%;height:60px;"></textarea>
      </div>
      <button type="button" class="btn btn-primary" onclick="submitForm()">Add Lead</button>
      <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
    </form>
    <script>
      function submitForm() {
        const data = {
          teamMember: document.getElementById('teamMember').value,
          issue: document.getElementById('issue').value,
          motivation: document.getElementById('motivation').value,
          address: document.getElementById('address').value,
          city: document.getElementById('city').value,
          state: document.getElementById('state').value,
          zip: document.getElementById('zip').value,
          price: document.getElementById('price').value,
          arv: document.getElementById('arv').value,
          sellerName: document.getElementById('sellerName').value,
          sellerPhone: document.getElementById('sellerPhone').value,
          notes: document.getElementById('notes').value
        };
        google.script.run
          .withSuccessHandler(() => { alert('Lead added!'); google.script.host.close(); })
          .withFailureHandler(err => alert('Error: ' + err.message))
          .addOffMarketLead(data);
      }
    </script>
  `).setWidth(450).setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Add Off Market Lead');
}

/**
 * Add On Market Lead from form
 */
function addOnMarketLead(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ON_MARKET);
  if (!sheet) throw new Error('On Market Leads sheet not found.');

  const id = 'OM-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmm');

  const rowData = [
    id, new Date(), data.teamMember, data.source, data.mls, data.address,
    data.city, '', data.state, data.zip, '', '', '', '', '', data.price, '',
    data.arv, '', '', '', 'No', 'No', data.notes, '', '', data.url
  ];

  sheet.appendRow(rowData);
}

/**
 * Add Off Market Lead from form
 */
function addOffMarketLead(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.OFF_MARKET);
  if (!sheet) throw new Error('Off Market Leads sheet not found.');

  const id = 'OF-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmm');

  const rowData = [
    id, new Date(), data.teamMember, data.issue, data.address, data.city,
    '', data.state, data.zip, '', '', '', '', '', data.price, data.arv, '',
    '', '', '', 'No', 'No', data.motivation, data.sellerName, data.sellerPhone, data.notes
  ];

  sheet.appendRow(rowData);
}

// ============================================================================
// ZIP CODE IMPORT
// ============================================================================

/**
 * Show zip code import dialog
 */
function showZipCodeImport() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      textarea { width: 100%; height: 200px; }
      .btn { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-top: 10px; }
      .btn-primary { background: #4A1F5C; color: white; }
    </style>
    <h3>Import Zip Codes</h3>
    <p>Enter zip codes (one per line or comma-separated):</p>
    <textarea id="zips" placeholder="75001&#10;75002&#10;75003"></textarea>
    <br>
    <button class="btn btn-primary" onclick="importZips()">Import</button>
    <script>
      function importZips() {
        const zips = document.getElementById('zips').value;
        google.script.run
          .withSuccessHandler(count => { alert(count + ' zip codes imported!'); google.script.host.close(); })
          .importZipCodes(zips);
      }
    </script>
  `).setWidth(400).setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(html, 'Import Zip Codes');
}

/**
 * Import zip codes
 */
function importZipCodes(zipText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ZIP_CODES);
  if (!sheet) throw new Error('Zip Codes sheet not found');

  const zips = zipText.split(/[\n,]+/).map(z => z.trim()).filter(z => z && /^\d{5}$/.test(z));
  const existingZips = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat().filter(z => z);

  let added = 0;
  zips.forEach(zip => {
    if (!existingZips.includes(zip)) {
      sheet.appendRow([zip, '', '', '', 'Active', 'Medium', '', '', '', '', '', '', new Date()]);
      added++;
    }
  });

  return added;
}

// ============================================================================
// HELP
// ============================================================================

function showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; line-height: 1.6; }
      h2 { color: #1E3A5F; }
      h3 { color: #2D5016; margin-top: 20px; }
      ul { padding-left: 20px; }
    </style>
    <h2>Acquisition Workspace v4.0</h2>
    <h3>Sheets</h3>
    <ul>
      <li><b>On Market Leads</b> - Zillow, Redfin, MLS, Crexi, LoopNet</li>
      <li><b>Off Market Leads</b> - Wholesalers, Direct Mail, Expired, Foreclosure, etc.</li>
      <li><b>Contract Pipeline</b> - Auto-populated when Contract Signed = Yes</li>
      <li><b>Lead KPIs</b> - Monthly metrics</li>
      <li><b>Yearly KPIs</b> - Track all months</li>
      <li><b>Zip Codes</b> - Target markets</li>
      <li><b>Dashboard</b> - Overview</li>
      <li><b>Buyers List</b> - Cash buyers (import from Dispo Workspace)</li>
    </ul>
    <h3>Automation</h3>
    <ul>
      <li>Lead ID & Date auto-generated</li>
      <li>MAO auto-calculated (70% rule)</li>
      <li>Contract Signed = Yes -> auto-moves to Pipeline</li>
      <li>"Start New Month" archives KPIs</li>
    </ul>
    <h3>Team Members</h3>
    <ul>
      <li>Tess Walter</li>
      <li>Nick Barr</li>
    </ul>
  `).setWidth(450).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Help');
}
