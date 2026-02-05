/**
 * Keystone Acquisition Workspace - Google Apps Script
 * Real Estate Lead Tracking & Automation System
 *
 * This script automates lead tracking, KPI calculations, and data management
 * for real estate acquisition operations.
 */

// ============================================================================
// CONFIGURATION - Customize these values for your workspace
// ============================================================================

const CONFIG = {
  // Sheet Names
  SHEETS: {
    ON_MARKET: 'On Market Leads',
    OFF_MARKET: 'Off Market Leads',
    CONTRACT_PIPELINE: 'Contract Pipeline',
    MAIN_PAGE: 'Main Page',
    LEAD_KPIS: 'Lead KPIs',
    ZIP_CODES: 'Zip Codes',
    DASHBOARD: 'Dashboard',
    BUYERS_LIST: 'Buyers List'
  },

  // Lead Sources
  SOURCES: {
    ON_MARKET: ['Zillow', 'Redfin', 'Realtor.com', 'MLS'],
    OFF_MARKET: ['Wholesaler', 'Agent Referral', 'Direct Mail', 'Cold Call', 'Driving for Dollars', 'Probate', 'Tax Lien', 'Other']
  },

  // Status Options
  STATUS: {
    LEAD: ['New', 'Contacted', 'Analyzing', 'Offer Pending', 'Under Contract', 'Closed', 'Dead'],
    DISPOSITION: ['Assigned', 'Double Close', 'Wholetail', 'Fix & Flip', 'Buy & Hold', 'Dead Deal']
  },

  // Monthly Goals (customize these)
  MONTHLY_GOALS: {
    LEADS_GENERATED: 100,
    LEADS_QUALIFIED: 50,
    OFFERS_MADE: 25,
    CONTRACTS_SIGNED: 5,
    DEALS_CLOSED: 3,
    REVENUE_TARGET: 50000
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
  ui.createMenu('🏠 Acquisition Tools')
    .addItem('📊 Refresh Dashboard', 'refreshDashboard')
    .addItem('📈 Update KPIs', 'updateKPIs')
    .addSeparator()
    .addSubMenu(ui.createMenu('➕ Add New Lead')
      .addItem('On Market Lead', 'showOnMarketForm')
      .addItem('Off Market Lead', 'showOffMarketForm'))
    .addItem('📋 Move to Contract Pipeline', 'moveToContractPipeline')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Setup')
      .addItem('Initialize All Sheets', 'initializeWorkspace')
      .addItem('Reset Monthly KPIs', 'resetMonthlyKPIs')
      .addItem('Import Zip Codes', 'showZipCodeImport'))
    .addSeparator()
    .addItem('❓ Help', 'showHelp')
    .addToUi();
}

/**
 * Initialize the entire workspace with all required sheets
 */
function initializeWorkspace() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Initialize Workspace',
    'This will create/update all sheets with the proper structure. Existing data will be preserved. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  createOnMarketSheet(ss);
  createOffMarketSheet(ss);
  createContractPipelineSheet(ss);
  createMainPageSheet(ss);
  createKPISheet(ss);
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
 * Create or update On Market Leads sheet (Zillow, Redfin, etc.)
 */
function createOnMarketSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ON_MARKET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.ON_MARKET);
  }

  const headers = [
    'Lead ID', 'Date Added', 'Source', 'MLS #', 'Property Address', 'City',
    'County', 'State', 'Zip Code', 'Property Type', 'Beds', 'Baths', 'Sqft',
    'Year Built', 'List Price', 'Days on Market', 'ARV Estimate', 'Repair Estimate',
    'MAO', 'Offer Amount', 'Qualified (Yes/No)', 'Passed to Team Lead (Yes/No)',
    'Offer Made (Yes/No)', 'Contract Signed (Yes/No)', 'Status', 'Notes',
    'Agent Name', 'Agent Phone', 'Listing URL'
  ];

  setupSheetHeaders(sheet, headers, '#1E3A5F');
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 80);   // Lead ID
  sheet.setColumnWidth(2, 100);  // Date Added
  sheet.setColumnWidth(5, 200);  // Address
  sheet.setColumnWidth(29, 250); // URL
}

/**
 * Create or update Off Market Leads sheet (Wholesalers, Agents, etc.)
 */
function createOffMarketSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.OFF_MARKET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.OFF_MARKET);
  }

  const headers = [
    'Lead ID', 'Date Added', 'Source', 'Source Contact', 'Source Phone', 'Source Email',
    'Property Address', 'City', 'County', 'State', 'Zip Code', 'Property Type',
    'Beds', 'Baths', 'Sqft', 'Year Built', 'Asking Price', 'ARV Estimate',
    'Repair Estimate', 'MAO', 'Offer Amount', 'Assignment Fee', 'Qualified (Yes/No)',
    'Passed to Team Lead (Yes/No)', 'Offer Made (Yes/No)', 'Contract Signed (Yes/No)',
    'Status', 'Motivation Level', 'Seller Name', 'Seller Phone', 'Notes'
  ];

  setupSheetHeaders(sheet, headers, '#2D5016');
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(7, 200);
}

/**
 * Create or update Contract Pipeline sheet
 */
function createContractPipelineSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.CONTRACT_PIPELINE);
  }

  const headers = [
    'Deal ID', 'Original Lead ID', 'Date Under Contract', 'Property Address', 'City',
    'County', 'State', 'Zip Code', 'Purchase Price', 'ARV', 'Repair Estimate',
    'Expected Profit', 'Earnest Money', 'Inspection Period End', 'Close Date',
    'Disposition Strategy', 'Disposition Outcome', 'Buyer Name', 'Buyer Phone',
    'Assignment Fee', 'Final Sale Price', 'Actual Profit', 'Status',
    'Title Company', 'Title Contact', 'Attorney', 'Notes', 'Days to Close'
  ];

  setupSheetHeaders(sheet, headers, '#5C1F1F');
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(4, 200);
}

/**
 * Create or update Main Page (consolidated view)
 */
function createMainPageSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.MAIN_PAGE);
  }

  const headers = [
    'Lead ID', 'Date Added', 'Lead Type', 'Source', 'Property Address', 'City',
    'County', 'Zip Code', 'Property Type', 'List/Ask Price', 'ARV', 'MAO',
    'Qualified (Yes/No)', 'Passed to Team Lead (Yes/No)', 'Offer Made (Yes/No)',
    'Contract Signed (Yes/No)', 'Disposition Outcome', 'Buyer Name', 'Final Sale Price',
    'Status', 'Notes'
  ];

  setupSheetHeaders(sheet, headers, '#1A4D4D');
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(5, 200);
}

/**
 * Create or update Lead KPIs sheet
 */
function createKPISheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.LEAD_KPIS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.LEAD_KPIS);
  }

  // Clear existing content
  sheet.clear();

  // Title
  sheet.getRange('A1').setValue('MONTHLY KPI DASHBOARD').setFontSize(18).setFontWeight('bold');
  sheet.getRange('A1:H1').merge().setBackground('#1E3A5F').setFontColor('white').setHorizontalAlignment('center');

  // Current Month Header
  const currentMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy');
  sheet.getRange('A3').setValue('Current Period: ' + currentMonth).setFontSize(12).setFontWeight('bold');

  // KPI Headers
  const kpiHeaders = ['Metric', 'Goal', 'Actual', 'Variance', '% of Goal', 'Status'];
  sheet.getRange('A5:F5').setValues([kpiHeaders]).setFontWeight('bold').setBackground('#E0E0E0');

  // KPI Rows
  const kpiData = [
    ['Total Leads Generated', CONFIG.MONTHLY_GOALS.LEADS_GENERATED, '=COUNTIF(\'Main Page\'!A:A,"<>")-1', '', '', ''],
    ['Leads Qualified', CONFIG.MONTHLY_GOALS.LEADS_QUALIFIED, '=COUNTIF(\'Main Page\'!M:M,"Yes")', '', '', ''],
    ['Passed to Team Lead', '', '=COUNTIF(\'Main Page\'!N:N,"Yes")', '', '', ''],
    ['Offers Made', CONFIG.MONTHLY_GOALS.OFFERS_MADE, '=COUNTIF(\'Main Page\'!O:O,"Yes")', '', '', ''],
    ['Contracts Signed', CONFIG.MONTHLY_GOALS.CONTRACTS_SIGNED, '=COUNTIF(\'Main Page\'!P:P,"Yes")', '', '', ''],
    ['Deals Closed', CONFIG.MONTHLY_GOALS.DEALS_CLOSED, '=COUNTIF(\'Contract Pipeline\'!W:W,"Closed")', '', '', ''],
    ['Total Revenue', CONFIG.MONTHLY_GOALS.REVENUE_TARGET, '=SUMIF(\'Contract Pipeline\'!W:W,"Closed",\'Contract Pipeline\'!V:V)', '', '', '']
  ];

  sheet.getRange('A6:F12').setValues(kpiData);

  // Add variance and percentage formulas
  for (let i = 6; i <= 12; i++) {
    // Variance formula
    sheet.getRange('D' + i).setFormula('=IF(B' + i + '<>"",C' + i + '-B' + i + ',"")');
    // Percentage formula
    sheet.getRange('E' + i).setFormula('=IF(B' + i + '<>"",C' + i + '/B' + i + ',"")').setNumberFormat('0%');
    // Status formula
    sheet.getRange('F' + i).setFormula('=IF(E' + i + '="","",IF(E' + i + '>=1,"✅ On Track",IF(E' + i + '>=0.75,"⚠️ Behind","❌ At Risk")))');
  }

  // Conversion Rates Section
  sheet.getRange('A15').setValue('CONVERSION RATES').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A15:D15').merge().setBackground('#2D5016').setFontColor('white');

  const conversionHeaders = ['Conversion Type', 'Formula', 'Rate', 'Industry Avg'];
  sheet.getRange('A17:D17').setValues([conversionHeaders]).setFontWeight('bold').setBackground('#E0E0E0');

  const conversionData = [
    ['Lead to Qualified', 'Qualified / Total Leads', '=IF(C6>0,C7/C6,0)', '50%'],
    ['Qualified to Offer', 'Offers / Qualified', '=IF(C7>0,C9/C7,0)', '50%'],
    ['Offer to Contract', 'Contracts / Offers', '=IF(C9>0,C10/C9,0)', '20%'],
    ['Contract to Close', 'Closed / Contracts', '=IF(C10>0,C11/C10,0)', '80%'],
    ['Overall Lead to Close', 'Closed / Total Leads', '=IF(C6>0,C11/C6,0)', '3%']
  ];

  sheet.getRange('A18:D22').setValues(conversionData);
  sheet.getRange('C18:D22').setNumberFormat('0%');

  // Source Performance Section
  sheet.getRange('A25').setValue('LEAD SOURCE PERFORMANCE').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A25:E25').merge().setBackground('#5C1F1F').setFontColor('white');

  const sourceHeaders = ['Source', 'Total Leads', 'Qualified', 'Contracts', 'Conversion Rate'];
  sheet.getRange('A27:E27').setValues([sourceHeaders]).setFontWeight('bold').setBackground('#E0E0E0');

  // Add COUNTIF formulas for each source
  const allSources = [...CONFIG.SOURCES.ON_MARKET, ...CONFIG.SOURCES.OFF_MARKET];
  let row = 28;
  allSources.forEach(source => {
    sheet.getRange('A' + row).setValue(source);
    sheet.getRange('B' + row).setFormula('=COUNTIF(\'Main Page\'!D:D,"' + source + '")');
    sheet.getRange('C' + row).setFormula('=COUNTIFS(\'Main Page\'!D:D,"' + source + '",\'Main Page\'!M:M,"Yes")');
    sheet.getRange('D' + row).setFormula('=COUNTIFS(\'Main Page\'!D:D,"' + source + '",\'Main Page\'!P:P,"Yes")');
    sheet.getRange('E' + row).setFormula('=IF(B' + row + '>0,D' + row + '/B' + row + ',0)').setNumberFormat('0%');
    row++;
  });

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 120);
}

/**
 * Create or update Zip Codes sheet
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

  // Add formulas for leads count and success rate in row 2 as example
  sheet.getRange('I2').setFormula('=COUNTIF(\'Main Page\'!H:H,A2)');
  sheet.getRange('J2').setFormula('=COUNTIFS(\'Main Page\'!H:H,A2,\'Main Page\'!P:P,"Yes")');
  sheet.getRange('K2').setFormula('=IF(I2>0,J2/I2,0)').setNumberFormat('0%');
}

/**
 * Create or update Dashboard sheet
 */
function createDashboardSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DASHBOARD);
  }

  sheet.clear();

  // Title
  sheet.getRange('A1').setValue('ACQUISITION DASHBOARD').setFontSize(24).setFontWeight('bold');
  sheet.getRange('A1:J1').merge().setBackground('#1E3A5F').setFontColor('white').setHorizontalAlignment('center');

  // Last Updated
  sheet.getRange('A2').setValue('Last Updated: ' + new Date().toLocaleString());
  sheet.getRange('A2:J2').merge().setHorizontalAlignment('center').setFontStyle('italic');

  // Quick Stats Row
  sheet.getRange('A4').setValue('QUICK STATS').setFontWeight('bold').setFontSize(14);

  // Stat boxes
  const statBoxes = [
    { label: 'Total Active Leads', formula: '=COUNTIF(\'Main Page\'!T:T,"<>Closed")-COUNTIF(\'Main Page\'!T:T,"Dead")', col: 'A' },
    { label: 'Pending Offers', formula: '=COUNTIF(\'Main Page\'!T:T,"Offer Pending")', col: 'C' },
    { label: 'Under Contract', formula: '=COUNTIF(\'Contract Pipeline\'!W:W,"Under Contract")', col: 'E' },
    { label: 'Closed This Month', formula: '=COUNTIF(\'Contract Pipeline\'!W:W,"Closed")', col: 'G' },
    { label: 'Revenue This Month', formula: '=SUMIF(\'Contract Pipeline\'!W:W,"Closed",\'Contract Pipeline\'!V:V)', col: 'I' }
  ];

  statBoxes.forEach((stat, index) => {
    const col = stat.col;
    sheet.getRange(col + '5').setValue(stat.label).setFontWeight('bold').setBackground('#E8E8E8');
    sheet.getRange(col + '6').setFormula(stat.formula).setFontSize(20).setHorizontalAlignment('center');
    if (stat.label.includes('Revenue')) {
      sheet.getRange(col + '6').setNumberFormat('$#,##0');
    }
  });

  // Pipeline Summary
  sheet.getRange('A9').setValue('PIPELINE SUMMARY').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A9:E9').merge().setBackground('#2D5016').setFontColor('white');

  const pipelineHeaders = ['Stage', 'Count', '% of Total', 'Avg Value', 'Total Value'];
  sheet.getRange('A10:E10').setValues([pipelineHeaders]).setFontWeight('bold').setBackground('#E0E0E0');

  const pipelineStages = ['New', 'Contacted', 'Analyzing', 'Offer Pending', 'Under Contract'];
  pipelineStages.forEach((stage, index) => {
    const row = 11 + index;
    sheet.getRange('A' + row).setValue(stage);
    sheet.getRange('B' + row).setFormula('=COUNTIF(\'Main Page\'!T:T,"' + stage + '")');
    sheet.getRange('C' + row).setFormula('=IF(COUNTA(\'Main Page\'!T:T)>1,B' + row + '/(COUNTA(\'Main Page\'!T:T)-1),0)').setNumberFormat('0%');
    sheet.getRange('D' + row).setFormula('=AVERAGEIF(\'Main Page\'!T:T,"' + stage + '",\'Main Page\'!J:J)').setNumberFormat('$#,##0');
    sheet.getRange('E' + row).setFormula('=SUMIF(\'Main Page\'!T:T,"' + stage + '",\'Main Page\'!J:J)').setNumberFormat('$#,##0');
  });

  // Top Performing Zip Codes
  sheet.getRange('G9').setValue('TOP ZIP CODES').setFontWeight('bold').setFontSize(14);
  sheet.getRange('G9:J9').merge().setBackground('#5C1F1F').setFontColor('white');

  const zipHeaders = ['Zip Code', 'Leads', 'Contracts', 'Rate'];
  sheet.getRange('G10:J10').setValues([zipHeaders]).setFontWeight('bold').setBackground('#E0E0E0');

  // Monthly Trend placeholder
  sheet.getRange('A18').setValue('MONTHLY GOALS PROGRESS').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A18:J18').merge().setBackground('#4A1F5C').setFontColor('white');

  // Goal progress bars (using formulas and conditional formatting)
  const goals = [
    { name: 'Leads Generated', goal: CONFIG.MONTHLY_GOALS.LEADS_GENERATED, actual: '=COUNTA(\'Main Page\'!A:A)-1' },
    { name: 'Qualified Leads', goal: CONFIG.MONTHLY_GOALS.LEADS_QUALIFIED, actual: '=COUNTIF(\'Main Page\'!M:M,"Yes")' },
    { name: 'Offers Made', goal: CONFIG.MONTHLY_GOALS.OFFERS_MADE, actual: '=COUNTIF(\'Main Page\'!O:O,"Yes")' },
    { name: 'Contracts Signed', goal: CONFIG.MONTHLY_GOALS.CONTRACTS_SIGNED, actual: '=COUNTIF(\'Main Page\'!P:P,"Yes")' },
    { name: 'Deals Closed', goal: CONFIG.MONTHLY_GOALS.DEALS_CLOSED, actual: '=COUNTIF(\'Contract Pipeline\'!W:W,"Closed")' }
  ];

  sheet.getRange('A19:E19').setValues([['Metric', 'Goal', 'Actual', 'Progress', 'Status']]).setFontWeight('bold').setBackground('#E0E0E0');

  goals.forEach((goal, index) => {
    const row = 20 + index;
    sheet.getRange('A' + row).setValue(goal.name);
    sheet.getRange('B' + row).setValue(goal.goal);
    sheet.getRange('C' + row).setFormula(goal.actual);
    sheet.getRange('D' + row).setFormula('=IF(B' + row + '>0,C' + row + '/B' + row + ',0)').setNumberFormat('0%');
    sheet.getRange('E' + row).setFormula('=IF(D' + row + '>=1,"✅ Complete",IF(D' + row + '>=0.75,"🟡 On Track",IF(D' + row + '>=0.5,"🟠 Behind","🔴 At Risk")))');
  });

  // Set column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
}

/**
 * Create or update Buyers List sheet
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
 * Setup data validation for dropdown fields
 */
function setupDataValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // On Market Leads validations
  const onMarket = ss.getSheetByName(CONFIG.SHEETS.ON_MARKET);
  if (onMarket) {
    // Source dropdown
    const sourceRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.SOURCES.ON_MARKET, true)
      .build();
    onMarket.getRange('C2:C1000').setDataValidation(sourceRule);

    // Yes/No dropdowns
    const yesNoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true)
      .build();
    onMarket.getRange('U2:X1000').setDataValidation(yesNoRule);

    // Status dropdown
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.STATUS.LEAD, true)
      .build();
    onMarket.getRange('Y2:Y1000').setDataValidation(statusRule);

    // Property Type
    const propertyRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Single Family', 'Multi Family', 'Townhouse', 'Condo', 'Land', 'Commercial', 'Other'], true)
      .build();
    onMarket.getRange('J2:J1000').setDataValidation(propertyRule);
  }

  // Off Market Leads validations
  const offMarket = ss.getSheetByName(CONFIG.SHEETS.OFF_MARKET);
  if (offMarket) {
    // Source dropdown
    const sourceRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.SOURCES.OFF_MARKET, true)
      .build();
    offMarket.getRange('C2:C1000').setDataValidation(sourceRule);

    // Yes/No dropdowns
    const yesNoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true)
      .build();
    offMarket.getRange('W2:Z1000').setDataValidation(yesNoRule);

    // Status dropdown
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.STATUS.LEAD, true)
      .build();
    offMarket.getRange('AA2:AA1000').setDataValidation(statusRule);

    // Motivation Level
    const motivationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Hot', 'Warm', 'Cold'], true)
      .build();
    offMarket.getRange('AB2:AB1000').setDataValidation(motivationRule);
  }

  // Contract Pipeline validations
  const pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  if (pipeline) {
    // Disposition Strategy
    const dispRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.STATUS.DISPOSITION, true)
      .build();
    pipeline.getRange('P2:Q1000').setDataValidation(dispRule);

    // Status
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Under Contract', 'Pending Close', 'Closed', 'Cancelled'], true)
      .build();
    pipeline.getRange('W2:W1000').setDataValidation(statusRule);
  }

  // Main Page validations
  const mainPage = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);
  if (mainPage) {
    const yesNoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true)
      .build();
    mainPage.getRange('M2:P1000').setDataValidation(yesNoRule);

    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.STATUS.LEAD, true)
      .build();
    mainPage.getRange('T2:T1000').setDataValidation(statusRule);
  }

  // Zip Codes validations
  const zipCodes = ss.getSheetByName(CONFIG.SHEETS.ZIP_CODES);
  if (zipCodes) {
    const marketRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active', 'Paused', 'Retired'], true)
      .build();
    zipCodes.getRange('E2:E1000').setDataValidation(marketRule);

    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['High', 'Medium', 'Low'], true)
      .build();
    zipCodes.getRange('F2:F1000').setDataValidation(priorityRule);
  }
}

/**
 * Setup conditional formatting for visual indicators
 */
function setupConditionalFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Main Page status colors
  const mainPage = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);
  if (mainPage) {
    const statusRange = mainPage.getRange('T2:T1000');

    // Clear existing rules
    statusRange.clearFormat();

    const rules = mainPage.getConditionalFormatRules();

    // New - Blue
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('New')
      .setBackground('#E3F2FD')
      .setRanges([statusRange])
      .build());

    // Under Contract - Green
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Under Contract')
      .setBackground('#E8F5E9')
      .setRanges([statusRange])
      .build());

    // Dead - Red
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Dead')
      .setBackground('#FFEBEE')
      .setRanges([statusRange])
      .build());

    // Closed - Dark Green
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Closed')
      .setBackground('#C8E6C9')
      .setRanges([statusRange])
      .build());

    mainPage.setConditionalFormatRules(rules);
  }
}

// ============================================================================
// AUTOMATION FUNCTIONS
// ============================================================================

/**
 * Auto-generate Lead ID on edit
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  // Skip header row
  if (row === 1) return;

  // Auto-generate Lead ID for new entries
  if (col > 1 && sheet.getRange(row, 1).getValue() === '') {
    if ([CONFIG.SHEETS.ON_MARKET, CONFIG.SHEETS.OFF_MARKET, CONFIG.SHEETS.MAIN_PAGE].includes(sheetName)) {
      const prefix = sheetName === CONFIG.SHEETS.ON_MARKET ? 'OM' :
                     sheetName === CONFIG.SHEETS.OFF_MARKET ? 'OF' : 'MP';
      const id = prefix + '-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd') + '-' + row;
      sheet.getRange(row, 1).setValue(id);
    }
  }

  // Auto-set date added
  if (col > 2 && sheet.getRange(row, 2).getValue() === '') {
    if ([CONFIG.SHEETS.ON_MARKET, CONFIG.SHEETS.OFF_MARKET, CONFIG.SHEETS.MAIN_PAGE].includes(sheetName)) {
      sheet.getRange(row, 2).setValue(new Date());
    }
  }

  // Calculate MAO (Maximum Allowable Offer) when ARV or Repairs change
  if ([CONFIG.SHEETS.ON_MARKET, CONFIG.SHEETS.OFF_MARKET].includes(sheetName)) {
    calculateMAO(sheet, row);
  }

  // Auto-move to Contract Pipeline when "Contract Signed" = Yes
  // Handle from all three lead sheets
  const contractSignedColumns = {
    [CONFIG.SHEETS.ON_MARKET]: 24,    // Column X - Contract Signed (Yes/No)
    [CONFIG.SHEETS.OFF_MARKET]: 26,   // Column Z - Contract Signed (Yes/No)
    [CONFIG.SHEETS.MAIN_PAGE]: 16     // Column P - Contract Signed (Yes/No)
  };

  if (contractSignedColumns[sheetName] && col === contractSignedColumns[sheetName]) {
    if (e.value === 'Yes') {
      moveLeadToContractPipeline(sheet, row, sheetName);
    }
  }
}

/**
 * Calculate Maximum Allowable Offer (70% Rule)
 */
function calculateMAO(sheet, row) {
  const sheetName = sheet.getName();
  let arvCol, repairCol, maoCol;

  if (sheetName === CONFIG.SHEETS.ON_MARKET) {
    arvCol = 17;    // ARV Estimate
    repairCol = 18; // Repair Estimate
    maoCol = 19;    // MAO
  } else if (sheetName === CONFIG.SHEETS.OFF_MARKET) {
    arvCol = 18;
    repairCol = 19;
    maoCol = 20;
  } else {
    return;
  }

  const arv = sheet.getRange(row, arvCol).getValue();
  const repairs = sheet.getRange(row, repairCol).getValue();

  if (arv && repairs) {
    const mao = (arv * 0.7) - repairs;
    sheet.getRange(row, maoCol).setValue(Math.max(0, mao)).setNumberFormat('$#,##0');
  }
}

/**
 * Move lead to Contract Pipeline - handles all three source sheets
 * @param {Sheet} sourceSheet - The sheet where contract was signed
 * @param {number} row - The row number
 * @param {string} sheetName - Name of the source sheet
 */
function moveLeadToContractPipeline(sourceSheet, row, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);

  if (!pipeline) return;

  const rowData = sourceSheet.getRange(row, 1, 1, sourceSheet.getLastColumn()).getValues()[0];

  let pipelineData;

  if (sheetName === CONFIG.SHEETS.ON_MARKET) {
    // On Market Leads column mapping
    // Columns: Lead ID(0), Date(1), Source(2), MLS#(3), Address(4), City(5), County(6), State(7), Zip(8),
    //          PropType(9), Beds(10), Baths(11), Sqft(12), YearBuilt(13), ListPrice(14), DOM(15),
    //          ARV(16), Repairs(17), MAO(18), OfferAmt(19), Qualified(20), PassedTL(21), OfferMade(22),
    //          ContractSigned(23), Status(24), Notes(25), AgentName(26), AgentPhone(27), URL(28)
    pipelineData = [
      'CP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmm'), // Deal ID
      rowData[0],   // Original Lead ID
      new Date(),   // Date Under Contract
      rowData[4],   // Property Address
      rowData[5],   // City
      rowData[6],   // County
      rowData[7],   // State
      rowData[8],   // Zip Code
      rowData[19] || rowData[14], // Purchase Price (Offer Amount or List Price)
      rowData[16],  // ARV
      rowData[17],  // Repair Estimate
      '',           // Expected Profit (calculated)
      '',           // Earnest Money
      '',           // Inspection Period End
      '',           // Close Date
      '',           // Disposition Strategy
      '',           // Disposition Outcome
      '',           // Buyer Name
      '',           // Buyer Phone
      '',           // Assignment Fee
      '',           // Final Sale Price
      '',           // Actual Profit
      'Under Contract', // Status
      '',           // Title Company
      '',           // Title Contact
      '',           // Attorney
      rowData[25],  // Notes
      ''            // Days to Close
    ];

    // Update status on source sheet
    sourceSheet.getRange(row, 25).setValue('Under Contract');

  } else if (sheetName === CONFIG.SHEETS.OFF_MARKET) {
    // Off Market Leads column mapping
    // Columns: Lead ID(0), Date(1), Source(2), SourceContact(3), SourcePhone(4), SourceEmail(5),
    //          Address(6), City(7), County(8), State(9), Zip(10), PropType(11), Beds(12), Baths(13),
    //          Sqft(14), YearBuilt(15), AskingPrice(16), ARV(17), Repairs(18), MAO(19), OfferAmt(20),
    //          AssignmentFee(21), Qualified(22), PassedTL(23), OfferMade(24), ContractSigned(25),
    //          Status(26), Motivation(27), SellerName(28), SellerPhone(29), Notes(30)
    pipelineData = [
      'CP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmm'), // Deal ID
      rowData[0],   // Original Lead ID
      new Date(),   // Date Under Contract
      rowData[6],   // Property Address
      rowData[7],   // City
      rowData[8],   // County
      rowData[9],   // State
      rowData[10],  // Zip Code
      rowData[20] || rowData[16], // Purchase Price (Offer Amount or Asking Price)
      rowData[17],  // ARV
      rowData[18],  // Repair Estimate
      '',           // Expected Profit (calculated)
      '',           // Earnest Money
      '',           // Inspection Period End
      '',           // Close Date
      '',           // Disposition Strategy
      '',           // Disposition Outcome
      '',           // Buyer Name
      '',           // Buyer Phone
      rowData[21],  // Assignment Fee
      '',           // Final Sale Price
      '',           // Actual Profit
      'Under Contract', // Status
      '',           // Title Company
      '',           // Title Contact
      '',           // Attorney
      rowData[30],  // Notes
      ''            // Days to Close
    ];

    // Update status on source sheet
    sourceSheet.getRange(row, 27).setValue('Under Contract');

  } else {
    // Main Page column mapping
    // Columns: Lead ID(0), Date(1), LeadType(2), Source(3), Address(4), City(5), County(6), Zip(7),
    //          PropType(8), Price(9), ARV(10), MAO(11), Qualified(12), PassedTL(13), OfferMade(14),
    //          ContractSigned(15), DispOutcome(16), BuyerName(17), FinalPrice(18), Status(19), Notes(20)
    pipelineData = [
      'CP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmm'), // Deal ID
      rowData[0],   // Original Lead ID
      new Date(),   // Date Under Contract
      rowData[4],   // Property Address
      rowData[5],   // City
      rowData[6],   // County
      '',           // State
      rowData[7],   // Zip Code
      rowData[9],   // Purchase Price
      rowData[10],  // ARV
      '',           // Repair Estimate
      '',           // Expected Profit
      '',           // Earnest Money
      '',           // Inspection Period End
      '',           // Close Date
      '',           // Disposition Strategy
      '',           // Disposition Outcome
      rowData[17],  // Buyer Name
      '',           // Buyer Phone
      '',           // Assignment Fee
      rowData[18],  // Final Sale Price
      '',           // Actual Profit
      'Under Contract', // Status
      '',           // Title Company
      '',           // Title Contact
      '',           // Attorney
      rowData[20],  // Notes
      ''            // Days to Close
    ];

    // Update status on Main Page
    sourceSheet.getRange(row, 20).setValue('Under Contract');
  }

  // Add to Contract Pipeline
  pipeline.appendRow(pipelineData);

  // Also update Main Page if the source was On Market or Off Market
  if (sheetName !== CONFIG.SHEETS.MAIN_PAGE) {
    syncToMainPage(rowData, sheetName);
  }
}

/**
 * Sync contract status to Main Page when updated from On/Off Market sheets
 */
function syncToMainPage(rowData, sourceSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainPage = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);

  if (!mainPage) return;

  // Get the Lead ID to find in Main Page
  const leadId = rowData[0];
  if (!leadId) return;

  // Find the lead in Main Page
  const mainData = mainPage.getDataRange().getValues();
  for (let i = 1; i < mainData.length; i++) {
    if (mainData[i][0] === leadId) {
      // Update Contract Signed and Status columns
      mainPage.getRange(i + 1, 16).setValue('Yes');  // Contract Signed column
      mainPage.getRange(i + 1, 20).setValue('Under Contract');  // Status column
      break;
    }
  }
}

/**
 * Legacy function - redirects to new function for backward compatibility
 */
function autoMoveToContractPipeline(sourceSheet, row) {
  moveLeadToContractPipeline(sourceSheet, row, CONFIG.SHEETS.MAIN_PAGE);
}

/**
 * Refresh all dashboard calculations
 */
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);

  if (dashboard) {
    // Update timestamp
    dashboard.getRange('A2').setValue('Last Updated: ' + new Date().toLocaleString());

    // Force recalculation
    SpreadsheetApp.flush();
  }

  SpreadsheetApp.getUi().alert('Dashboard refreshed!');
}

/**
 * Update all KPI calculations
 */
function updateKPIs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createKPISheet(ss);
  SpreadsheetApp.getUi().alert('KPIs updated!');
}

/**
 * Reset monthly KPIs (typically run at month end)
 */
function resetMonthlyKPIs() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset Monthly KPIs',
    'This will archive current KPIs and reset for a new month. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Archive current data (you could expand this to save to a separate archive sheet)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createKPISheet(ss);

  ui.alert('Monthly KPIs have been reset!');
}

// ============================================================================
// FORM DIALOGS
// ============================================================================

/**
 * Show On Market Lead entry form
 */
function showOnMarketForm() {
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
      <div class="form-group">
        <label>Source *</label>
        <select id="source" required>
          <option value="">Select Source</option>
          <option value="Zillow">Zillow</option>
          <option value="Redfin">Redfin</option>
          <option value="Realtor.com">Realtor.com</option>
          <option value="MLS">MLS</option>
        </select>
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
      </div>

      <div class="row">
        <div class="col form-group">
          <label>County</label>
          <input type="text" id="county">
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

      <div class="row">
        <div class="col form-group">
          <label>Beds</label>
          <input type="number" id="beds">
        </div>
        <div class="col form-group">
          <label>Baths</label>
          <input type="number" id="baths" step="0.5">
        </div>
        <div class="col form-group">
          <label>Sqft</label>
          <input type="number" id="sqft">
        </div>
      </div>

      <div class="form-group">
        <label>Listing URL</label>
        <input type="url" id="url">
      </div>

      <div class="form-group">
        <label>Notes</label>
        <textarea id="notes" style="width:100%; height:60px;"></textarea>
      </div>

      <button type="button" class="btn btn-primary" onclick="submitForm()">Add Lead</button>
      <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
    </form>

    <script>
      function submitForm() {
        const data = {
          source: document.getElementById('source').value,
          mls: document.getElementById('mls').value,
          address: document.getElementById('address').value,
          city: document.getElementById('city').value,
          state: document.getElementById('state').value,
          county: document.getElementById('county').value,
          zip: document.getElementById('zip').value,
          price: document.getElementById('price').value,
          arv: document.getElementById('arv').value,
          beds: document.getElementById('beds').value,
          baths: document.getElementById('baths').value,
          sqft: document.getElementById('sqft').value,
          url: document.getElementById('url').value,
          notes: document.getElementById('notes').value
        };

        google.script.run
          .withSuccessHandler(() => {
            alert('Lead added successfully!');
            google.script.host.close();
          })
          .withFailureHandler(err => alert('Error: ' + err.message))
          .addOnMarketLead(data);
      }
    </script>
  `)
  .setWidth(450)
  .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Add On Market Lead');
}

/**
 * Show Off Market Lead entry form
 */
function showOffMarketForm() {
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
          <label>Source *</label>
          <select id="source" required>
            <option value="">Select Source</option>
            <option value="Wholesaler">Wholesaler</option>
            <option value="Agent Referral">Agent Referral</option>
            <option value="Direct Mail">Direct Mail</option>
            <option value="Cold Call">Cold Call</option>
            <option value="Driving for Dollars">Driving for Dollars</option>
            <option value="Probate">Probate</option>
            <option value="Tax Lien">Tax Lien</option>
            <option value="Other">Other</option>
          </select>
        </div>
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
        <label>Source Contact Name</label>
        <input type="text" id="sourceContact">
      </div>

      <div class="row">
        <div class="col form-group">
          <label>Source Phone</label>
          <input type="tel" id="sourcePhone">
        </div>
        <div class="col form-group">
          <label>Source Email</label>
          <input type="email" id="sourceEmail">
        </div>
      </div>

      <hr>

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
      </div>

      <div class="row">
        <div class="col form-group">
          <label>County</label>
          <input type="text" id="county">
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
          <label>Repair Estimate</label>
          <input type="number" id="repairs">
        </div>
        <div class="col form-group">
          <label>Assignment Fee</label>
          <input type="number" id="assignmentFee">
        </div>
      </div>

      <hr>

      <div class="form-group">
        <label>Seller Name</label>
        <input type="text" id="sellerName">
      </div>

      <div class="form-group">
        <label>Seller Phone</label>
        <input type="tel" id="sellerPhone">
      </div>

      <div class="form-group">
        <label>Notes</label>
        <textarea id="notes" style="width:100%; height:60px;"></textarea>
      </div>

      <button type="button" class="btn btn-primary" onclick="submitForm()">Add Lead</button>
      <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
    </form>

    <script>
      function submitForm() {
        const data = {
          source: document.getElementById('source').value,
          motivation: document.getElementById('motivation').value,
          sourceContact: document.getElementById('sourceContact').value,
          sourcePhone: document.getElementById('sourcePhone').value,
          sourceEmail: document.getElementById('sourceEmail').value,
          address: document.getElementById('address').value,
          city: document.getElementById('city').value,
          state: document.getElementById('state').value,
          county: document.getElementById('county').value,
          zip: document.getElementById('zip').value,
          price: document.getElementById('price').value,
          arv: document.getElementById('arv').value,
          repairs: document.getElementById('repairs').value,
          assignmentFee: document.getElementById('assignmentFee').value,
          sellerName: document.getElementById('sellerName').value,
          sellerPhone: document.getElementById('sellerPhone').value,
          notes: document.getElementById('notes').value
        };

        google.script.run
          .withSuccessHandler(() => {
            alert('Lead added successfully!');
            google.script.host.close();
          })
          .withFailureHandler(err => alert('Error: ' + err.message))
          .addOffMarketLead(data);
      }
    </script>
  `)
  .setWidth(450)
  .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, 'Add Off Market Lead');
}

/**
 * Add On Market Lead from form
 */
function addOnMarketLead(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ON_MARKET);

  if (!sheet) throw new Error('On Market Leads sheet not found. Please initialize workspace first.');

  const id = 'OM-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmm');

  const rowData = [
    id,           // Lead ID
    new Date(),   // Date Added
    data.source,  // Source
    data.mls,     // MLS #
    data.address, // Property Address
    data.city,    // City
    data.county,  // County
    data.state,   // State
    data.zip,     // Zip Code
    '',           // Property Type
    data.beds,    // Beds
    data.baths,   // Baths
    data.sqft,    // Sqft
    '',           // Year Built
    data.price,   // List Price
    '',           // Days on Market
    data.arv,     // ARV Estimate
    '',           // Repair Estimate
    '',           // MAO
    '',           // Offer Amount
    'No',         // Qualified
    'No',         // Passed to Team Lead
    'No',         // Offer Made
    'No',         // Contract Signed
    'New',        // Status
    data.notes,   // Notes
    '',           // Agent Name
    '',           // Agent Phone
    data.url      // Listing URL
  ];

  sheet.appendRow(rowData);

  // Also add to Main Page
  addToMainPage(id, 'On Market', data);
}

/**
 * Add Off Market Lead from form
 */
function addOffMarketLead(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.OFF_MARKET);

  if (!sheet) throw new Error('Off Market Leads sheet not found. Please initialize workspace first.');

  const id = 'OF-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmm');

  // Calculate MAO
  let mao = '';
  if (data.arv && data.repairs) {
    mao = (parseFloat(data.arv) * 0.7) - parseFloat(data.repairs);
  }

  const rowData = [
    id,                 // Lead ID
    new Date(),         // Date Added
    data.source,        // Source
    data.sourceContact, // Source Contact
    data.sourcePhone,   // Source Phone
    data.sourceEmail,   // Source Email
    data.address,       // Property Address
    data.city,          // City
    data.county,        // County
    data.state,         // State
    data.zip,           // Zip Code
    '',                 // Property Type
    '',                 // Beds
    '',                 // Baths
    '',                 // Sqft
    '',                 // Year Built
    data.price,         // Asking Price
    data.arv,           // ARV Estimate
    data.repairs,       // Repair Estimate
    mao,                // MAO
    '',                 // Offer Amount
    data.assignmentFee, // Assignment Fee
    'No',               // Qualified
    'No',               // Passed to Team Lead
    'No',               // Offer Made
    'No',               // Contract Signed
    'New',              // Status
    data.motivation,    // Motivation Level
    data.sellerName,    // Seller Name
    data.sellerPhone,   // Seller Phone
    data.notes          // Notes
  ];

  sheet.appendRow(rowData);

  // Also add to Main Page
  addToMainPage(id, 'Off Market', data);
}

/**
 * Add lead to Main Page (consolidated view)
 */
function addToMainPage(id, leadType, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);

  if (!sheet) return;

  const rowData = [
    id,           // Lead ID
    new Date(),   // Date Added
    leadType,     // Lead Type
    data.source,  // Source
    data.address, // Property Address
    data.city,    // City
    data.county,  // County
    data.zip,     // Zip Code
    '',           // Property Type
    data.price,   // List/Ask Price
    data.arv || '', // ARV
    '',           // MAO
    'No',         // Qualified
    'No',         // Passed to Team Lead
    'No',         // Offer Made
    'No',         // Contract Signed
    '',           // Disposition Outcome
    '',           // Buyer Name
    '',           // Final Sale Price
    'New',        // Status
    data.notes    // Notes
  ];

  sheet.appendRow(rowData);
}

/**
 * Move selected lead to Contract Pipeline
 */
function moveToContractPipeline() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert('Please select a data row (not header)');
    return;
  }

  const sheetName = sheet.getName();
  if (![CONFIG.SHEETS.ON_MARKET, CONFIG.SHEETS.OFF_MARKET, CONFIG.SHEETS.MAIN_PAGE].includes(sheetName)) {
    SpreadsheetApp.getUi().alert('Please run this from a leads sheet');
    return;
  }

  autoMoveToContractPipeline(sheet, row);

  // Update status on source sheet
  const statusCol = sheet.getLastColumn();
  for (let col = 1; col <= statusCol; col++) {
    if (sheet.getRange(1, col).getValue() === 'Status') {
      sheet.getRange(row, col).setValue('Under Contract');
      break;
    }
  }

  SpreadsheetApp.getUi().alert('Lead moved to Contract Pipeline!');
}

/**
 * Show Zip Code import dialog
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
          .withSuccessHandler(count => {
            alert(count + ' zip codes imported!');
            google.script.host.close();
          })
          .importZipCodes(zips);
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(html, 'Import Zip Codes');
}

/**
 * Import zip codes from text
 */
function importZipCodes(zipText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ZIP_CODES);

  if (!sheet) throw new Error('Zip Codes sheet not found');

  // Parse zip codes
  const zips = zipText.split(/[\n,]+/).map(z => z.trim()).filter(z => z && /^\d{5}$/.test(z));

  // Get existing zips
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

/**
 * Show help dialog
 */
function showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; line-height: 1.6; }
      h2 { color: #1E3A5F; }
      h3 { color: #2D5016; margin-top: 20px; }
      ul { padding-left: 20px; }
      .tip { background: #FFF3CD; padding: 10px; border-radius: 4px; margin: 10px 0; }
    </style>

    <h2>Acquisition Workspace Help</h2>

    <h3>Sheet Overview</h3>
    <ul>
      <li><strong>On Market Leads</strong> - Track properties from Zillow, Redfin, MLS</li>
      <li><strong>Off Market Leads</strong> - Track deals from wholesalers, agents, direct mail</li>
      <li><strong>Contract Pipeline</strong> - Track deals under contract to close</li>
      <li><strong>Main Page</strong> - Consolidated view of all leads</li>
      <li><strong>Lead KPIs</strong> - Monthly performance metrics and conversion rates</li>
      <li><strong>Zip Codes</strong> - Target markets and performance by area</li>
      <li><strong>Dashboard</strong> - Visual overview of pipeline and goals</li>
      <li><strong>Buyers List</strong> - Track your cash buyers</li>
    </ul>

    <h3>Automation Features</h3>
    <ul>
      <li><strong>Auto Lead ID</strong> - IDs are generated automatically when you enter data</li>
      <li><strong>Auto Date</strong> - Date Added is filled automatically</li>
      <li><strong>MAO Calculator</strong> - Calculates 70% rule automatically</li>
      <li><strong>Contract Pipeline</strong> - Leads auto-move when Contract Signed = Yes</li>
      <li><strong>KPI Tracking</strong> - All metrics update automatically</li>
    </ul>

    <h3>Quick Tips</h3>
    <div class="tip">
      <strong>Tip:</strong> Use the dropdown menus for Status, Source, and Yes/No fields to ensure data consistency.
    </div>
    <div class="tip">
      <strong>Tip:</strong> Click "Refresh Dashboard" after making changes to update all metrics.
    </div>

    <h3>Keyboard Shortcuts</h3>
    <ul>
      <li><strong>Ctrl/Cmd + Shift + \\</strong> - Open Extensions menu</li>
      <li><strong>F5</strong> - Refresh formulas</li>
    </ul>
  `)
  .setWidth(500)
  .setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, 'Acquisition Workspace Help');
}

// ============================================================================
// SCHEDULED TRIGGERS (Optional - Set up manually in Apps Script)
// ============================================================================

/**
 * Daily summary email (configure trigger to run daily)
 */
function sendDailySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainPage = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);
  const pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);

  // Get counts
  const newLeads = mainPage ? mainPage.getRange('T:T').getValues().flat().filter(s => s === 'New').length : 0;
  const underContract = pipeline ? pipeline.getRange('W:W').getValues().flat().filter(s => s === 'Under Contract').length : 0;

  const subject = 'Daily Acquisition Summary - ' + new Date().toLocaleDateString();
  const body = `
Daily Acquisition Summary

New Leads Today: ${newLeads}
Properties Under Contract: ${underContract}

View Dashboard: ${ss.getUrl()}
  `;

  // Uncomment and add email to enable:
  // MailApp.sendEmail('your-email@example.com', subject, body);
}

/**
 * Weekly KPI report (configure trigger to run weekly)
 */
function sendWeeklyKPIReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Build report from KPI sheet
  const subject = 'Weekly KPI Report - ' + new Date().toLocaleDateString();
  const body = `
Weekly KPI Report

View full KPIs: ${ss.getUrl()}#gid=${ss.getSheetByName(CONFIG.SHEETS.LEAD_KPIS)?.getSheetId() || 0}
  `;

  // Uncomment and add email to enable:
  // MailApp.sendEmail('your-email@example.com', subject, body);
}
