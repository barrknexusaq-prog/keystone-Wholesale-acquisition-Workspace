/**
 * Utilities.gs - Helper Functions & Custom Formulas
 * Acquisition Workspace v4.0
 */

// ============================================================================
// CUSTOM SPREADSHEET FORMULAS
// ============================================================================

/**
 * Calculate Maximum Allowable Offer (70% Rule)
 * Usage: =MAO(ARV, Repairs, WholesaleFee)
 * @param {number} arv - After Repair Value
 * @param {number} repairs - Repair cost estimate
 * @param {number} fee - Your wholesale fee (default 10000)
 * @return {number} Maximum allowable offer
 * @customfunction
 */
function MAO(arv, repairs, fee) {
  if (!arv || arv <= 0) return 0;
  const wholesaleFee = fee || 10000;
  return Math.max(0, (arv * 0.7) - (repairs || 0) - wholesaleFee);
}

/**
 * Calculate potential profit
 * Usage: =PROFIT(PurchasePrice, ARV, Repairs, HoldingCosts, SellingCosts)
 * @customfunction
 */
function PROFIT(purchasePrice, arv, repairs, holdingCosts, sellingCosts) {
  if (!purchasePrice || !arv) return 0;
  const costs = (purchasePrice || 0) + (repairs || 0) + (holdingCosts || 0);
  const sellCosts = sellingCosts || (arv * 0.08); // Default 8%
  return arv - sellCosts - costs;
}

/**
 * Calculate assignment fee
 * Usage: =ASSIGNMENT_FEE(ContractPrice, EndBuyerPrice)
 * @customfunction
 */
function ASSIGNMENT_FEE(contractPrice, endBuyerPrice) {
  return (endBuyerPrice || 0) - (contractPrice || 0);
}

/**
 * Calculate days until deadline
 * Usage: =DAYS_UNTIL(date)
 * @customfunction
 */
function DAYS_UNTIL(deadline) {
  if (!deadline) return '';
  const today = new Date();
  const target = new Date(deadline);
  return Math.floor((target - today) / (24 * 60 * 60 * 1000));
}

/**
 * Calculate days between dates
 * Usage: =DAYS_BETWEEN(start, end)
 * @customfunction
 */
function DAYS_BETWEEN(startDate, endDate) {
  if (!startDate || !endDate) return '';
  const start = new Date(startDate);
  const end = new Date(endDate);
  return Math.floor((end - start) / (24 * 60 * 60 * 1000));
}

/**
 * Format phone number
 * Usage: =FORMAT_PHONE(phone)
 * @customfunction
 */
function FORMAT_PHONE(phone) {
  if (!phone) return '';
  const cleaned = phone.toString().replace(/\D/g, '');
  if (cleaned.length === 10) {
    return `(${cleaned.slice(0,3)}) ${cleaned.slice(3,6)}-${cleaned.slice(6)}`;
  }
  return phone;
}

/**
 * Extract zip code from address
 * Usage: =EXTRACT_ZIP(address)
 * @customfunction
 */
function EXTRACT_ZIP(address) {
  if (!address) return '';
  const match = address.toString().match(/\b(\d{5})(-\d{4})?\b/);
  return match ? match[1] : '';
}

// ============================================================================
// ZIP CODE UTILITIES
// ============================================================================

/**
 * Update zip code statistics
 */
function updateZipCodeStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zipSheet = ss.getSheetByName(CONFIG.SHEETS.ZIP_CODES);
  const onMarket = CONFIG.SHEETS.ON_MARKET;
  const offMarket = CONFIG.SHEETS.OFF_MARKET;

  if (!zipSheet || zipSheet.getLastRow() < 2) return;

  const lastRow = zipSheet.getLastRow();

  for (let row = 2; row <= lastRow; row++) {
    const zip = zipSheet.getRange('A' + row).getValue();
    if (!zip) continue;

    // Set formulas for leads and contracts
    zipSheet.getRange('I' + row).setFormula(
      `=COUNTIF('${onMarket}'!J:J,"${zip}")+COUNTIF('${offMarket}'!I:I,"${zip}")`
    );
    zipSheet.getRange('J' + row).setFormula(
      `=COUNTIFS('${onMarket}'!J:J,"${zip}",'${onMarket}'!W:W,"Yes")+COUNTIFS('${offMarket}'!I:I,"${zip}",'${offMarket}'!V:V,"Yes")`
    );
    zipSheet.getRange('K' + row).setFormula(`=IF(I${row}>0,J${row}/I${row},0)`).setNumberFormat('0%');
    zipSheet.getRange('M' + row).setValue(new Date());
  }

  SpreadsheetApp.getUi().alert('Zip code statistics updated!');
}

/**
 * Get performance for a specific zip code
 */
function getZipCodePerformance(zipCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let totalLeads = 0;
  let contracts = 0;

  // Count from On Market
  const onMarket = ss.getSheetByName(CONFIG.SHEETS.ON_MARKET);
  if (onMarket && onMarket.getLastRow() > 1) {
    const data = onMarket.getRange('J2:W' + onMarket.getLastRow()).getValues();
    data.forEach(row => {
      if (row[0] == zipCode) {
        totalLeads++;
        if (row[13] === 'Yes') contracts++; // Contract Signed column
      }
    });
  }

  // Count from Off Market
  const offMarket = ss.getSheetByName(CONFIG.SHEETS.OFF_MARKET);
  if (offMarket && offMarket.getLastRow() > 1) {
    const data = offMarket.getRange('I2:V' + offMarket.getLastRow()).getValues();
    data.forEach(row => {
      if (row[0] == zipCode) {
        totalLeads++;
        if (row[13] === 'Yes') contracts++; // Contract Signed column
      }
    });
  }

  return {
    zipCode,
    totalLeads,
    contracts,
    conversionRate: totalLeads > 0 ? (contracts / totalLeads) : 0
  };
}

// ============================================================================
// BUYER MATCHING
// ============================================================================

/**
 * Find buyers matching property criteria
/**
 * Find buyers matching property criteria
 * Buyers List columns (Dispo Workspace 3.0 format):
 *   0: Timestamp, 1: Name, 2: Phone, 3: Email, 4: Type of Buyer,
 *   5: Asset Type, 6: Deal Preference, 7: Exit Strategy, 8: Market,
 *   9: City/County/Zip-code(s), 10: Criteria
 *
 * @param {string} zipCode - Property zip code
 * @param {string} propertyType - Property type (e.g., "Single Family")
 * @param {string} market - Optional state code (FL, TN, TX, GA)
 * @return {Array} Matching buyer names
 */
function findMatchingBuyers(zipCode, propertyType, market) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const buyerSheet = ss.getSheetByName(CONFIG.SHEETS.BUYERS_LIST);

  if (!buyerSheet || buyerSheet.getLastRow() < 2) return [];

  const data = buyerSheet.getDataRange().getValues();
  const matches = [];

  for (let i = 1; i < data.length; i++) {
    const name = (data[i][1] || '').toString().trim();
    if (!name) continue; // Skip empty rows

    const assetTypes = (data[i][5] || '').toString().toLowerCase();
    const buyerMarket = (data[i][8] || '').toString().toUpperCase();
    const locations = (data[i][9] || '').toString().toLowerCase();

    // Match on asset type (property type must appear in buyer's asset types)
    const typeMatch = !assetTypes || !propertyType ||
      assetTypes.includes((propertyType || '').toLowerCase());

    // Match on market (state) or zip code
    let locationMatch = true;
    if (market) {
      locationMatch = !buyerMarket || buyerMarket.includes(market.toUpperCase());
    }
    if (zipCode) {
      const zipStr = zipCode.toString();
      locationMatch = locationMatch || locations.includes(zipStr);
    }

    if (typeMatch && locationMatch) {
      matches.push(name);
    }
  }

  return matches;
}

    }
  }

  return matches;
}

/**
 * Custom formula to show matching buyers
 * Usage: =MATCHING_BUYERS(ZipCode, PropertyType, Market)
 * @param {string} zipCode - Zip code to match against buyer locations
 * @param {string} propertyType - Property type (Single Family, Multi-Family, etc.)
 * @param {string} market - State code (FL, TN, TX, GA)
 * @customfunction
 */
function MATCHING_BUYERS(zipCode, propertyType, market) {
  const matches = findMatchingBuyers(zipCode, propertyType, market);
  return matches.length > 0 ? matches.join(', ') : 'No matches';
}

// ============================================================================
// EXPORT FUNCTIONS
// ============================================================================

/**
 * Export On Market leads to CSV
 */
function exportOnMarketToCSV() {
  exportSheetToCSV(CONFIG.SHEETS.ON_MARKET);
}

/**
 * Export Off Market leads to CSV
 */
function exportOffMarketToCSV() {
  exportSheetToCSV(CONFIG.SHEETS.OFF_MARKET);
}

/**
 * Export Contract Pipeline to CSV
 */
function exportPipelineToCSV() {
  exportSheetToCSV(CONFIG.SHEETS.CONTRACT_PIPELINE);
}

/**
 * Generic CSV export
 */
function exportSheetToCSV(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`${sheetName} sheet not found`);
    return;
  }

  const data = sheet.getDataRange().getValues();
  let csv = '';

  data.forEach(row => {
    csv += row.map(cell => {
      let value = cell.toString();
      if (value.includes(',') || value.includes('"') || value.includes('\n')) {
        value = '"' + value.replace(/"/g, '""') + '"';
      }
      return value;
    }).join(',') + '\n';
  });

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      textarea { width: 100%; height: 300px; font-family: monospace; }
      .btn { padding: 10px 20px; margin-top: 10px; cursor: pointer; }
    </style>
    <h3>Export: ${sheetName}</h3>
    <p>Copy the data below and save as .csv file:</p>
    <textarea id="csv">${csv}</textarea>
    <br>
    <button class="btn" onclick="copyToClipboard()">Copy to Clipboard</button>
    <script>
      function copyToClipboard() {
        const textarea = document.getElementById('csv');
        textarea.select();
        document.execCommand('copy');
        alert('Copied to clipboard!');
      }
    </script>
  `).setWidth(600).setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, 'Export to CSV');
}

// ============================================================================
// BULK OPERATIONS
// ============================================================================

/**
 * Bulk update offer status for selected rows
 */
function bulkMarkOfferMade() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  const selection = sheet.getActiveRange();

  let offerCol;
  if (sheetName === CONFIG.SHEETS.ON_MARKET) {
    offerCol = 22; // Column V
  } else if (sheetName === CONFIG.SHEETS.OFF_MARKET) {
    offerCol = 21; // Column U
  } else {
    SpreadsheetApp.getUi().alert('Please run from On Market or Off Market sheet');
    return;
  }

  const startRow = Math.max(2, selection.getRow());
  const endRow = selection.getLastRow();
  let count = 0;

  for (let row = startRow; row <= endRow; row++) {
    sheet.getRange(row, offerCol).setValue('Yes');
    count++;
  }

  SpreadsheetApp.getUi().alert(`Marked ${count} leads as Offer Made`);
}

/**
 * Bulk update contract signed for selected rows
 */
function bulkMarkContractSigned() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  const selection = sheet.getActiveRange();

  let contractCol;
  if (sheetName === CONFIG.SHEETS.ON_MARKET) {
    contractCol = 23; // Column W
  } else if (sheetName === CONFIG.SHEETS.OFF_MARKET) {
    contractCol = 22; // Column V
  } else {
    SpreadsheetApp.getUi().alert('Please run from On Market or Off Market sheet');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Bulk Contract Signed',
    'This will mark selected rows as Contract Signed and move them to the Pipeline. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const startRow = Math.max(2, selection.getRow());
  const endRow = selection.getLastRow();
  let count = 0;

  for (let row = startRow; row <= endRow; row++) {
    sheet.getRange(row, contractCol).setValue('Yes');
    count++;
  }

  SpreadsheetApp.getUi().alert(`Marked ${count} leads as Contract Signed. They will be added to Pipeline.`);
}

// ============================================================================
// REPORTING
// ============================================================================

/**
 * Generate performance report
 */
function generatePerformanceReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get all data
  const onMarket = ss.getSheetByName(CONFIG.SHEETS.ON_MARKET);
  const offMarket = ss.getSheetByName(CONFIG.SHEETS.OFF_MARKET);
  const pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);

  const totalOnMarket = onMarket ? Math.max(0, onMarket.getLastRow() - 1) : 0;
  const totalOffMarket = offMarket ? Math.max(0, offMarket.getLastRow() - 1) : 0;

  // Count pipeline stats
  let underContract = 0;
  let pendingClose = 0;
  let closed = 0;
  let totalRevenue = 0;

  if (pipeline && pipeline.getLastRow() > 1) {
    const data = pipeline.getRange('X2:Y' + pipeline.getLastRow()).getValues();
    data.forEach(row => {
      const profit = row[0] || 0;
      const status = row[1];
      if (status === 'Under Contract') underContract++;
      if (status === 'Pending Close') pendingClose++;
      if (status === 'Closed') {
        closed++;
        totalRevenue += profit;
      }
    });
  }

  // Team performance
  const teamPerf = {};
  CONFIG.TEAM_MEMBERS.forEach(m => teamPerf[m] = { leads: 0, contracts: 0 });

  if (onMarket && onMarket.getLastRow() > 1) {
    const data = onMarket.getRange('C2:W' + onMarket.getLastRow()).getValues();
    data.forEach(row => {
      if (teamPerf[row[0]]) {
        teamPerf[row[0]].leads++;
        if (row[20] === 'Yes') teamPerf[row[0]].contracts++;
      }
    });
  }

  if (offMarket && offMarket.getLastRow() > 1) {
    const data = offMarket.getRange('C2:V' + offMarket.getLastRow()).getValues();
    data.forEach(row => {
      if (teamPerf[row[0]]) {
        teamPerf[row[0]].leads++;
        if (row[19] === 'Yes') teamPerf[row[0]].contracts++;
      }
    });
  }

  // Build report
  let report = 'ACQUISITION PERFORMANCE REPORT\n';
  report += '================================\n\n';
  report += `Generated: ${new Date().toLocaleString()}\n\n`;

  report += 'LEAD SUMMARY\n';
  report += '------------\n';
  report += `On Market Leads: ${totalOnMarket}\n`;
  report += `Off Market Leads: ${totalOffMarket}\n`;
  report += `Total Leads: ${totalOnMarket + totalOffMarket}\n\n`;

  report += 'PIPELINE\n';
  report += '--------\n';
  report += `Under Contract: ${underContract}\n`;
  report += `Pending Close: ${pendingClose}\n`;
  report += `Closed: ${closed}\n`;
  report += `Total Revenue: $${totalRevenue.toLocaleString()}\n\n`;

  report += 'TEAM PERFORMANCE\n';
  report += '----------------\n';
  Object.keys(teamPerf).forEach(member => {
    const p = teamPerf[member];
    const rate = p.leads > 0 ? ((p.contracts / p.leads) * 100).toFixed(1) : '0';
    report += `${member}: ${p.leads} leads, ${p.contracts} contracts (${rate}%)\n`;
  });

  // Show report
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: monospace; padding: 20px; white-space: pre-wrap; }
    </style>
    <pre>${report}</pre>
  `).setWidth(500).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Performance Report');

  return report;
}

// ============================================================================
// DATA CLEANUP
// ============================================================================

/**
 * Clean up empty rows
 */
function removeEmptyRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const isEmpty = row.every(cell => cell === '' || cell === null);
    if (isEmpty) {
      rowsToDelete.push(i + 1);
    }
  }

  rowsToDelete.forEach(row => sheet.deleteRow(row));

  SpreadsheetApp.getUi().alert(`Removed ${rowsToDelete.length} empty rows`);
}

/**
 * Sort sheet by date (newest first)
 */
function sortByDateDesc() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getLastRow() < 3) return;

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  range.sort({ column: 2, ascending: false }); // Date Added column

  SpreadsheetApp.getUi().alert('Sorted by date (newest first)');
}

/**
 * Sort sheet by date (oldest first)
 */
function sortByDateAsc() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getLastRow() < 3) return;

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  range.sort({ column: 2, ascending: true }); // Date Added column

  SpreadsheetApp.getUi().alert('Sorted by date (oldest first)');
}
