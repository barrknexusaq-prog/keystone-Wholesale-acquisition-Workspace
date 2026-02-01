/**
 * Utilities.gs - Helper Functions & Data Processing
 *
 * This file contains utility functions for data manipulation,
 * calculations, and integration helpers
 */

// ============================================================================
// PROPERTY ANALYSIS FUNCTIONS
// ============================================================================

/**
 * Calculate Maximum Allowable Offer using the 70% rule
 * @param {number} arv - After Repair Value
 * @param {number} repairs - Estimated repair costs
 * @param {number} wholesaleFee - Your target wholesale fee (default $10,000)
 * @return {number} Maximum allowable offer
 */
function calculateMAOFormula(arv, repairs, wholesaleFee = 10000) {
  if (!arv || arv <= 0) return 0;
  const mao = (arv * 0.7) - (repairs || 0) - wholesaleFee;
  return Math.max(0, mao);
}

/**
 * Custom function for use in spreadsheet cells
 * Usage: =MAO(ARV, Repairs, WholesaleFee)
 */
function MAO(arv, repairs, wholesaleFee) {
  return calculateMAOFormula(arv, repairs, wholesaleFee || 10000);
}

/**
 * Calculate potential profit on a deal
 * @param {number} purchasePrice
 * @param {number} arv
 * @param {number} repairs
 * @param {number} holdingCosts
 * @param {number} sellingCosts
 * @return {number} Potential profit
 */
function POTENTIAL_PROFIT(purchasePrice, arv, repairs, holdingCosts, sellingCosts) {
  if (!purchasePrice || !arv) return 0;
  const totalCosts = purchasePrice + (repairs || 0) + (holdingCosts || 0);
  const netSale = arv - (sellingCosts || arv * 0.08); // Default 8% selling costs
  return netSale - totalCosts;
}

/**
 * Calculate assignment fee potential
 * @param {number} contractPrice - Your contract price with seller
 * @param {number} endBuyerPrice - Price you can sell assignment for
 * @return {number} Assignment fee
 */
function ASSIGNMENT_FEE(contractPrice, endBuyerPrice) {
  return (endBuyerPrice || 0) - (contractPrice || 0);
}

/**
 * Calculate cash-on-cash return for rental analysis
 */
function CASH_ON_CASH(annualCashFlow, totalCashInvested) {
  if (!totalCashInvested || totalCashInvested <= 0) return 0;
  return annualCashFlow / totalCashInvested;
}

// ============================================================================
// DATA FORMATTING FUNCTIONS
// ============================================================================

/**
 * Format phone number to (XXX) XXX-XXXX
 * @param {string} phone - Raw phone number
 * @return {string} Formatted phone number
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
 * Format address properly
 * @param {string} address
 * @return {string} Properly formatted address
 */
function FORMAT_ADDRESS(address) {
  if (!address) return '';
  return address.toString()
    .trim()
    .replace(/\s+/g, ' ')
    .split(' ')
    .map(word => {
      if (['NE', 'NW', 'SE', 'SW', 'N', 'S', 'E', 'W'].includes(word.toUpperCase())) {
        return word.toUpperCase();
      }
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(' ');
}

/**
 * Extract zip code from full address
 * @param {string} address
 * @return {string} Zip code
 */
function EXTRACT_ZIP(address) {
  if (!address) return '';
  const match = address.toString().match(/\b(\d{5})(-\d{4})?\b/);
  return match ? match[1] : '';
}

/**
 * Calculate days between two dates
 * @param {Date} startDate
 * @param {Date} endDate
 * @return {number} Number of days
 */
function DAYS_BETWEEN(startDate, endDate) {
  if (!startDate || !endDate) return '';
  const start = new Date(startDate);
  const end = new Date(endDate);
  return Math.floor((end - start) / (24 * 60 * 60 * 1000));
}

/**
 * Calculate days until a deadline
 * @param {Date} deadline
 * @return {number} Days until deadline (negative if passed)
 */
function DAYS_UNTIL(deadline) {
  if (!deadline) return '';
  const today = new Date();
  const target = new Date(deadline);
  return Math.floor((target - today) / (24 * 60 * 60 * 1000));
}

// ============================================================================
// LEAD SCORING FUNCTIONS
// ============================================================================

/**
 * Calculate lead score based on various factors
 * Returns 1-100 score
 */
function LEAD_SCORE(equityPercent, motivation, daysOld, propertyCondition) {
  let score = 50; // Base score

  // Equity factor (up to 25 points)
  if (equityPercent >= 50) score += 25;
  else if (equityPercent >= 30) score += 15;
  else if (equityPercent >= 20) score += 10;

  // Motivation factor (up to 25 points)
  const motivationLower = (motivation || '').toString().toLowerCase();
  if (motivationLower === 'hot') score += 25;
  else if (motivationLower === 'warm') score += 15;
  else if (motivationLower === 'cold') score += 5;

  // Freshness factor (up to 15 points)
  if (daysOld <= 1) score += 15;
  else if (daysOld <= 3) score += 10;
  else if (daysOld <= 7) score += 5;
  else if (daysOld > 30) score -= 10;

  // Property condition factor (up to 10 points)
  const condition = (propertyCondition || '').toString().toLowerCase();
  if (condition === 'poor' || condition === 'needs work') score += 10;
  else if (condition === 'fair') score += 5;

  return Math.max(0, Math.min(100, score));
}

/**
 * Get lead priority based on score
 */
function LEAD_PRIORITY(score) {
  if (score >= 80) return 'Hot';
  if (score >= 60) return 'Warm';
  if (score >= 40) return 'Cool';
  return 'Cold';
}

// ============================================================================
// ZIP CODE ANALYSIS FUNCTIONS
// ============================================================================

/**
 * Get lead count for a specific zip code
 */
function getLeadCountByZip(zipCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainPage = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);
  if (!mainPage) return 0;

  const data = mainPage.getRange('H2:H' + mainPage.getLastRow()).getValues();
  return data.flat().filter(z => z.toString() === zipCode.toString()).length;
}

/**
 * Get contract count for a specific zip code
 */
function getContractCountByZip(zipCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainPage = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);
  if (!mainPage) return 0;

  const data = mainPage.getDataRange().getValues();
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][7].toString() === zipCode.toString() && data[i][15] === 'Yes') {
      count++;
    }
  }

  return count;
}

/**
 * Update all zip code statistics
 */
function updateZipCodeStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zipSheet = ss.getSheetByName(CONFIG.SHEETS.ZIP_CODES);

  if (!zipSheet) return;

  const lastRow = zipSheet.getLastRow();
  if (lastRow < 2) return;

  // Get all zip codes
  const zips = zipSheet.getRange('A2:A' + lastRow).getValues();

  // Update stats for each zip
  for (let i = 0; i < zips.length; i++) {
    const zip = zips[i][0];
    if (!zip) continue;

    const row = i + 2;

    // Set formulas for lead count and contract count
    zipSheet.getRange('I' + row).setFormula(`=COUNTIF('Main Page'!H:H,"${zip}")`);
    zipSheet.getRange('J' + row).setFormula(`=COUNTIFS('Main Page'!H:H,"${zip}",'Main Page'!P:P,"Yes")`);
    zipSheet.getRange('K' + row).setFormula(`=IF(I${row}>0,J${row}/I${row},0)`);
    zipSheet.getRange('M' + row).setValue(new Date());
  }

  SpreadsheetApp.getUi().alert('Zip code statistics updated!');
}

// ============================================================================
// BUYER MATCHING FUNCTIONS
// ============================================================================

/**
 * Find matching buyers for a property
 * @param {string} zipCode
 * @param {string} propertyType
 * @param {number} price
 * @return {Array} Array of matching buyer names
 */
function findMatchingBuyers(zipCode, propertyType, price) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const buyerSheet = ss.getSheetByName(CONFIG.SHEETS.BUYERS_LIST);

  if (!buyerSheet) return [];

  const data = buyerSheet.getDataRange().getValues();
  const matches = [];

  for (let i = 1; i < data.length; i++) {
    const buyer = {
      name: data[i][2],        // Contact Name
      company: data[i][1],     // Company Name
      preferredTypes: data[i][6]?.toString().toLowerCase() || '',
      preferredZips: data[i][7]?.toString() || '',
      minPrice: data[i][8] || 0,
      maxPrice: data[i][9] || Infinity,
      active: data[i][15]
    };

    // Skip inactive buyers
    if (buyer.active !== 'Yes') continue;

    // Check zip code match
    const zipMatch = !buyer.preferredZips ||
                     buyer.preferredZips.includes(zipCode.toString());

    // Check property type match
    const typeMatch = !buyer.preferredTypes ||
                      buyer.preferredTypes.includes(propertyType?.toString().toLowerCase());

    // Check price range
    const priceMatch = (!buyer.minPrice || price >= buyer.minPrice) &&
                       (!buyer.maxPrice || price <= buyer.maxPrice);

    if (zipMatch && typeMatch && priceMatch) {
      matches.push(buyer.company || buyer.name);
    }
  }

  return matches;
}

/**
 * Custom function to show matching buyers in a cell
 */
function MATCHING_BUYERS(zipCode, propertyType, price) {
  const matches = findMatchingBuyers(zipCode, propertyType, price);
  return matches.length > 0 ? matches.join(', ') : 'No matches';
}

// ============================================================================
// EXPORT FUNCTIONS
// ============================================================================

/**
 * Export leads to CSV format
 */
function exportLeadsToCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainPage = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE);

  if (!mainPage) {
    SpreadsheetApp.getUi().alert('Main Page sheet not found');
    return;
  }

  const data = mainPage.getDataRange().getValues();
  let csv = '';

  data.forEach(row => {
    csv += row.map(cell => {
      let value = cell.toString();
      // Escape quotes and wrap in quotes if contains comma
      if (value.includes(',') || value.includes('"')) {
        value = '"' + value.replace(/"/g, '""') + '"';
      }
      return value;
    }).join(',') + '\n';
  });

  // Create download link
  const blob = Utilities.newBlob(csv, 'text/csv', 'leads_export.csv');
  const url = 'data:text/csv;charset=utf-8,' + encodeURIComponent(csv);

  const html = HtmlService.createHtmlOutput(`
    <p>Your export is ready:</p>
    <textarea style="width:100%;height:200px;">${csv}</textarea>
    <p>Copy the above data and paste into a text file, save as .csv</p>
  `).setWidth(500).setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(html, 'Export Complete');
}

/**
 * Export contract pipeline to CSV
 */
function exportPipelineToCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);

  if (!pipeline) {
    SpreadsheetApp.getUi().alert('Contract Pipeline sheet not found');
    return;
  }

  const data = pipeline.getDataRange().getValues();
  let csv = '';

  data.forEach(row => {
    csv += row.map(cell => {
      let value = cell.toString();
      if (value.includes(',') || value.includes('"')) {
        value = '"' + value.replace(/"/g, '""') + '"';
      }
      return value;
    }).join(',') + '\n';
  });

  const html = HtmlService.createHtmlOutput(`
    <p>Your pipeline export is ready:</p>
    <textarea style="width:100%;height:200px;">${csv}</textarea>
    <p>Copy the above data and paste into a text file, save as .csv</p>
  `).setWidth(500).setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(html, 'Pipeline Export Complete');
}

// ============================================================================
// BULK OPERATIONS
// ============================================================================

/**
 * Bulk update status for selected rows
 */
function bulkUpdateStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const selection = sheet.getActiveRange();

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      select, button { padding: 10px; margin: 5px 0; }
      button { background: #1E3A5F; color: white; border: none; cursor: pointer; }
    </style>
    <h3>Bulk Update Status</h3>
    <p>Selected rows: ${selection.getRow()} to ${selection.getLastRow()}</p>
    <p>New Status:</p>
    <select id="status">
      <option value="New">New</option>
      <option value="Contacted">Contacted</option>
      <option value="Analyzing">Analyzing</option>
      <option value="Offer Pending">Offer Pending</option>
      <option value="Under Contract">Under Contract</option>
      <option value="Closed">Closed</option>
      <option value="Dead">Dead</option>
    </select>
    <br><br>
    <button onclick="updateStatus()">Update</button>
    <script>
      function updateStatus() {
        const status = document.getElementById('status').value;
        google.script.run
          .withSuccessHandler(() => {
            alert('Status updated!');
            google.script.host.close();
          })
          .applyBulkStatus(${selection.getRow()}, ${selection.getLastRow()}, status);
      }
    </script>
  `).setWidth(300).setHeight(250);

  SpreadsheetApp.getUi().showModalDialog(html, 'Bulk Update');
}

/**
 * Apply bulk status update
 */
function applyBulkStatus(startRow, endRow, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // Find status column
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('Status') + 1;

  if (statusCol === 0) {
    throw new Error('Status column not found');
  }

  // Update each row
  for (let row = startRow; row <= endRow; row++) {
    sheet.getRange(row, statusCol).setValue(newStatus);
  }
}

/**
 * Bulk mark leads as qualified
 */
function bulkMarkQualified() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const selection = sheet.getActiveRange();

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const qualifiedCol = headers.indexOf('Qualified (Yes/No)') + 1;

  if (qualifiedCol === 0) {
    SpreadsheetApp.getUi().alert('Qualified column not found');
    return;
  }

  const startRow = selection.getRow();
  const endRow = selection.getLastRow();

  for (let row = startRow; row <= endRow; row++) {
    if (row > 1) { // Skip header
      sheet.getRange(row, qualifiedCol).setValue('Yes');
    }
  }

  SpreadsheetApp.getUi().alert(`Marked ${endRow - startRow + 1} leads as Qualified`);
}

// ============================================================================
// QUICK FILTERS
// ============================================================================

/**
 * Show only hot leads (filter)
 */
function filterHotLeads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MAIN_PAGE) || ss.getActiveSheet();

  // Create filter if doesn't exist
  if (!sheet.getFilter()) {
    sheet.getDataRange().createFilter();
  }

  const filter = sheet.getFilter();
  const statusCol = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0].indexOf('Status') + 1;

  if (statusCol > 0) {
    const criteria = SpreadsheetApp.newFilterCriteria()
      .whenTextEqualTo('New')
      .build();
    filter.setColumnFilterCriteria(statusCol, criteria);
  }
}

/**
 * Show deals under contract (filter)
 */
function filterUnderContract() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE) || ss.getActiveSheet();

  if (!sheet.getFilter()) {
    sheet.getDataRange().createFilter();
  }

  const filter = sheet.getFilter();
  const statusCol = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0].indexOf('Status') + 1;

  if (statusCol > 0) {
    const criteria = SpreadsheetApp.newFilterCriteria()
      .whenTextEqualTo('Under Contract')
      .build();
    filter.setColumnFilterCriteria(statusCol, criteria);
  }
}

/**
 * Clear all filters
 */
function clearAllFilters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }

  SpreadsheetApp.getUi().alert('Filters cleared');
}
