/**
 * JFrog Customer Questionnaire — Google Apps Script Backend
 *
 * Setup steps:
 *  1. Create a new Google Sheet
 *  2. Extensions → Apps Script → paste this file → Save
 *  3. Deploy → New deployment → Web App
 *       Execute as: Me
 *       Who has access: Anyone
 *  4. Copy the Web App URL and paste it into the questionnaire form
 *
 * Each questionnaire submission appends one row to the active sheet.
 */

const SHEET_NAME = 'Responses'; // Change if needed

function doPost(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet = ss.getSheetByName(SHEET_NAME);

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Write header row on first use
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Submitted At',
        'Company', 'Industry', 'Contact Name', 'Contact Email',
        'Developer Count', 'Dev Count Note',
        'DC Sites', 'Deployment Type',
        'Storage (Current)', 'Existing Tools',
        'Security Tools', 'Compliance Requirements',
        'Package Types (Current)', 'Package Types (Planned)',
        'Pain Points', 'Objectives',
        'Timeline', 'Budget', 'Additional Notes',
        'Fit Score', 'High Priority', 'Medium Priority', 'All Recommendations'
      ]);

      // Style header row
      const header = sheet.getRange(1, 1, 1, 24);
      header.setBackground('#1A1A2E').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const d    = JSON.parse(e.postData.contents);
    const recs = d.recommendations || [];
    const score = recs.length
      ? Math.round(recs.reduce((a, r) => a + r.s, 0) / recs.length)
      : 0;

    const highRecs = recs.filter(r => r.priority === 'high').map(r => r.product).join(', ') || '—';
    const medRecs  = recs.filter(r => r.priority === 'medium').map(r => r.product).join(', ') || '—';
    const allRecs  = recs.map(r => `${r.product} (${r.s})`).join(', ');

    sheet.appendRow([
      new Date(d.submittedAt),
      d.company?.name        || '',
      d.company?.industry    || '',
      d.company?.contactName || '',
      d.company?.contactEmail|| '',
      d.infrastructure?.devCount     || '',
      d.infrastructure?.devCountNote || '',
      (d.infrastructure?.dataCenterSites || []).join(', '),
      d.infrastructure?.deploymentType   || '',
      d.infrastructure?.storageCurrentGB || '',
      (d.infrastructure?.existingTools || []).join(', ') || '',
      (d.security?.toolsInUse || []).join(', '),
      d.security?.compliance || '',
      (d.packages?.currentTypes || []).join(', '),
      (d.packages?.plannedTypes || []).join(', '),
      (d.painPoints?.pains      || []).join(', '),
      (d.painPoints?.objectives || []).join(', '),
      d.painPoints?.timeline || '',
      d.painPoints?.budget   || '',
      d.painPoints?.freeText || '',
      score,
      highRecs,
      medRecs,
      allRecs
    ]);

    // Auto-resize columns on first few rows
    if (sheet.getLastRow() <= 3) {
      sheet.autoResizeColumns(1, 24);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', row: sheet.getLastRow() }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Health check — GET request returns status
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'JFrog Questionnaire endpoint is live' }))
    .setMimeType(ContentService.MimeType.JSON);
}
