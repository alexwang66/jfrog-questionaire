/**
 * JFrog Customer Questionnaire — Google Apps Script Backend
 *
 * Setup steps:
 *  1. Create a new Google Sheet
 *  2. Extensions → Apps Script → paste this file → Save
 *  3. Change SUBMIT_TOKEN below to your own secret string
 *  4. Deploy → New deployment → Web App
 *       Execute as: Me
 *       Who has access: Anyone
 *  5. Copy the Web App URL and paste it into the questionnaire SE Tools panel
 *  6. Paste the same token value into the SE Tools "Submit Token" field
 *
 * Each questionnaire submission appends one row to the active sheet.
 */

// ── Anti-abuse config ──────────────────────────────────────────────────────────
// Change this to any secret string you like. Must match the token configured
// in the SE Tools panel of the questionnaire form.
const SUBMIT_TOKEN   = 'jfrog-se-2026';   // ← change before deploying

// Minimum seconds a user must spend on the form (blocks instant bot submissions)
const MIN_DURATION_S = 10;

// How long (seconds) to block re-submissions from the same email address
const RATE_LIMIT_S   = 600;  // 10 minutes

const SHEET_NAME = 'Responses';

// ── Helper ─────────────────────────────────────────────────────────────────────
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST handler ───────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);

    // 1. Token validation ──────────────────────────────────────────────────────
    if (!d.token || d.token !== SUBMIT_TOKEN) {
      return json({ status: 'rejected', reason: 'invalid token' });
    }

    // 2. Minimum form-fill duration ───────────────────────────────────────────
    const duration = Number(d.durationSeconds) || 0;
    if (duration < MIN_DURATION_S) {
      return json({ status: 'rejected', reason: 'form submitted too quickly' });
    }

    // 3. Required fields ───────────────────────────────────────────────────────
    const name  = d.company?.name?.trim();
    const email = d.company?.contactEmail?.trim();
    if (!name || !email || !email.includes('@')) {
      return json({ status: 'rejected', reason: 'missing required fields' });
    }

    // 4. Rate limiting by email (one submission per 10 minutes) ───────────────
    const cache    = CacheService.getScriptCache();
    const cacheKey = 'sub_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');
    if (cache.get(cacheKey)) {
      return json({ status: 'rejected', reason: 'duplicate submission — please wait before resubmitting' });
    }
    // Mark this email as submitted; key expires after RATE_LIMIT_S seconds
    cache.put(cacheKey, '1', RATE_LIMIT_S);

    // 5. Write to Sheet ────────────────────────────────────────────────────────
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

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
      const header = sheet.getRange(1, 1, 1, 24);
      header.setBackground('#1A1A2E').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

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

    if (sheet.getLastRow() <= 3) {
      sheet.autoResizeColumns(1, 24);
    }

    return json({ status: 'ok', row: sheet.getLastRow() });

  } catch (err) {
    return json({ status: 'error', message: err.message });
  }
}

// Health check — GET request
function doGet() {
  return json({ status: 'ok', message: 'JFrog Questionnaire endpoint is live' });
}
