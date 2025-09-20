/********* Colibri — Bugfix Pack: Long Notes + Safe Color Scales + Mapping + Sections (CH) *********/
const SHEET_CANDIDATES = {
  MODEL: ['MODEL', 'Financial Model', 'FINANCIAL MODEL', 'FINANCIAL_MODEL'],
  ASSUMPTIONS: ['ASSUMPTIONS', 'Assumptions'],
  SUMMARY: ['SUMMARY', 'Summary'],
  README: ['README', 'Readme', 'READ ME'],
  FORMULA_LIB: ['FORMULA_LIBRARY', 'Formula Library', 'FORMULA LIBRARY', 'Formula_Library'],
  GROWTH: ['Growth Hypothesis', 'GROWTH HYPOTHESIS', 'Growth_Hypothesis'],
  COST: ['cost hypothesis', 'Cost Hypothesis', 'COST HYPOTHESIS', 'cost_hypothesis'],
  MAP: ['MAPPING_CATEGORIES', 'Mapping', 'MAP'],
  IS: ['INCOME_STATEMENT', 'Income Statement', 'P&L'],
  CF: ['CASH_FLOW', 'Cash Flow', 'Cashflow'],
  BS: ['BALANCE_SHEET', 'Balance Sheet', 'BS']
};

/* -------------------- utilities -------------------- */
function normalize_(s) { return String(s || '').toLowerCase().replace(/\s|[_-]+/g, ''); }
function getSheetByCandidates_(cands) {
  const ss = SpreadsheetApp.getActive(), all = ss.getSheets();
  for (const c of cands) { const t = normalize_(c); for (const s of all) { if (normalize_(s.getName()) === t) return s; } }
  for (const s of all) { const n = normalize_(s.getName()); for (const c of cands) { if (n.includes(normalize_(c))) return s; } }
  return null;
}
function getOrInsert_(key, fallback) { return getSheetByCandidates_(SHEET_CANDIDATES[key]) || SpreadsheetApp.getActive().insertSheet(fallback); }
function columnLetter_(col) { let t = ''; while (col > 0) { let r = (col - 1) % 26; t = String.fromCharCode(65 + r) + t; col = Math.floor((col - 1) / 26); } return t; }

// Non-destructive note append (single cell)
function appendNoteCell_(cell, text) {
  const tag = `— Colibri: ${text}`;
  const cur = cell.getNote() || '';
  if (cur.indexOf(tag) !== -1) return;
  cell.setNote(cur ? `${cur}\n\n${tag}` : tag);
}
// Non-destructive for a rectangle
function appendNotesRect_(rng, text, onlyIfHasValue = false) {
  const tag = `— Colibri: ${text}`;
  const vals = rng.getValues(), notes = rng.getNotes();
  for (let r = 0; r < rng.getNumRows(); r++) {
    for (let c = 0; c < rng.getNumColumns(); c++) {
      if (onlyIfHasValue && (vals[r][c] === null || vals[r][c] === '')) continue;
      const cur = notes[r][c] || '';
      if (cur.indexOf(tag) === -1) notes[r][c] = cur ? `${cur}\n\n${tag}` : tag;
    }
  }
  rng.setNotes(notes);
}
function blockNote_(o) {
  return [
    `🧩 What: ${o.what}`,
    `📄 Source: ${o.source}`,
    `🧮 How: ${o.how}`,
    `🇨🇭 Typical in CH: ${o.typical}`,
    `🕹️ When to change: ${o.when}`,
    `⚠️ Warnings: ${o.warn}`
  ].join('\n');
}

// Document-wide settings
function getDocProp_(key, def) {
  try { const v = PropertiesService.getDocumentProperties().getProperty(key); return v === null || v === undefined ? def : v; } catch (e) { return def; }
}
function setDocProp_(key, val) {
  try { PropertiesService.getDocumentProperties().setProperty(key, String(val)); } catch (e) { }
}

function toggleFormulaLibExamples() {
  const key = 'FORMULA_LIB_SHOW_EXAMPLES';
  const cur = getDocProp_(key, 'note'); // 'note' | 'text'
  const next = cur === 'text' ? 'note' : 'text';
  setDocProp_(key, next);
  try { enrichFormulaLib_(); } catch (e) { }
  SpreadsheetApp.getActive().toast(`Formula Library examples now shown as: ${next === 'text' ? 'Visible Text' : 'Notes Only'}`, 'Colibri', 5);
}

/* -------------------- menu -------------------- */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Colibri')
    .addItem('① Normalize Dates & Sections', 'normalizeDatesAndSections')
    .addItem('② Fix Assumptions: Types & Formats', 'fixAssumptionsTypesAndFormats')
    .addItem('③ Build/Refresh MODEL & SUMMARY', 'confirmBuildModelSummary_')
    .addItem('③ Preview (Dry Run): MODEL & SUMMARY', 'previewBuildModelSummary_')
    .addItem('④ Build/Refresh Mapping', 'confirmBuildMapping_')
    .addItem('④ Preview (Dry Run): Mapping', 'previewBuildMapping_')
    .addItem('⑤ Build / Refresh Financial Statements', 'confirmBuildFinancials_')
    .addItem('⑤ Preview (Dry Run): Financial Statements', 'previewBuildFinancials_')
    .addItem('⑥ Append Long Notes (everywhere)', 'appendLongNotesEverywhere')
    .addItem('⑦ Apply MODEL Color Scales (blue)', 'applyModelColorScales')
    .addItem('⑧ EU Formatting & Clean Comments', 'applyEUFormattingAndCleanComments')
    .addItem('⑨ Audit ASSUMPTIONS for Outliers (CH)', 'auditAssumptionsCH')
    .addItem('⑩ Enrich Notes & Guides', 'enrichNotesAndGuides')
    .addItem('⑪ Create Onboarding Walkthrough', 'confirmCreateOnboarding_')
    .addItem('⑪ Preview (Dry Run): Onboarding', 'previewCreateOnboarding_')
    .addItem('⑫ Scan Formula Errors & Report', 'confirmScanFormulaErrors_')
    .addItem('⑫ Preview (Dry Run): Formula Errors Report', 'previewScanFormulaErrors_')
    .addItem('Formula Library: Toggle Examples', 'toggleFormulaLibExamples')
    .addItem('Create Safety Backup (copy file)', 'createSafetyBackupCopy_')
    .addSeparator()
    .addItem('Open Assumptions Coach', 'showAssumptionsCoach')
    .addItem('Quick Fix: Apply Sensible Defaults', 'applySensibleDefaults')
    .addSeparator()
    .addItem('Diagnostics', 'diagnostics')
    ;
  // Coaches submenu
  const coach = ui.createMenu('Coaches 🧠')
    .addItem('ASSUMPTIONS Coach', 'showAssumptionsCoach')
    .addItem('MODEL Coach', 'showModelCoach')
    .addItem('SUMMARY Coach', 'showSummaryCoach')
    .addItem('Income Statement Coach', 'showISCoach')
    .addItem('Cash Flow Coach', 'showCFCoach')
    .addItem('Balance Sheet Coach', 'showBSCoach')
    .addItem('Mapping Coach', 'showMappingCoach')
    .addItem('README Coach', 'showReadmeCoach')
    .addItem('Formula Library Coach', 'showFormulaLibCoach');
  menu.addSubMenu(coach).addToUi();
  SpreadsheetApp.getActive().toast('Colibri menu ready ✅', 'Colibri', 5);

  // If user opens file on ASSUMPTIONS tab, show the coach automatically (gentle)
  try {
    const sh = SpreadsheetApp.getActive().getActiveSheet();
    if (!sh) return;
    const name = normalize_(sh.getName());
    if (SHEET_CANDIDATES.ASSUMPTIONS.some(c => name === normalize_(c))) showAssumptionsCoach();
    else if (SHEET_CANDIDATES.MODEL.some(c => name === normalize_(c))) showModelCoach();
    else if (SHEET_CANDIDATES.SUMMARY.some(c => name === normalize_(c))) showSummaryCoach();
    else if (SHEET_CANDIDATES.IS.some(c => name === normalize_(c))) showISCoach();
    else if (SHEET_CANDIDATES.CF.some(c => name === normalize_(c))) showCFCoach();
    else if (SHEET_CANDIDATES.BS.some(c => name === normalize_(c))) showBSCoach();
    else if (SHEET_CANDIDATES.MAP.some(c => name === normalize_(c))) showMappingCoach();
    else if (SHEET_CANDIDATES.README.some(c => name === normalize_(c))) showReadmeCoach();
    else if (SHEET_CANDIDATES.FORMULA_LIB.some(c => name === normalize_(c))) showFormulaLibCoach();
  } catch (e) { }
}

// Generic confirm + helpers
function runWithConfirm_(title, message) {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(title, message + '\n\nTip: Use “Create Safety Backup (copy file)” in the Colibri menu first if you want a snapshot.', ui.ButtonSet.YES_NO);
  return res === ui.Button.YES;
}
function showRecoveryHint_(affected) {
  const ss = SpreadsheetApp.getActive();
  const hint = `If this wasn’t what you expected:\n• Edit > Undo (Cmd/Ctrl+Z)\n• File > Version history to restore\n• Next time, use Colibri > Create Safety Backup first.\n${affected ? 'Affected: ' + affected : ''}`;
  ss.toast(hint, 'Recovery', 8);
}

function createSafetyBackupCopy_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const file = DriveApp.getFileById(ss.getId());
    const name = `${ss.getName()} — Backup ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}`;
    const copy = file.makeCopy(name);
    const url = `https://docs.google.com/spreadsheets/d/${copy.getId()}/edit`;
    SpreadsheetApp.getUi().alert('Backup created', `A backup copy was created in your Drive:\n\n${name}\n\nOpen: ${url}`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Backup failed', String(e && e.message || e), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Confirm wrappers
function confirmBuildModelSummary_() {
  if (!runWithConfirm_('Rebuild MODEL & SUMMARY?', 'This will overwrite the MODEL and SUMMARY tabs. Proceed?')) return;
  try { buildOrRefreshModelAndSummary(); showRecoveryHint_('MODEL, SUMMARY'); } catch (e) { SpreadsheetApp.getUi().alert('Build failed', String(e && e.message || e), SpreadsheetApp.getUi().ButtonSet.OK); }
}
function confirmBuildMapping_() {
  if (!runWithConfirm_('Rebuild Mapping?', 'This will overwrite the MAP (MAPPING_CATEGORIES) tab. Proceed?')) return;
  try { buildOrRefreshMapping(); showRecoveryHint_('MAP'); } catch (e) { SpreadsheetApp.getUi().alert('Mapping failed', String(e && e.message || e), SpreadsheetApp.getUi().ButtonSet.OK); }
}
function confirmBuildFinancials_() {
  if (!runWithConfirm_('Rebuild Financial Statements?', 'This will overwrite the IS, CF, and BS tabs. Proceed?')) return;
  try { buildFinancialStatements(); showRecoveryHint_('IS, CF, BS'); } catch (e) { SpreadsheetApp.getUi().alert('Financial build failed', String(e && e.message || e), SpreadsheetApp.getUi().ButtonSet.OK); }
}
function confirmCreateOnboarding_() {
  if (!runWithConfirm_('Create/Reset ONBOARDING?', 'This will overwrite the ONBOARDING tab. Proceed?')) return;
  try { createOnboardingWalkthrough(); showRecoveryHint_('ONBOARDING'); } catch (e) { SpreadsheetApp.getUi().alert('Onboarding failed', String(e && e.message || e), SpreadsheetApp.getUi().ButtonSet.OK); }
}
function confirmScanFormulaErrors_() {
  if (!runWithConfirm_('Generate Formula Errors report?', 'This will overwrite the FORMULA_ERRORS tab. Proceed?')) return;
  try { scanFormulaErrorsReport(); showRecoveryHint_('FORMULA_ERRORS'); } catch (e) { SpreadsheetApp.getUi().alert('Scan failed', String(e && e.message || e), SpreadsheetApp.getUi().ButtonSet.OK); }
}

// -------- Dry Run (Preview) helpers --------
function infoForCandidates_(cands) {
  const sh = getSheetByCandidates_(cands);
  if (!sh) return { exists: false, name: `(missing)`, rows: 0, cols: 0 };
  return { exists: true, name: sh.getName(), rows: sh.getLastRow(), cols: sh.getLastColumn() };
}
function previewBuildModelSummary_() {
  const ui = SpreadsheetApp.getUi();
  const a = infoForCandidates_(SHEET_CANDIDATES.MODEL);
  const b = infoForCandidates_(SHEET_CANDIDATES.SUMMARY);
  const msg = [
    'This is a non-destructive preview. No changes were made.',
    '',
    'Affected tabs and actions:',
    `• MODEL — overwrite (clear & rebuild)\n   - Found: ${a.exists ? 'Yes' : 'No'} (name: ${a.name})\n   - Current size: ${a.rows} rows × ${a.cols} cols`,
    `• SUMMARY — overwrite (clear & rebuild)\n   - Found: ${b.exists ? 'Yes' : 'No'} (name: ${b.name})\n   - Current size: ${b.rows} rows × ${b.cols} cols`,
    '',
    'Tip: Use “Create Safety Backup (copy file)” before running the rebuild.'
  ].join('\n');
  ui.alert('Preview — MODEL & SUMMARY Rebuild', msg, ui.ButtonSet.OK);
}
function previewBuildMapping_() {
  const ui = SpreadsheetApp.getUi();
  const m = infoForCandidates_(SHEET_CANDIDATES.MAP);
  const msg = [
    'This is a non-destructive preview. No changes were made.',
    '',
    'Affected tabs and actions:',
    `• MAP (MAPPING_CATEGORIES) — overwrite (clear & rebuild)\n   - Found: ${m.exists ? 'Yes' : 'No'} (name: ${m.name})\n   - Current size: ${m.rows} rows × ${m.cols} cols`,
    '',
    'Tip: Use “Create Safety Backup (copy file)” before running the rebuild.'
  ].join('\n');
  ui.alert('Preview — Mapping Rebuild', msg, ui.ButtonSet.OK);
}
function previewBuildFinancials_() {
  const ui = SpreadsheetApp.getUi();
  const is = infoForCandidates_(SHEET_CANDIDATES.IS);
  const cf = infoForCandidates_(SHEET_CANDIDATES.CF);
  const bs = infoForCandidates_(SHEET_CANDIDATES.BS);
  const msg = [
    'This is a non-destructive preview. No changes were made.',
    '',
    'Affected tabs and actions:',
    `• IS (Income Statement) — overwrite (clear & rebuild)\n   - Found: ${is.exists ? 'Yes' : 'No'} (name: ${is.name})\n   - Current size: ${is.rows} rows × ${is.cols} cols`,
    `• CF (Cash Flow) — overwrite (clear & rebuild)\n   - Found: ${cf.exists ? 'Yes' : 'No'} (name: ${cf.name})\n   - Current size: ${cf.rows} rows × ${cf.cols} cols`,
    `• BS (Balance Sheet) — overwrite (clear & rebuild)\n   - Found: ${bs.exists ? 'Yes' : 'No'} (name: ${bs.name})\n   - Current size: ${bs.rows} rows × ${bs.cols} cols`,
    '',
    'Tip: Use “Create Safety Backup (copy file)” before running the rebuild.'
  ].join('\n');
  ui.alert('Preview — Financial Statements Rebuild', msg, ui.ButtonSet.OK);
}
function previewCreateOnboarding_() {
  const ui = SpreadsheetApp.getUi();
  const ob = infoForCandidates_(['ONBOARDING', 'Onboarding', 'Walkthrough']);
  const msg = [
    'This is a non-destructive preview. No changes were made.',
    '',
    'Affected tabs and actions:',
    `• ONBOARDING — overwrite (clear & rebuild)\n   - Found: ${ob.exists ? 'Yes' : 'No'} (name: ${ob.name})\n   - Current size: ${ob.rows} rows × ${ob.cols} cols`,
    '',
    'Tip: Use “Create Safety Backup (copy file)” before running the rebuild.'
  ].join('\n');
  ui.alert('Preview — Onboarding', msg, ui.ButtonSet.OK);
}
function previewScanFormulaErrors_() {
  const ui = SpreadsheetApp.getUi();
  const fe = infoForCandidates_(['FORMULA_ERRORS', 'Formula Errors', 'Errors']);
  const msg = [
    'This is a non-destructive preview. No changes were made.',
    '',
    'Affected tabs and actions:',
    `• FORMULA_ERRORS — overwrite (clear & rebuild)\n   - Found: ${fe.exists ? 'Yes' : 'No'} (name: ${fe.name})\n   - Current size: ${fe.rows} rows × ${fe.cols} cols`,
    '',
    'Tip: Use “Create Safety Backup (copy file)” before running the rebuild.'
  ].join('\n');
  ui.alert('Preview — Formula Errors Report', msg, ui.ButtonSet.OK);
}

/* -------------------- 9) Audit ASSUMPTIONS for Outliers (CH) -------------------- */
function auditAssumptionsCH() {
  const ass = getOrInsert_('ASSUMPTIONS', 'ASSUMPTIONS');
  const lr = ass.getLastRow(); if (lr < 2) { SpreadsheetApp.getActive().toast('ASSUMPTIONS empty', 'Colibri', 5); return; }
  const labels = ass.getRange(2, 2, lr - 1, 1).getValues().map(r => String(r[0] || '').trim());
  const valsR = ass.getRange(2, 1, lr - 1, 1); const vals = valsR.getValues();
  const auditSh = getOrInsert_('ASSUMPTIONS_AUDIT', 'ASSUMPTIONS_AUDIT'); auditSh.clear();
  const out = [['Row', 'Cell', 'Label', 'Value', 'Status', 'Expected', 'Note']];
  const rules = buildCHRanges_();

  // Reset formats before highlighting
  try { ass.getRange(2, 1, lr - 1, 1).setBackground(null); } catch (e) { /* ignore */ }

  const fmtVal = (v) => {
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (typeof v === 'number') return v;
    return String(v || '');
  };

  for (let i = 0; i < labels.length; i++) {
    try {
      const row = i + 2; const label = labels[i]; const cell = ass.getRange(row, 1); const v = vals[i][0];
      const rule = rules.find(r => r.re.test(label)) || null;
      if (!rule) { continue; }
      let status = '', note = '', expected = '';
      // Type checks
      if (rule.type === 'date') {
        if (!(v instanceof Date)) { status = 'TYPE'; note = 'Expected a Date (yyyy-mm-dd).'; try { cell.setBackground('#FFF3CD'); } catch (e) { } }
      } else if (rule.type === 'percent') {
        if (typeof v !== 'number') { status = 'TYPE'; note = 'Expected a percentage number (0–100).'; try { cell.setBackground('#FFF3CD'); } catch (e) { } }
      } else if (rule.type === 'ratio') {
        if (typeof v !== 'number') { status = 'TYPE'; note = 'Expected a ratio (0–1).'; try { cell.setBackground('#FFF3CD'); } catch (e) { } }
      } else {
        if (typeof v !== 'number') { status = 'TYPE'; note = 'Expected a number (CHF or count).'; try { cell.setBackground('#FFF3CD'); } catch (e) { } }
      }
      // Range checks
      if (typeof v === 'number') {
        expected = rule.min !== undefined && rule.max !== undefined ? `[${rule.min} .. ${rule.max}]` : '';
        if (rule.min !== undefined && v < rule.min) { status = status || 'HARD'; note = note || `Too low for ${rule.ch || 'CH reality'}`; try { cell.setBackground('#F8D7DA'); } catch (e) { } }
        if (rule.max !== undefined && v > rule.max) { status = status || 'HARD'; note = note || `Too high for ${rule.ch || 'CH reality'}`; try { cell.setBackground('#F8D7DA'); } catch (e) { } }
        if (!status && rule.softMin !== undefined && v < rule.softMin) { status = 'SOFT'; note = `Lower than typical ${rule.ch || ''}`; try { cell.setBackground('#FFE8A1'); } catch (e) { } }
        if (!status && rule.softMax !== undefined && v > rule.softMax) { status = 'SOFT'; note = `Higher than typical ${rule.ch || ''}`; try { cell.setBackground('#FFE8A1'); } catch (e) { } }
      }
      if (status) {
        try { appendNoteCell_(cell, `${status === 'HARD' ? '🚫' : '⚠️'} ${note}`); } catch (e) { }
        out.push([row, `A${row}`, label, fmtVal(v), status, expected, rule.desc || '']);
      }
    } catch (err) {
      const row = i + 2; const label = labels[i]; const v = vals[i] ? vals[i][0] : '';
      out.push([row, `A${row}`, label, String(v), 'ERROR', '', String(err && err.message || err)]);
    }
  }
  if (out.length === 1) out.push(['—', '—', '—', '—', 'OK', '—', 'No issues detected']);
  auditSh.getRange(1, 1, out.length, out[0].length).setValues(out);
  auditSh.setFrozenRows(1);
  SpreadsheetApp.getActive().toast('ASSUMPTIONS audited. See ASSUMPTIONS_AUDIT ✅', 'Colibri', 6);
}

function buildCHRanges_() {
  const R = (re, type, o) => ({ re: new RegExp(re, 'i'), type, ...o });
  return [
    // Dates
    R('^Start Date', 'date', { desc: 'Forecast start date' }),
    R('^End Date', 'date', { desc: 'Forecast end date' }),

    // Percentages
    R('Tax rate %', 'percent', { min: 0, max: 100, softMax: 35, ch: 'CH taxes', desc: 'Corporate tax rate %' }),
    R('Cash yield %', 'percent', { min: 0, max: 100, softMax: 5, ch: 'CH deposit yields', desc: 'Cash yield on balances %' }),
    R('Cost of Revenue % \(COGS\)', 'percent', { min: 0, max: 100, softMax: 80, desc: 'COGS share of revenue' }),

    // Ratios (0–1)
    R('Churn .*\(0.?–.?1\)|Churn .*\(0-1\)|Churn monthly', 'ratio', { min: 0, max: 1, softMax: 0.2, desc: 'Monthly churn ratio' }),
    R('Awareness.?→?Conversion .*\(0.?–.?1\)|Awareness.?->?Conversion', 'ratio', { min: 0, max: 1, softMax: 0.3, desc: 'Lead→customer conversion ratio' }),
    R('Other digital monthly growth .*\(0.?–.?1\)|monthly growth .*\(0-1\)', 'ratio', { min: 0, max: 1, softMax: 0.2, desc: 'Other digital monthly growth ratio' }),

    // Salaries (CHF/yr)
    R('Salary L1 .*CHF/yr', 'number', { min: 30000, max: 400000, softMin: 90000, softMax: 220000, ch: 'CH salaries', desc: 'Exec/Architect' }),
    R('Salary L2 .*CHF/yr', 'number', { min: 30000, max: 300000, softMin: 80000, softMax: 180000, ch: 'CH salaries', desc: 'Lead/Owner' }),
    R('Salary L3 .*CHF/yr', 'number', { min: 30000, max: 250000, softMin: 70000, softMax: 150000, ch: 'CH salaries', desc: 'Senior' }),
    R('Salary L4 .*CHF/yr', 'number', { min: 30000, max: 200000, softMin: 60000, softMax: 110000, ch: 'CH salaries', desc: 'Associate' }),
    R('Payroll Overhead %', 'percent', { min: 0, max: 50, softMin: 10, softMax: 25, desc: 'Employer on-top costs %' }),

    // Revenue unit economics
    R('ARPU .*CaaS monthly .*CHF', 'number', { min: 0, max: 50000, softMax: 2500, desc: 'Average monthly revenue per customer' }),
    R('CAC blended .*CHF', 'number', { min: 0, max: 100000, softMax: 25000, desc: 'Customer acquisition cost' }),
    R('LTV months', 'number', { min: 1, max: 120, softMax: 84, desc: 'Lifetime in months' }),

    // Customers and funnel
    R('Initial CaaS customers', 'number', { min: 0, max: 100000, softMax: 5000, desc: 'Starting active customers' }),
    R('Monthly leads', 'number', { min: 0, max: 1000000, softMax: 50000, desc: 'Leads per month' }),

    // AIX pricing
    R('AIX Monthly price .*CHF', 'number', { min: 0, max: 10000, softMax: 2000, desc: 'AIX monthly price' }),
    R('AIX Yearly price .*CHF', 'number', { min: 0, max: 120000, softMax: 24000, desc: 'AIX yearly price' }),

    // Services revenue
    R('Day rate .*training .*CHF', 'number', { min: 200, max: 10000, softMax: 4000, desc: 'Day rate for training' }),
    R('Training days / month', 'number', { min: 0, max: 22, softMax: 15, desc: 'Billable training days per month' }),

    // Budgets (monthly CHF)
    R('Other digital products .*start .*CHF', 'number', { min: 0, max: 1000000, softMax: 50000, desc: 'Other digital start revenue' }),
    R('Culture .*Learning per FTE .*CHF', 'number', { min: 0, max: 2000, softMax: 400, desc: 'Learning per FTE / month' }),
    R('Tech stack per FTE .*CHF', 'number', { min: 0, max: 2000, softMax: 400, desc: 'Tools per FTE / month' }),
    R('Labs .*Universities monthly .*CHF', 'number', { min: 0, max: 100000, softMax: 10000, desc: 'Labs & Universities monthly' }),
    R('R&D AIX monthly .*CHF', 'number', { min: 0, max: 500000, softMax: 50000, desc: 'AIX R&D monthly' }),
    R('Product dev .*CaaS.* monthly .*CHF', 'number', { min: 0, max: 500000, softMax: 50000, desc: 'CaaS product dev monthly' }),
    R('Media / PR monthly .*CHF', 'number', { min: 0, max: 1000000, softMax: 50000, desc: 'Media/PR monthly' }),
    R('Purchased Services .*Micro monthly .*CHF', 'number', { min: 0, max: 500000, softMax: 20000, desc: 'Purchased services Micro' }),
    R('Purchased Services .*Meso monthly .*CHF', 'number', { min: 0, max: 500000, softMax: 50000, desc: 'Purchased services Meso' }),
    R('Purchased Services .*Macro monthly .*CHF', 'number', { min: 0, max: 500000, softMax: 100000, desc: 'Purchased services Macro' }),
    R('Purchased Services .*Mundo monthly .*CHF', 'number', { min: 0, max: 500000, softMax: 150000, desc: 'Purchased services Mundo' }),

    // Cash and timing
    R('Starting Cash .*CHF', 'number', { min: 0, max: 50000000, softMax: 5000000, desc: 'Starting cash balance' }),
    R('Depreciation months', 'number', { min: 1, max: 120, softMax: 60, desc: 'Depreciation period in months' }),
    R('DSO|DPO|DIO', 'number', { min: 0, max: 180, softMax: 90, desc: 'Working capital timing (days)' }),
  ];
}

/* -------------------- 1) Dates + Sections on ASSUMPTIONS -------------------- */
function normalizeDatesAndSections() {
  const ass = getOrInsert_('ASSUMPTIONS', 'ASSUMPTIONS');
  // ensure columns A/B/C exist
  ass.getRange('A1').setValue('Value');
  ass.getRange('B1').setValue('Assumption');
  ass.getRange('C1').setValue('Source / Notes');
  // Ensure Section column (D)
  ass.getRange('D1').setValue('Section');

  // Start/End base
  ensureAssumptionValueWithNote_(ass, 'Start Date (yyyy-mm-dd)', new Date(2025, 10, 1), {
    what: 'Start month of forecast.',
    source: 'Timeline for MODEL.',
    how: 'Month_1 = Start; next months use EDATE(prev,1).',
    typical: 'First of any month.',
    when: 'Change to shift the entire horizon.',
    warn: 'Blank or invalid date breaks the timeline.'
  });
  const startRow = findLabelRow_(ass, 'Start Date (yyyy-mm-dd)');
  const start = ass.getRange(startRow, 1).getValue();
  const end = new Date(start); end.setMonth(end.getMonth() + 72);
  ensureAssumptionValueWithNote_(ass, 'End Date (yyyy-mm-dd)', end, {
    what: 'Last forecast month.',
    source: 'Timeline for MODEL.',
    how: 'In_Horizon = 1 between Start and End.',
    typical: '5–7 years.',
    when: 'Extend for longer planning.',
    warn: 'End < Start truncates series.'
  });

  // Section tagging based on your flow
  const sectionRules = buildSectionRules_();
  const last = ass.getLastRow(); if (last < 2) return;
  const labels = ass.getRange(2, 2, last - 1, 1).getValues().map(r => String(r[0] || '').trim());
  const sectCol = ass.getRange(2, 4, last - 1, 1);
  const sectVals = sectCol.getValues();
  for (let i = 0; i < labels.length; i++) {
    sectVals[i][0] = sectionForLabel_(labels[i], sectionRules);
  }
  sectCol.setValues(sectVals);

  // Colour band sections for scannability (pastel backgrounds)
  const colors = {
    'Customer Journey': '#FFF7E6',
    'Revenue & Customer Metrics': '#E8F4FD',
    'CAC & Efficiency': '#FFF9C4',
    'Revenue → Profit': '#E8F5E9',
    'Margins': '#F3E5F5',
    'Cash Flow & Balance Sheet': '#F1F8E9',
    'Investments': '#EDE7F6',
    'Other': '#F5F5F5'
  };
  for (let i = 0; i < labels.length; i++) {
    const clr = colors[sectVals[i][0]] || '#FFFFFF';
    ass.getRange(2 + i, 1, 1, 4).setBackground(clr);
  }
  ass.setFrozenRows(1);
  appendNoteCell_(ass.getRange('D1'), '📚 Logical grouping that mirrors your flowchart.');
  SpreadsheetApp.getActive().toast('Dates normalised + Sections applied ✅', 'Colibri', 5);
}
function buildSectionRules_() {
  return [
    { section: 'Customer Journey', keys: ['Monthly leads', 'Awareness→Conversion'] },
    { section: 'Revenue & Customer Metrics', keys: ['ARPU', 'LTV months', 'Gross Margin %', 'Churn', 'Initial CaaS customers', 'AIX Monthly price', 'AIX Yearly price'] },
    { section: 'CAC & Efficiency', keys: ['CAC blended'] },
    { section: 'Revenue → Profit', keys: ['Cost of Revenue % (COGS)', 'AI Agents cost per active customer per month (CHF)', 'Day rate – training', 'Training days', 'Other digital'] },
    { section: 'Margins', keys: ['Gross Margin %'] },
    { section: 'Cash Flow & Balance Sheet', keys: ['Starting Cash', 'DSO', 'DPO', 'DIO'] },
    { section: 'Investments', keys: ['CapEx', 'Depreciation'] },
  ];
}
function sectionForLabel_(label, rules) {
  const s = String(label || '');
  for (const r of rules) {
    if (r.keys.some(k => s.indexOf(k) > -1)) return r.section;
  }
  return 'Other';
}
function ensureAssumptionValueWithNote_(ass, label, value, note) {
  const r = findOrCreateLabel_(ass, label);
  if (!ass.getRange(r, 1).getValue()) ass.getRange(r, 1).setValue(value);
  appendNoteCell_(ass.getRange(r, 2), blockNote_(note));
  appendNoteCell_(ass.getRange(r, 1), '🟢 Edit here. This value drives the model.');
}
function findLabelRow_(ass, label) {
  const last = ass.getLastRow(); if (last < 2) return null;
  const labels = ass.getRange(2, 2, last - 1, 1).getValues().map(r => String(r[0] || ''));
  const idx = labels.findIndex(x => x === label);
  return idx === -1 ? null : (2 + idx);
}
function findOrCreateLabel_(ass, label) {
  let r = findLabelRow_(ass, label);
  if (r) return r;
  const nr = ass.getLastRow() + 1; ass.getRange(nr, 2).setValue(label); return nr;
}

/* -------------------- 2) Mapping (drivers → categories & bucket) -------------------- */
function buildOrRefreshMapping() {
  const ass = getOrInsert_('ASSUMPTIONS', 'ASSUMPTIONS');
  const mapSh = getOrInsert_('MAP', 'MAPPING_CATEGORIES'); mapSh.clear();
  const header = ['Source', 'Label', 'Category', 'Bucket', 'Notes'];
  const rows = [header];

  // Pull from driver tabs if they exist
  const growth = readKV_(getSheetByCandidates_(SHEET_CANDIDATES.GROWTH));
  const cost = readKV_(getSheetByCandidates_(SHEET_CANDIDATES.COST));
  const haveDriverTabs = growth.length || cost.length;

  // If missing, seed from ASSUMPTIONS cost-related labels
  const fallbackLabels = [
    'Cost of Revenue % (COGS)', 'AI Agents cost per active customer per month (CHF)',
    'Media / PR monthly (CHF)', 'R&D AIX monthly (CHF)', 'Product dev (CaaS) monthly (CHF)',
    'Tech stack per FTE / month (CHF)', 'Culture & Learning per FTE / month (CHF)',
    'Labs & Universities monthly (CHF)',
    'Purchased Services – Micro monthly (CHF)', 'Purchased Services – Meso monthly (CHF)', 'Purchased Services – Macro monthly (CHF)', 'Purchased Services – Mundo monthly (CHF)'
  ];
  const fromAss = labelsFromAssumptions_(ass, fallbackLabels);

  const all = haveDriverTabs
    ? [...growth.map(([k]) => ['Growth Hypothesis', k]), ...cost.map(([k]) => ['cost hypothesis', k])]
    : fromAss.map(k => ['ASSUMPTIONS', k]);

  const defaultCatBucket = labelToDefaultCatBucket_;
  all.forEach(([src, label]) => {
    const { cat, bucket } = defaultCatBucket(label);
    rows.push([src, label, cat, bucket, '']);
  });

  mapSh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  const cats = ['Hosting/Infra', 'AI Runtime', 'Payment Fees', 'Marketing/Media', 'R&D / Product', 'Tech & Tools', 'Services & Labs', 'Culture & Learning', 'Other'];
  const buckets = ['COGS', 'Opex'];
  mapSh.getRange(2, 3, rows.length - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(cats, true).build());
  mapSh.getRange(2, 4, rows.length - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(buckets, true).build());
  mapSh.setFrozenRows(1);

  appendNoteCell_(mapSh.getRange('A1'), '✍️ Edit Category/Bucket per line. This drives IS split across COGS & Opex buckets.');
  SpreadsheetApp.getActive().toast('MAPPING_CATEGORIES filled ✅', 'Colibri', 5);
}
function readKV_(sheet) {
  if (!sheet) return [];
  const lr = sheet.getLastRow(), lc = sheet.getLastColumn(); if (lr < 2) return [];
  const hdrs = sheet.getRange(1, 1, 1, Math.min(lc, 6)).getValues()[0].map(x => String(x || '').toLowerCase());
  let colKey = 1, colVal = 2; const keyH = ['label', 'metric', 'assumption', 'name', 'kpi']; const valH = ['value', 'amount', 'chf', 'val', 'number'];
  for (let c = 1; c <= Math.min(lc, 6); c++) { const h = hdrs[c - 1]; if (keyH.some(k => h.indexOf(k) > -1)) colKey = c; if (valH.some(k => h.indexOf(k) > -1)) colVal = c; }
  const data = sheet.getRange(2, 1, lr - 1, Math.max(colKey, colVal)).getValues();
  const out = []; for (const row of data) { const k = String(row[colKey - 1] || '').trim(); const v = row[colVal - 1]; if (k) out.push([k, v]); }
  return out;
}
function labelsFromAssumptions_(ass, want) {
  const last = ass.getLastRow(); if (last < 2) return [];
  const labels = ass.getRange(2, 2, last - 1, 1).getValues().map(r => String(r[0] || '').trim());
  return labels.filter(l => want.some(w => l.indexOf(w) > -1) || /(monthly|per FTE|Purchased Services|COGS|AI Agents|Media|R&D|Product dev|Tech stack|Labs)/i.test(l));
}
function labelToDefaultCatBucket_(label) {
  const l = String(label || '').toLowerCase();
  if (/(hosting|cloud|infra|server|cdn)/.test(l)) return { cat: 'Hosting/Infra', bucket: 'COGS' };
  if (/ai agent|ai runtime|ai cost|ai agents/.test(l)) return { cat: 'AI Runtime', bucket: 'COGS' };
  if (/payment|stripe|fees/.test(l)) return { cat: 'Payment Fees', bucket: 'COGS' };
  if (/media|marketing|ads|advert|pr/.test(l)) return { cat: 'Marketing/Media', bucket: 'Opex' };
  if (/r&d|research|aix|product dev/.test(l)) return { cat: 'R&D / Product', bucket: 'Opex' };
  if (/tech stack|tools|saas|license|software|cloud/.test(l)) return { cat: 'Tech & Tools', bucket: 'Opex' };
  if (/labs|universit/.test(l)) return { cat: 'Services & Labs', bucket: 'Opex' };
  if (/culture .*learning/.test(l)) return { cat: 'Culture & Learning', bucket: 'Opex' };
  if (/purchased services/.test(l)) return { cat: 'Services & Labs', bucket: 'Opex' };
  return { cat: 'Other', bucket: 'Opex' };
}

/* -------------------- 3) Build IS/CF/BS (reads mapping) -------------------- */
function buildFinancialStatements() {
  const model = getSheetByCandidates_(SHEET_CANDIDATES.MODEL);
  if (!model) throw new Error('MODEL tab not found.');

  // MODEL headers
  const headers = model.getRange(1, 1, 1, model.getLastColumn()).getValues()[0];
  const idx = n => { const i = headers.indexOf(n); if (i < 0) throw new Error(`MODEL header not found: ${n}`); return i + 1; };
  const cMonth = idx('Month'), cInHor = idx('In_Horizon'), cRev = idx('Revenue_Display'), cCosts = idx('Costs_Display'), cCash = idx('Cum_Cash');

  // Horizon rows
  const hv = model.getRange(2, cInHor, model.getLastRow() - 1, 1).getValues();
  let last = 1; for (let i = 0; i < hv.length; i++) if (hv[i][0] === 1 || hv[i][0] === '1') last = i + 1;
  const rows = Math.max(last, 1), lastCol = 1 + rows;
  const dates = model.getRange(2, cMonth, rows, 1).getValues();

  // Mapping
  const mapSh = getOrInsert_('MAP', 'MAPPING_CATEGORIES');
  const mLR = mapSh.getLastRow();
  const mapRows = mLR >= 2 ? mapSh.getRange(2, 1, mLR - 1, 4).getValues() : [];
  const catAgg = { 'Hosting/Infra': 0, 'AI Runtime': 0, 'Payment Fees': 0, 'Marketing/Media': 0, 'R&D / Product': 0, 'Tech & Tools': 0, 'Services & Labs': 0, 'Culture & Learning': 0, 'Other': 0 };
  mapRows.forEach(r => {
    const label = String(r[1] || '').trim(), cat = r[2] || 'Other', bucket = r[3] || 'Opex';
    // Pull amounts from ASSUMPTIONS if present
    const ass = getOrInsert_('ASSUMPTIONS', 'ASSUMPTIONS');
    const row = findLabelRow_(ass, label);
    const val = (row ? ass.getRange(row, 1).getValue() : 0) || 0;
    // Only aggregate numeric amounts (ignores % labels like COGS%)
    if (typeof val === 'number') catAgg[cat] = (catAgg[cat] || 0) + val;
  });

  /* ----- INCOME STATEMENT ----- */
  const is = getOrInsert_('IS', 'INCOME_STATEMENT'); is.clear();
  const hdr = ['Metric', ...dates.map(d => d[0])];
  const lines = [
    'Revenue (Total) 💵',                // 2
    'COGS – Hosting/Infra 🧪',          // 3
    'COGS – AI Runtime 🤖',             // 4
    'COGS – Payment Fees 💳',           // 5
    'COGS (Direct Total) 🧪',           // 6
    'Gross Profit 💎',                  // 7
    'Gross Margin % 📈',                // 8
    'Opex – Marketing/Media 📢',        // 9
    'Opex – R&D / Product 🔬🏗️',        // 10
    'Opex – Tech & Tools 🧰',           // 11
    'Opex – Services & Labs 🛒🎓',      // 12
    'Opex – Culture & Learning 🌱',     // 13
    'Opex – Other 🗂️',                 // 14
    'Opex (Total) 🧰',                  // 15
    'EBITDA 📊',                        // 16
    'Depreciation 🧱',                  // 17
    'EBIT 🧮',                          // 18
    'Interest (net) 💳',                // 19
    'EBT 💼',                           // 20
    'Taxes 🧾',                         // 21
    'Net Income 🟢'                     // 22
  ];
  is.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  for (let r = 0; r < lines.length; r++) is.getRange(2 + r, 1).setValue(lines[r]);
  for (let c = 2; c <= lastCol; c++) {
    const REV = `IFERROR(INDEX(MODEL!${columnLetter_(cRev)}:${columnLetter_(cRev)}, MATCH(${is.getRange(1, c).getA1Notation()}, MODEL!${columnLetter_(cMonth)}:${columnLetter_(cMonth)}, 0)),0)`;
    is.getRange(2, c).setFormula(`=${REV}`);
    // COGS buckets (flat monthly, mapping-driven)
    is.getRange(3, c).setValue(catAgg['Hosting/Infra'] || 0);
    is.getRange(4, c).setValue(catAgg['AI Runtime'] || 0);
    is.getRange(5, c).setValue(catAgg['Payment Fees'] || 0);
    is.getRange(6, c).setFormula(`=SUM(${is.getRange(3, c).getA1Notation()}:${is.getRange(5, c).getA1Notation()})`);
    // GP / GM%
    is.getRange(7, c).setFormula(`=${is.getRange(2, c).getA1Notation()}-${is.getRange(6, c).getA1Notation()}`);
    is.getRange(8, c).setFormula(`=IFERROR(${is.getRange(7, c).getA1Notation()}/MAX(0.0001,${is.getRange(2, c).getA1Notation()}),0)`);
    // Opex buckets
    is.getRange(9, c).setValue(catAgg['Marketing/Media'] || 0);
    is.getRange(10, c).setValue(catAgg['R&D / Product'] || 0);
    is.getRange(11, c).setValue(catAgg['Tech & Tools'] || 0);
    is.getRange(12, c).setValue(catAgg['Services & Labs'] || 0);
    is.getRange(13, c).setValue(catAgg['Culture & Learning'] || 0);
    is.getRange(14, c).setValue(catAgg['Other'] || 0);
    is.getRange(15, c).setFormula(`=SUM(${is.getRange(9, c).getA1Notation()}:${is.getRange(14, c).getA1Notation()})`);
    // EBITDA / Dep / EBIT
    is.getRange(16, c).setFormula(`=${is.getRange(7, c).getA1Notation()}-${is.getRange(15, c).getA1Notation()}`);
    is.getRange(17, c).setFormula(`=IFERROR(INDEX(ASSUMPTIONS!A:A, MATCH("CapEx per month (CHF)", ASSUMPTIONS!B:B,0))/MAX(1, INDEX(ASSUMPTIONS!A:A, MATCH("Depreciation months", ASSUMPTIONS!B:B,0))), 0)`);
    is.getRange(18, c).setFormula(`=${is.getRange(16, c).getA1Notation()}-${is.getRange(17, c).getA1Notation()}`);
    // Interest (cash yield only, no debt schedule)
    const CF = getOrInsert_('CF', 'CASH_FLOW');
    const CASH_t = CF.getRange(10, c).getA1Notation();
    const CASH_prev = c > 2 ? CF.getRange(10, c - 1).getA1Notation() : CASH_t;
    const AVG_CASH = `(${CASH_t}+${CASH_prev})/2`;
    const CASH_INT = `IFERROR(${AVG_CASH} * INDEX(ASSUMPTIONS!A:A, MATCH("Cash yield %", ASSUMPTIONS!B:B,0))/100/12, 0)`;
    is.getRange(19, c).setFormula(`=${CASH_INT}`);
    is.getRange(20, c).setFormula(`=${is.getRange(18, c).getA1Notation()}+${is.getRange(19, c).getA1Notation()}`);
    is.getRange(21, c).setFormula(`=IFERROR(MAX(0, ${is.getRange(20, c).getA1Notation()} * IFERROR(INDEX(ASSUMPTIONS!A:A, MATCH("Tax rate %", ASSUMPTIONS!B:B,0))/100,0)),0)`);
    is.getRange(22, c).setFormula(`=IFERROR(${is.getRange(20, c).getA1Notation()}-${is.getRange(21, c).getA1Notation()},0)`);
  }

  /* ----- CASH FLOW ----- */
  const cf = getOrInsert_('CF', 'CASH_FLOW'); cf.clear();
  const cfHdr = ['Metric', ...dates.map(d => d[0])];
  const cfLines = ['Net Income 🧾', 'Non-cash: Depreciation 🧱', 'Working Capital Δ 🔄', 'Operating Cash Flow 💧', 'CapEx 🛠️', 'Free Cash Flow 🟢', 'Financing (Debt/Equity) 💳', 'Net Cash Change 💱', 'Ending Cash 🏦'];
  cf.getRange(1, 1, 1, cfHdr.length).setValues([cfHdr]);
  cfLines.forEach((L, i) => cf.getRange(2 + i, 1).setValue(L));
  for (let c = 2; c <= lastCol; c++) {
    cf.getRange(2, c).setFormula(`=IFERROR(${is.getRange(22, c).getA1Notation()},0)`);
    cf.getRange(3, c).setFormula(`=IFERROR(${is.getRange(17, c).getA1Notation()},0)`);
    // keep WC simple (0) unless you enable DSO/DPO/DIO later
    cf.getRange(4, c).setValue(0);
    cf.getRange(5, c).setFormula(`=${cf.getRange(2, c).getA1Notation()}+${cf.getRange(3, c).getA1Notation()}+${cf.getRange(4, c).getA1Notation()}`);
    cf.getRange(6, c).setFormula(`=IFERROR(INDEX(ASSUMPTIONS!A:A, MATCH("CapEx per month (CHF)", ASSUMPTIONS!B:B,0)),0)`);
    cf.getRange(7, c).setFormula(`=${cf.getRange(5, c).getA1Notation()}-${cf.getRange(6, c).getA1Notation()}`);
    cf.getRange(8, c).setValue(0);
    cf.getRange(9, c).setFormula(`=${cf.getRange(7, c).getA1Notation()}+${cf.getRange(8, c).getA1Notation()}`);
    const CASH_MODEL = `INDEX(MODEL!${columnLetter_(cCash)}:${columnLetter_(cCash)}, MATCH(${cf.getRange(1, c).getA1Notation()}, MODEL!${columnLetter_(cMonth)}:${columnLetter_(cMonth)},0))`;
    cf.getRange(10, c).setFormula(`=${CASH_MODEL}`);
  }

  /* ----- BALANCE SHEET ----- */
  const bs = getOrInsert_('BS', 'BALANCE_SHEET'); bs.clear();
  const bsHdr = ['Metric', ...dates.map(d => d[0])];
  const bsLines = ['Cash 🏦', 'A/R 📬', 'Inventory 📦', 'Prepaids & Other 🗃️', 'PP&E (net) 🏭', 'Total Assets 💼', 'A/P 🧾', 'Other Liab 📑', 'Deferred Rev ⏳', 'Debt 💳', 'Total Liabilities 🧮', 'Equity 📈', 'Liabilities + Equity ⚖️', 'Balance Check ✅'];
  bs.getRange(1, 1, 1, bsHdr.length).setValues([bsHdr]);
  bsLines.forEach((L, i) => bs.getRange(2 + i, 1).setValue(L));
  for (let c = 2; c <= lastCol; c++) {
    // Cash mirrors CF Ending Cash (row 10) at matching date; use explicit sheet name
    bs.getRange(2, c).setFormula(`=${cf.getSheetName()}!${cf.getRange(10, c).getA1Notation()}`);
    bs.getRange(3, c).setValue(0);
    bs.getRange(4, c).setValue(0);
    const prev = c > 2 ? bs.getRange(5, c - 1).getA1Notation() : '0';
    // PP&E (net) = prev PP&E + CapEx - Depreciation; use explicit sheet names to avoid self-references
    bs.getRange(5, c).setFormula(`=IFERROR(${prev}+IFERROR(${cf.getSheetName()}!${cf.getRange(6, c).getA1Notation()},0)-IFERROR(${is.getSheetName()}!${is.getRange(17, c).getA1Notation()},0),0)`);
    bs.getRange(6, c).setFormula(`=IFERROR(SUM(${bs.getRange(2, c).getA1Notation()}:${bs.getRange(5, c).getA1Notation()}),0)`);
    bs.getRange(7, c).setValue(0);
    bs.getRange(8, c).setValue(0);
    bs.getRange(9, c).setValue(0);
    bs.getRange(10, c).setValue(0);
    bs.getRange(11, c).setFormula(`=IFERROR(SUM(${bs.getRange(7, c).getA1Notation()}:${bs.getRange(10, c).getA1Notation()}),0)`);
    bs.getRange(12, c).setFormula(`=IFERROR(${bs.getRange(6, c).getA1Notation()}-${bs.getRange(11, c).getA1Notation()},0)`);
    bs.getRange(13, c).setFormula(`=IFERROR(${bs.getRange(11, c).getA1Notation()}+${bs.getRange(12, c).getA1Notation()},0)`);
    bs.getRange(14, c).setFormula(`=IFERROR(${bs.getRange(6, c).getA1Notation()}-${bs.getRange(13, c).getA1Notation()},0)`);
    // Row 15: Balance Check = Total Assets - (Liabilities + Equity)
    if (bs.getLastRow() < 15) bs.insertRowsAfter(14, 1);
    bs.getRange(15, 1).setValue('Balance Check ✅');
    bs.getRange(15, c).setFormula(`=IFERROR(${bs.getRange(6, c).getA1Notation()}-${bs.getRange(13, c).getA1Notation()},0)`);
  }

  SpreadsheetApp.getActive().toast('IS/CF/BS rebuilt ✅', 'Colibri', 5);
}

/* -------------------- 4) Long notes (titles + all editable values) -------------------- */
function appendLongNotesEverywhere() {
  // ASSUMPTIONS — headers
  const ass = getOrInsert_('ASSUMPTIONS', 'ASSUMPTIONS');
  appendNoteCell_(ass.getRange('A1'), '🔢 Enter numbers/dates. This is the only column you usually edit.');
  appendNoteCell_(ass.getRange('B1'), '🏷️ Labels. Formulas look up these exact names.');
  appendNoteCell_(ass.getRange('C1'), '🔗 Sources / rationale. Paste links here.');
  appendNoteCell_(ass.getRange('D1'), '📚 Section for scannability (matches your flowchart).');

  // ASSUMPTIONS — per-value notes (full 6-part). Known labels map below; others get a generic helper.
  const notes = assumptionNotesCH_();
  const last = ass.getLastRow(); if (last >= 2) {
    const labR = ass.getRange(2, 2, last - 1, 1).getValues(), valR = ass.getRange(2, 1, last - 1, 1);
    const valNotes = valR.getNotes();
    for (let r = 0; r < labR.length; r++) {
      const lab = String(labR[r][0] || '').trim();
      const n = notes[lab] || notes.__generic(lab);
      const tag = `— Colibri: ${blockNote_(n)}`;
      const cur = valNotes[r][0] || '';
      if (cur.indexOf(tag) === -1) valNotes[r][0] = cur ? `${cur}\n\n${tag}` : tag;
    }
    valR.setNotes(valNotes);
  }

  // MODEL / SUMMARY / IS / CF / BS — header title notes
  headerTitleNotes_();

  // Formula library short helper notes
  const fl = getOrInsert_('FORMULA_LIB', 'FORMULA_LIBRARY');
  appendNoteCell_(fl.getRange('A1'), '📚 Handy functions used in this model.');
  SpreadsheetApp.getActive().toast('Long notes appended across tabs ✅', 'Colibri', 5);
}
function assumptionNotesCH_() {
  const A = (t) => `INDEX(ASSUMPTIONS!A:A, MATCH("${t}", ASSUMPTIONS!B:B, 0))`;
  return {
    'Starting Cash (CHF)': {
      what: 'Money in bank at the start month.',
      source: 'Cash Flow (begin) & Balance Sheet (Cash).',
      how: 'Used to seed MODEL Cum_Cash;\nRunway = Starting Cash ÷ Latest Net Burn.',
      typical: 'CHF 50k–500k seed stage (varies).',
      when: 'Change when you raise/spend before start.',
      warn: '0 with positive burn → runway 0.'
    },
    'Payroll Overhead %': {
      what: 'Employer on-top costs (AHV/ALV/BVG/accident).',
      source: 'Income Statement (Opex).',
      how: 'Payroll/mo = (Σ HC×Salary ÷12)×(1+Overhead%).',
      typical: '~12–22% depending on benefits.',
      when: 'Adjust per benefits set.',
      warn: 'Too low inflates profits.'
    },
    'Salary L1 (CHF/yr)': { what: 'Yearly salary Level 1 (Exec/Architect).', source: 'Opex (Payroll).', how: 'Part of Σ HC×Salary ÷12.', typical: 'CHF 140k–220k+', when: 'Adjust to actual contracts.', warn: 'Outliers skew payroll.' },
    'Salary L2 (CHF/yr)': { what: 'Yearly salary Level 2 (Lead/Owner).', source: 'Opex (Payroll).', how: 'Part of Σ HC×Salary ÷12.', typical: 'CHF 110k–160k', when: 'Adjust to actual.', warn: '—' },
    'Salary L3 (CHF/yr)': { what: 'Yearly salary Level 3 (Senior).', source: 'Opex (Payroll).', how: 'Part of Σ HC×Salary ÷12.', typical: 'CHF 90k–130k', when: 'Adjust to actual.', warn: '—' },
    'Salary L4 (CHF/yr)': { what: 'Yearly salary Level 4 (Associate).', source: 'Opex (Payroll).', how: 'Part of Σ HC×Salary ÷12.', typical: 'CHF 60k–95k', when: 'Adjust to actual.', warn: '—' },
    'Starting Headcount': { what: 'People on payroll at start.', source: 'Opex (Payroll).', how: 'Drives Σ HC in MODEL.', typical: 'Founders 2–4 + early hires.', when: 'Update as you hire.', warn: '—' },
    'Target Headcount (by year 3)': { what: 'Goal headcount ~36 months.', source: 'Opex (Payroll).', how: 'MODEL ramps towards this.', typical: 'Context-specific.', when: 'Update as plans evolve.', warn: '—' },
    'Cost of Revenue % (COGS)': {
      what: 'Direct costs share of revenue.',
      source: 'Income Statement (COGS).',
      how: `COGS = Revenue × ${A('Cost of Revenue % (COGS)')}/100.`,
      typical: 'SaaS net GM 60–90%.',
      when: 'Update with infra/support data.',
      warn: '0% likely unrealistic.'
    },
    'AI Agents cost per active customer per month (CHF)': {
      what: 'AI runtime CHF per active customer.',
      source: 'COGS (direct).',
      how: 'COGS += Active_Customers × AI_cost_per_customer.',
      typical: 'Highly variable (usage).',
      when: 'Update with observed usage.',
      warn: 'If 0, GM% can look too high.'
    },
    'ARPU – CaaS monthly (CHF)': {
      what: 'Average revenue per customer per month.',
      source: 'Income Statement (Revenue).',
      how: 'MRR = Active_Customers × ARPU.',
      typical: 'CHF 100–600 B2B (varies).',
      when: 'Change when repricing.',
      warn: '0 → zero MRR.'
    },
    'Churn monthly (0–1)': {
      what: 'Monthly % of customers that cancel.',
      source: 'Revenue dynamics.',
      how: 'Active_t = Active_{t-1}×(1−churn)+New.',
      typical: '1–5% B2B (enterprise <1%).',
      when: 'Update with real data.',
      warn: 'High churn kills LTV.'
    },
    'Awareness→Conversion (0–1)': {
      what: '% of leads that become paying customers.',
      source: 'Top-of-funnel.',
      how: 'New = Leads × Conversion.',
      typical: '1–5% cold; higher warm.',
      when: 'Update from funnel metrics.',
      warn: '0 → no growth.'
    },
    'CAC blended (CHF)': {
      what: 'Average cost to acquire one customer.',
      source: 'Unit economics.',
      how: 'CAC Payback = CAC ÷ (ARPU × GM%).',
      typical: 'CHF 1k–8k+ (B2B).',
      when: 'Update from real campaigns.',
      warn: 'Huge CAC + low ARPU is bad.'
    },
    'LTV months': {
      what: 'Typical months a customer stays.',
      source: 'Unit economics.',
      how: 'LTV = ARPU × months × GM%.',
      typical: '12–60 based on segment.',
      when: 'Update as retention matures.',
      warn: 'Too low harms LTV/CAC.'
    },
    'Gross Margin %': {
      what: '% kept after direct costs.',
      source: 'P&L.',
      how: 'GM% = GrossProfit ÷ Revenue.',
      typical: 'SaaS 60–90%.',
      when: 'Derived; no manual change.',
      warn: 'If forcing manual, beware.'
    },
    'AIX Monthly price (CHF)': {
      what: 'Price per month for AIX content.',
      source: 'Revenue.',
      how: 'AIX_Monthly = subs × price.',
      typical: 'Context-specific.',
      when: 'Change when repricing.',
      warn: '—'
    },
    'AIX Yearly price (CHF)': {
      what: 'Yearly price (recognized monthly /12).',
      source: 'Revenue + Deferred Rev.',
      how: 'Monthly recog = yearly/12.',
      typical: 'Context-specific.',
      when: 'Change when repricing.',
      warn: '—'
    },
    'Initial CaaS customers': {
      what: 'Customers at month 1.',
      source: 'Revenue dynamics.',
      how: 'Seed Active_Customers.',
      typical: '0–20 early stage.',
      when: 'Set actual count.',
      warn: '—'
    },
    'Monthly leads': {
      what: 'Leads entering funnel per month.',
      source: 'Funnel.',
      how: 'New customers = leads × conversion.',
      typical: 'Depends on channel.',
      when: 'Update with marketing plan.',
      warn: '—'
    },
    'Day rate – training (CHF)': {
      what: 'Training/consulting fee per day.',
      source: 'Revenue (services).',
      how: 'Revenue = rate × days.',
      typical: 'CHF 1.5k–5k+.',
      when: 'Update your offer.',
      warn: '—'
    },
    'Training days / month': {
      what: 'Billable training/consulting days.',
      source: 'Revenue (services).',
      how: 'Revenue = rate × days.',
      typical: '0–10 early stage.',
      when: 'Update capacity.',
      warn: '—'
    },
    'Other digital products – start (CHF)': {
      what: 'Starting revenue for experiments.',
      source: 'Revenue (other).',
      how: 'Growth compounding per month.',
      typical: 'Small pilot values.',
      when: 'Set when testing ideas.',
      warn: '—'
    },
    'Other digital monthly growth (0–1)': {
      what: 'Monthly growth rate for other digital.',
      source: 'Revenue (other).',
      how: 'Value_t = start×(1+g)^t.',
      typical: '0–10%/mo.',
      when: 'Tune to learning pace.',
      warn: '>20%/mo may be optimistic.'
    },
    'Culture & Learning per FTE / month (CHF)': {
      what: 'Monthly budget per employee for learning.',
      source: 'Opex.',
      how: '= Headcount × budget/FTE.',
      typical: 'CHF 50–200.',
      when: 'Change with policy.',
      warn: '—'
    },
    'Tech stack per FTE / month (CHF)': {
      what: 'SaaS tools/cloud per employee.',
      source: 'Opex.',
      how: '= Headcount × tools/FTE.',
      typical: 'CHF 50–300.',
      when: 'Change with stack.',
      warn: '—'
    },
    'Labs & Universities monthly (CHF)': {
      what: 'Collaboration budget (Aalto etc.).',
      source: 'Opex.',
      how: 'Flat monthly.',
      typical: 'CHF 500–5k+.',
      when: 'Change with contracts.',
      warn: '—'
    },
    'R&D AIX monthly (CHF)': {
      what: 'AIX R&D budget.',
      source: 'Opex.',
      how: 'Flat monthly.',
      typical: 'CHF 1k–10k+.',
      when: 'Change by roadmap.',
      warn: '—'
    },
    'Product dev (CaaS) monthly (CHF)': {
      what: 'CaaS product build budget.',
      source: 'Opex.',
      how: 'Flat monthly.',
      typical: 'CHF 2k–20k+.',
      when: 'Change by roadmap.',
      warn: '—'
    },
    'Media / PR monthly (CHF)': {
      what: 'Paid marketing/PR.',
      source: 'Opex.',
      how: 'Flat monthly.',
      typical: 'CHF 500–20k+.',
      when: 'Change by plan.',
      warn: '—'
    },
    'Purchased Services – Micro monthly (CHF)': {
      what: 'Coaching/individual services.',
      source: 'Opex.',
      how: 'Flat monthly.',
      typical: 'Varies.',
      when: 'Change by contracts.',
      warn: '—'
    },
    'Purchased Services – Meso monthly (CHF)': {
      what: 'Team-level external services.',
      source: 'Opex.',
      how: 'Flat monthly.',
      typical: 'Varies.',
      when: 'Change by contracts.',
      warn: '—'
    },
    'Purchased Services – Macro monthly (CHF)': {
      what: 'Org-level consulting.',
      source: 'Opex.',
      how: 'Flat monthly.',
      typical: 'Varies.',
      when: 'Change by contracts.',
      warn: '—'
    },
    'Purchased Services – Mundo monthly (CHF)': {
      what: 'Specialist expertise (e.g., MLOps).',
      source: 'Opex.',
      how: 'Flat monthly.',
      typical: 'Varies.',
      when: 'Change by contracts.',
      warn: '—'
    },
    // Programmatic fallback:
    __generic: (lab) => ({
      what: `Input for ${lab}.`,
      source: 'Feeds MODEL/IS as appropriate.',
      how: 'Referenced via INDEX/MATCH by label.',
      typical: 'Set to your reality.',
      when: 'Edit as your plan updates.',
      warn: 'Blank may propagate zeros.'
    })
  };
}
function headerTitleNotes_() {
  const addHead = (sh, map) => {
    if (!sh) return;
    const lc = sh.getLastColumn(); const hdr = sh.getRange(1, 1, 1, lc);
    const hvals = hdr.getValues()[0]; const notes = hdr.getNotes();
    for (let i = 0; i < hvals.length; i++) {
      const k = String(hvals[i] || '').trim();
      const t = map[k]; if (!t) continue;
      const tag = `— Colibri: ${t}`; const cur = notes[0][i] || '';
      if (cur.indexOf(tag) === -1) notes[0][i] = cur ? `${cur}\n\n${tag}` : tag;
    }
    hdr.setNotes(notes);
  };
  // MODEL notes are handled by enrichModel_ to avoid duplication/conflicts.
  addHead(getOrInsert_('SUMMARY', 'SUMMARY'), {
    'Metric': '🏷️ KPI name',
    'Value': '🔢 Value (CHF)',
    'Description': 'ℹ️ Explanation'
  });
  addHead(getOrInsert_('IS', 'INCOME_STATEMENT'), { 'Metric': '🏷️ P&L line item' });
  addHead(getOrInsert_('CF', 'CASH_FLOW'), { 'Metric': '🏷️ Cash flow line item' });
  addHead(getOrInsert_('BS', 'BALANCE_SHEET'), { 'Metric': '🏷️ Balance sheet line item' });
}

/* -------------------- 5) MODEL blue color scales (safe) -------------------- */
function applyModelColorScales() {
  const sh = getSheetByCandidates_(SHEET_CANDIDATES.MODEL);
  if (!sh) { SpreadsheetApp.getActive().toast('MODEL not found', 'Colibri', 5); return; }
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const targets = ['MRR_Display', 'ARR_Display', 'Revenue_Display', 'Costs_Display', 'Net_Burn', 'Active_Customers', 'Headcount_Total']
    .map(h => headers.indexOf(h) + 1).filter(i => i > 0);

  // Keep all existing rules, just ADD our per-column gradients (no deletes)
  const rules = sh.getConditionalFormatRules();
  targets.forEach(col => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([sh.getRange(2, col, Math.max(lr - 1, 1), 1)])
      .setGradientMinpoint('#E3F2FD')
      .setGradientMaxpoint('#0D47A1')
      .build();
    rules.push(rule);
  });
  sh.setConditionalFormatRules(rules);
  SpreadsheetApp.getActive().toast('Blue gradients applied (MODEL) ✅', 'Colibri', 5);
}

/* -------------------- 6) Fix assumption types & formats -------------------- */
function fixAssumptionsTypesAndFormats() {
  const ass = getOrInsert_('ASSUMPTIONS', 'ASSUMPTIONS');
  const lr = ass.getLastRow(); if (lr < 2) { SpreadsheetApp.getActive().toast('ASSUMPTIONS empty', 'Colibri', 5); return; }
  const labels = ass.getRange(2, 2, lr - 1, 1).getValues().map(r => String(r[0] || '').trim());
  const valsR = ass.getRange(2, 1, lr - 1, 1);
  const vals = valsR.getValues();

  const isDateLabel = (lab) => /start date|end date/i.test(lab);
  const isPercentLabel = (lab) => /%(?!.*month)/i.test(lab) || /(yield|tax rate)/i.test(lab);
  const isRatioLabel = (lab) => /(0\s*[–-]\s*1|0-1|churn|conversion|growth)/i.test(lab);
  const isCurrencyLabel = (lab) => /(CHF|\$|price|capex|budget|monthly|yearly|salary|cost|fees)/i.test(lab) && !isPercentLabel(lab) && !isRatioLabel(lab);
  const isCountLabel = (lab) => /(headcount|customers|days|months)/i.test(lab) && !isCurrencyLabel(lab) && !isPercentLabel(lab);

  let fixed = 0, warned = 0;
  for (let i = 0; i < labels.length; i++) {
    const lab = labels[i];
    const row = i + 2;
    const cell = ass.getRange(row, 1);
    const v = vals[i][0];

    // Apply number formats first (non-destructive)
    if (isDateLabel(lab)) cell.setNumberFormat('yyyy-mm-dd');
    else if (isPercentLabel(lab)) cell.setNumberFormat('0.00%');
    else if (isRatioLabel(lab)) cell.setNumberFormat('0.000');
    else if (isCurrencyLabel(lab)) cell.setNumberFormat('#,##0.00');
    else if (isCountLabel(lab)) cell.setNumberFormat('0');

    // Convert mis-typed Date objects back to numeric where a number is expected
    const expectsDate = isDateLabel(lab);
    const expectsNumber = !expectsDate;
    if (expectsNumber && v instanceof Date) {
      const disp = cell.getDisplayValue();
      const m = disp.replace(/\s/g, '').match(/[-+]?\d{1,3}(?:[\'\s]?\d{3})*(?:[\.,]\d+)?|[-+]?\d+(?:[\.,]\d+)?/);
      if (m) {
        const num = parseFloat(m[0].replace(/'/g, '').replace(',', '.'));
        if (!isNaN(num)) { vals[i][0] = num; fixed++; continue; }
      }
      // Fallback: convert date to serial days (Excel/Sheets base)
      const serial = Math.round((v - new Date('1899-12-30')) / 86400000);
      vals[i][0] = serial; fixed++;
      appendNoteCell_(cell, 'Converted from Date; please verify expected numeric.');
    }
    if (expectsDate && typeof v === 'number') {
      if (v > 59 && v < 60000) {
        const d = new Date(1899, 11, 30); d.setDate(d.getDate() + v);
        vals[i][0] = d; fixed++;
      } else {
        warned++;
        appendNoteCell_(cell, 'Numeric found where Date expected. Please enter yyyy-mm-dd.');
      }
    }
  }
  valsR.setValues(vals);
  SpreadsheetApp.getActive().toast(`Assumptions formats fixed: ${fixed} changed, ${warned} warned ✅`, 'Colibri', 5);
}

/* -------------------- 7) Build/Refresh MODEL & SUMMARY -------------------- */
function buildOrRefreshModelAndSummary() {
  const ass = getOrInsert_('ASSUMPTIONS', 'ASSUMPTIONS');
  const model = getOrInsert_('MODEL', 'MODEL');

  const A = (t) => `INDEX(ASSUMPTIONS!A:A, MATCH("${t}", ASSUMPTIONS!B:B, 0))`;
  const A0 = (t) => `IFERROR(${A(t)},0)`;

  const sr = findLabelRow_(ass, 'Start Date (yyyy-mm-dd)');
  const er = findLabelRow_(ass, 'End Date (yyyy-mm-dd)');
  const start = sr ? ass.getRange(sr, 1).getValue() : null;
  const end = er ? ass.getRange(er, 1).getValue() : null;
  if (!(start instanceof Date) || !(end instanceof Date)) throw new Error('Start/End Date assumptions must be valid dates.');
  const months = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth()) + 1;
  const rows = Math.max(12, Math.min(180, months));

  const headers = [
    'Month', 'In_Horizon',
    'Active_Customers', 'New_Customers', 'Churned_Customers',
    'ARPU_CaaS', 'MRR_Display', 'ARR_Display',
    'Headcount_Total', 'Payroll_CHF',
    'COGS_CHF', 'Other_Opex_CHF',
    'Revenue_Display', 'Costs_Display', 'Net_Burn', 'Cum_Cash'
  ];
  model.clear();
  model.getRange(1, 1, 1, headers.length).setValues([headers]);

  for (let r = 0; r < rows; r++) {
    const row = 2 + r;
    model.getRange(row, 1).setFormula(`=EDATE(${A('Start Date (yyyy-mm-dd)')}, ${r})`); // Month
    model.getRange(row, 2).setFormula(`=IF(${model.getRange(row, 1).getA1Notation()}<=${A('End Date (yyyy-mm-dd)')},1,)`); // In_Horizon
  }

  for (let r = 0; r < rows; r++) {
    const row = 2 + r;
    const colAC = 3, colNew = 4, colChurn = 5;
    if (r === 0) {
      model.getRange(row, colAC).setFormula(`=${A0('Initial CaaS customers')}`);
    } else {
      const prevAC = model.getRange(row - 1, colAC).getA1Notation();
      const churnRate = A0('Churn monthly (0–1)');
      const newCust = model.getRange(row, colNew).getA1Notation();
      model.getRange(row, colAC).setFormula(`=${prevAC}*(1-${churnRate})+IFERROR(${newCust},0)`);
    }
    model.getRange(row, colNew).setFormula(`=IF(${model.getRange(row, 2).getA1Notation()}=1, ${A0('Monthly leads')}*${A0('Awareness→Conversion (0–1)')}, )`);
    const prevACRef = r === 0 ? A0('Initial CaaS customers') : model.getRange(row - 1, colAC).getA1Notation();
    model.getRange(row, colChurn).setFormula(`=IFERROR(${prevACRef}*${A0('Churn monthly (0–1)')},)`);
  }

  for (let r = 0; r < rows; r++) {
    const row = 2 + r; const ac = model.getRange(row, 3).getA1Notation();
    model.getRange(row, 6).setFormula(`=${A0('ARPU – CaaS monthly (CHF)')}`);
    model.getRange(row, 7).setFormula(`=IFERROR(${ac}*${model.getRange(row, 6).getA1Notation()},)`);
    model.getRange(row, 8).setFormula(`=${model.getRange(row, 7).getA1Notation()}*12`);
  }

  // Headcount: ramp plus dated hires if provided
  for (let r = 0; r < rows; r++) {
    const row = 2 + r; const m = model.getRange(row, 1).getA1Notation();
    const ramp = `ROUND( ${A0('Starting Headcount')} + ( ${A0('Target Headcount (by year 3)')} - ${A0('Starting Headcount')} ) * MIN(DATEDIF(${A('Start Date (yyyy-mm-dd)')}, ${m}, "M"),36) / 36 )`;
    const hires = Array.from({ length: 8 }, (_, i) => i + 1)
      .map(i => `IF(AND(NOT(ISBLANK(${A(`Hire ${i} – Start (yyyy-mm-dd)`)})), ${m}>=${A(`Hire ${i} – Start (yyyy-mm-dd)`)}), 1, 0)`) // proper IF(AND(...),1,0)
      .join('+');
    const hcExpr = `IFERROR(${ramp} + (${hires}), ${ramp})`;
    model.getRange(row, 9).setFormula(`=${hcExpr}`);
  }

  // Payroll uses level map per hire if provided, fallback to L3 average
  for (let r = 0; r < rows; r++) {
    const row = 2 + r; const hc = model.getRange(row, 9).getA1Notation(); const mo = model.getRange(row, 1).getA1Notation();
    const lvlToSal = (lvl) => `IF(${lvl}="L1", ${A0('Salary L1 (CHF/yr)')}, IF(${lvl}="L2", ${A0('Salary L2 (CHF/yr)')}, IF(${lvl}="L3", ${A0('Salary L3 (CHF/yr)')}, IF(${lvl}="L4", ${A0('Salary L4 (CHF/yr)')}, ${A0('Salary L3 (CHF/yr)')}))))`;
    const perHire = Array.from({ length: 8 }, (_, i) => {
      const idx = i + 1; const start = A(`Hire ${idx} – Start (yyyy-mm-dd)`); const lvl = A(`Hire ${idx} – Level (L1–L4)`);
      return `IF(AND(NOT(ISBLANK(${start})), ${mo}>=${start}), ${lvlToSal(lvl)}/12, 0)`;
    }).join('+');
    const baseAvg = `(${hc} * (${A0('Salary L3 (CHF/yr)')}/12))`;
    const salary = `IFERROR(${perHire} + MAX(0, ${hc} - (${Array.from({ length: 8 }, (_, i) => `IF(AND(NOT(ISBLANK(${A(`Hire ${i + 1} – Start (yyyy-mm-dd)`)})), ${mo}>=${A(`Hire ${i + 1} – Start (yyyy-mm-dd)`)}), 1, 0)`).join('+')})) * (${A0('Salary L3 (CHF/yr)')}/12), ${baseAvg})`;
    model.getRange(row, 10).setFormula(`=IFERROR( (${salary}) * (1+${A0('Payroll Overhead %')}/100), )`);
  }

  for (let r = 0; r < rows; r++) {
    const row = 2 + r; const ac = model.getRange(row, 3).getA1Notation();
    const mrr = model.getRange(row, 7).getA1Notation();
    model.getRange(row, 11).setFormula(`=IFERROR(${mrr}*${A0('Cost of Revenue % (COGS)')}/100 + ${ac}*${A0('AI Agents cost per active customer per month (CHF)')}, )`);
  }

  const opexFormulaParts = [
    `IFERROR(${A0('Culture & Learning per FTE / month (CHF)')}*${'@HC'},0)`,
    `IFERROR(${A0('Tech stack per FTE / month (CHF)')}*${'@HC'},0)`,
    `IFERROR(${A0('Media / PR monthly (CHF)')},0)`,
    `IFERROR(${A0('R&D AIX monthly (CHF)')},0)`,
    `IFERROR(${A0('Product dev (CaaS) monthly (CHF)')},0)`,
    `IFERROR(${A0('Labs & Universities monthly (CHF)')},0)`,
    `IFERROR(${A0('Purchased Services – Micro monthly (CHF)')},0)`,
    `IFERROR(${A0('Purchased Services – Meso monthly (CHF)')},0)`,
    `IFERROR(${A0('Purchased Services – Macro monthly (CHF)')},0)`,
    `IFERROR(${A0('Purchased Services – Mundo monthly (CHF)')},0)`
  ];
  for (let r = 0; r < rows; r++) {
    const row = 2 + r; const hcAddr = model.getRange(row, 9).getA1Notation();
    const expr = opexFormulaParts.map(p => p.replace('@HC', hcAddr)).join('+');
    model.getRange(row, 12).setFormula(`=${expr}`);
  }

  for (let r = 0; r < rows; r++) {
    const row = 2 + r; const mrr = model.getRange(row, 7).getA1Notation();
    const training = `IFERROR(${A0('Day rate – training (CHF)')} * ${A0('Training days / month')}, 0)`;
    const otherDigital = `IFERROR(${A0('Other digital products – start (CHF)')} * (1+${A0('Other digital monthly growth (0–1)')})^DATEDIF(${A('Start Date (yyyy-mm-dd)')}, ${model.getRange(row, 1).getA1Notation()}, "M"), 0)`;
    model.getRange(row, 13).setFormula(`=${mrr} + ${training} + ${otherDigital}`);
  }

  for (let r = 0; r < rows; r++) {
    const row = 2 + r;
    model.getRange(row, 14).setFormula(`=IFERROR(SUM(${model.getRange(row, 10).getA1Notation()},${model.getRange(row, 11).getA1Notation()},${model.getRange(row, 12).getA1Notation()}),0)`);
  }

  for (let r = 0; r < rows; r++) {
    const row = 2 + r; const burnCol = 15, cashCol = 16;
    model.getRange(row, burnCol).setFormula(`=IFERROR(${model.getRange(row, 14).getA1Notation()}-${model.getRange(row, 13).getA1Notation()},0)`);
    if (r === 0) {
      model.getRange(row, cashCol).setFormula(`=IFERROR(${A0('Starting Cash (CHF)')}-${model.getRange(row, burnCol).getA1Notation()}, ${A0('Starting Cash (CHF)')})`);
    } else {
      model.getRange(row, cashCol).setFormula(`=IFERROR(${model.getRange(row - 1, cashCol).getA1Notation()}-${model.getRange(row, burnCol).getA1Notation()}, ${model.getRange(row - 1, cashCol).getA1Notation()})`);
    }
  }

  model.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  model.setFrozenRows(1);
  applyModelColorScales();

  buildOrRefreshSummary_(model);

  try {
    const charts = model.getCharts();
    const existing = charts.find(c => c.getOptions().get('title') === 'Revenue vs Costs');
    if (existing) model.removeChart(existing);
    const lr = model.getLastRow();
    const chart = model.newChart()
      .asLineChart()
      .setOption('title', 'Revenue vs Costs')
      .addRange(model.getRange(1, 1, lr, 1))
      .addRange(model.getRange(1, 13, lr, 1))
      .addRange(model.getRange(1, 14, lr, 1))
      .setPosition(1, headers.length + 2, 0, 0)
      .build();
    model.insertChart(chart);
  } catch (e) { }

  SpreadsheetApp.getActive().toast('MODEL & SUMMARY rebuilt ✅', 'Colibri', 5);
}

function buildOrRefreshSummary_(model) {
  const sum = getOrInsert_('SUMMARY', 'SUMMARY'); sum.clear();
  const hdr = ['Metric', 'Value', 'Description'];
  sum.getRange(1, 1, 1, hdr.length).setValues([hdr]); sum.setFrozenRows(1);
  const lc = model.getLastColumn();
  const headers = model.getRange(1, 1, 1, lc).getValues()[0];
  const cMonth = headers.indexOf('Month') + 1, cIn = headers.indexOf('In_Horizon') + 1, cBurn = headers.indexOf('Net_Burn') + 1, cCash = headers.indexOf('Cum_Cash') + 1;
  if (!cMonth || !cIn || !cBurn || !cCash) return;

  sum.getRange(2, 1).setValue('Latest Month');
  sum.getRange(2, 2).setFormula(`=INDEX(${model.getSheetName()}!${columnLetter_(cMonth)}:${columnLetter_(cMonth)}, MATCH(2, ${model.getSheetName()}!${columnLetter_(cIn)}:${columnLetter_(cIn)}, 1))`);
  sum.getRange(2, 3).setValue('Last period in horizon');

  sum.getRange(3, 1).setValue('MRR');
  sum.getRange(3, 2).setFormula(`=INDEX(${model.getSheetName()}!G:G, MATCH(${sum.getRange(2, 2).getA1Notation()}, ${model.getSheetName()}!${columnLetter_(cMonth)}:${columnLetter_(cMonth)},0))`);
  sum.getRange(3, 3).setValue('Monthly Recurring Revenue');

  sum.getRange(4, 1).setValue('Net Burn');
  sum.getRange(4, 2).setFormula(`=INDEX(${model.getSheetName()}!${columnLetter_(cBurn)}:${columnLetter_(cBurn)}, MATCH(${sum.getRange(2, 2).getA1Notation()}, ${model.getSheetName()}!${columnLetter_(cMonth)}:${columnLetter_(cMonth)},0))`);
  sum.getRange(4, 3).setValue('Costs - Revenue (positive = burn)');

  sum.getRange(5, 1).setValue('Runway (months)');
  // Use average burn over last 3 months to avoid noisy spikes
  const idxCur = `MATCH(${sum.getRange(2, 2).getA1Notation()}, ${model.getSheetName()}!${columnLetter_(cMonth)}:${columnLetter_(cMonth)},0)`;
  const avgBurn3 = `AVERAGE(OFFSET(${model.getSheetName()}!${columnLetter_(cBurn)}1, ${idxCur}-1-2, 0, 3, 1))`;
  sum.getRange(5, 2).setFormula(`=IFERROR( INDEX(${model.getSheetName()}!${columnLetter_(cCash)}:${columnLetter_(cCash)}, ${idxCur}) / MAX(0.0001, ${avgBurn3}), 0)`);
  sum.getRange(5, 3).setValue('Months of cash at average burn (last 3 months). Aim > 10.');

  sum.autoResizeColumns(1, 3);
}

/* -------------------- helpers -------------------- */
function diagnostics() {
  const ss = SpreadsheetApp.getActive();
  const found = key => !!getSheetByCandidates_(SHEET_CANDIDATES[key]);
  const parts = [];
  parts.push(`🗂️ File: ${ss.getName()}`);
  parts.push('');
  // Presence checks
  parts.push('✅ Tabs present:');
  parts.push(`• MODEL: ${found('MODEL') ? '✅' : '❌'}  • ASSUMPTIONS: ${found('ASSUMPTIONS') ? '✅' : '❌'}  • MAP: ${found('MAP') ? '✅' : '❌'}`);
  parts.push(`• IS: ${found('IS') ? '✅' : '❌'}  • CF: ${found('CF') ? '✅' : '❌'}  • BS: ${found('BS') ? '✅' : '❌'}  • SUMMARY: ${found('SUMMARY') ? '✅' : '❌'}`);
  parts.push('');
  const next = [];
  if (!found('MODEL') || !found('SUMMARY')) next.push('Run ③ Build/Refresh MODEL & SUMMARY.');
  if (!found('IS') || !found('CF') || !found('BS')) next.push('Run ⑤ Build / Refresh Financial Statements.');
  if (!found('MAP')) next.push('Optional: Run ④ Build/Refresh Mapping to set categories/buckets.');

  // Formula errors report
  let errCount = 0;
  try {
    const rep = getSheetByCandidates_(['FORMULA_ERRORS', 'Formula Errors', 'Errors']);
    if (rep) { const lr = rep.getLastRow(); errCount = Math.max(0, lr - 1); }
  } catch (e) { }
  if (errCount > 0) {
    parts.push(`🛑 Formula issues: ${errCount} cells flagged in FORMULA_ERRORS`);
    next.push('Open FORMULA_ERRORS and fix the listed cells, or re-run ⑫ Scan Formula Errors & Report.');
  } else {
    parts.push('🧪 Formula check: No known errors. If in doubt, run ⑫ Scan Formula Errors & Report.');
  }

  // Common gaps & startup flags
  try {
    const ass = getOrInsert_('ASSUMPTIONS', 'ASSUMPTIONS');
    const A_ = (lab) => { try { const r = findLabelRow_(ass, lab); return r ? ass.getRange(r, 1).getValue() : null; } catch (e) { return null; } };
    const cg = [];
    // Timeline horizon
    let monthsH = null; try {
      const s = A_('Start Date (yyyy-mm-dd)'); const e = A_('End Date (yyyy-mm-dd)');
      if (s instanceof Date && e instanceof Date) {
        monthsH = (e.getFullYear() - s.getFullYear()) * 12 + (e.getMonth() - s.getMonth()) + 1;
        if (monthsH < 12) { cg.push('🟠 Horizon < 12 months — extend End Date to at least 12–36 months for planning.'); next.push('In ASSUMPTIONS, set End Date to >= +12 months, then run ③.'); }
        if (monthsH > 120) { cg.push('🟠 Horizon > 10 years — consider shortening to 3–7 years.'); }
      }
    } catch (e) { }

    // Cash and runway basics
    const startCash = Number(A_('Starting Cash (CHF)') || 0);
    if (startCash <= 0) { cg.push('🟠 Starting Cash is 0 — runway will be 0 until you set a positive value.'); next.push('Set “Starting Cash (CHF)” in ASSUMPTIONS and rebuild ③ → ⑤.'); }

    // CapEx / Depreciation consistency
    const capex = Number(A_('CapEx per month (CHF)') || 0);
    const depm = Number(A_('Depreciation months') || 0);
    if (capex > 0 && (!depm || depm <= 0)) { cg.push('🟠 CapEx > 0 but Depreciation months is not set — PP&E won’t decline.'); next.push('Set “Depreciation months” (e.g., 36) in ASSUMPTIONS.'); }
    if (depm > 0 && capex === 0) { cg.push('🟠 Depreciation months set but CapEx = 0 — PP&E will trend to 0.'); }

    // Taxes / yield sanity
    const tax = Number(A_('Tax rate %') || 0);
    if (tax > 50) { cg.push('🟠 Tax rate % > 50 — unusually high. Typical CH corporate taxes 12–35%.'); }
    if (tax < 0) { cg.push('🟠 Tax rate % < 0 — invalid.'); }
    const yieldP = Number(A_('Cash yield %') || 0);
    if (yieldP > 10) { cg.push('🟠 Cash yield % > 10 — unusually high for cash balances.'); }

    // Growth drivers sanity
    const arpu = Number(A_('ARPU – CaaS monthly (CHF)') || 0);
    if (arpu <= 0) { cg.push('🟠 ARPU (CaaS) is 0 — MRR will be 0.'); next.push('Set “ARPU – CaaS monthly (CHF)” to a realistic value (e.g., 100–600).'); }
    const leads = Number(A_('Monthly leads') || 0);
    const conv = Number(A_('Awareness→Conversion (0–1)') || 0);
    const churn = Number(A_('Churn monthly (0–1)') || 0);
    if (leads <= 0) { cg.push('🟠 Monthly leads is 0 — New_Customers will be 0.'); }
    if (conv <= 0) { cg.push('🟠 Conversion is 0 — growth impossible.'); }
    if (conv > 0.5) { cg.push('🟠 Conversion > 50% — likely unrealistic at scale.'); }
    if (churn < 0 || churn > 1) { cg.push('🛑 Churn outside 0–1 — must be a ratio (e.g., 0.03 for 3%).'); next.push('Fix “Churn monthly (0–1)” to a 0–1 value.'); }
    else if (churn > 0.2) { cg.push('🟠 Churn > 20%/mo — extremely high; revisit retention.'); }

    // Runway threshold (low)
    try {
      const sum = getSheetByCandidates_(SHEET_CANDIDATES.SUMMARY);
      if (sum && sum.getLastRow() >= 5) {
        const rw = Number(sum.getRange(5, 2).getValue() || 0);
        if (rw > 0 && rw < 6) { cg.push(`🟠 Low runway ≈ ${rw.toFixed(1)} months — consider cutting burn, boosting revenue, or fundraising.`); }
      }
    } catch (e) { }

    // GM% sanity at latest column
    try {
      const isSh = getOrInsert_('IS', 'INCOME_STATEMENT');
      const lc = isSh.getLastColumn(); if (lc >= 2) {
        const gm = Number(isSh.getRange(8, lc).getValue() || 0);
        if (gm > 0.95) cg.push('🟠 Gross Margin > 95% — check COGS assumptions.');
        if (gm > 0 && gm < 0.2) cg.push('🟠 Gross Margin < 20% — check pricing/COGS.');
      }
    } catch (e) { }

    // MRR presence by latest column
    try {
      const model = getSheetByCandidates_(SHEET_CANDIDATES.MODEL);
      if (model) {
        const mHdr = model.getRange(1, 1, 1, model.getLastColumn()).getValues()[0];
        const cMRR = mHdr.indexOf('MRR_Display') + 1;
        const cIn = mHdr.indexOf('In_Horizon') + 1;
        if (cMRR > 0 && cIn > 0) {
          // Find last in-horizon row
          const hv = model.getRange(2, cIn, model.getLastRow() - 1, 1).getValues();
          let last = 1; for (let i = 0; i < hv.length; i++) if (hv[i][0] === 1 || hv[i][0] === '1') last = i + 1;
          const mrr = Number(model.getRange(1 + last, cMRR).getValue() || 0);
          const initCust = Number(A_('Initial CaaS customers') || 0);
          if (mrr <= 0 && (initCust > 0 || leads > 0)) cg.push('🟠 MRR is 0 despite customers/leads — re-run ③/⑤ or check ARPU/conversion.');
        }
      }
    } catch (e) { }

    // Mapping "Other" dominance
    try {
      const mapSh = getOrInsert_('MAP', 'MAPPING_CATEGORIES');
      const lr = mapSh.getLastRow();
      if (lr >= 2) {
        const rows = mapSh.getRange(2, 1, lr - 1, 4).getValues();
        const catAgg = {};
        rows.forEach(r => {
          const label = String(r[1] || '').trim(); const cat = r[2] || 'Other';
          const rr = findLabelRow_(ass, label); const val = rr ? ass.getRange(rr, 1).getValue() : 0;
          if (typeof val === 'number') catAgg[cat] = (catAgg[cat] || 0) + val;
        });
        const total = Object.values(catAgg).reduce((a, b) => a + b, 0);
        const other = catAgg['Other'] || 0;
        if (total > 0 && other / total > 0.5 && other > 500) {
          cg.push('🟠 >50% of mapped spend is “Other” — refine categories for better insight.');
          next.push('Open MAPPING_CATEGORIES and reassign big labels from “Other” to a specific category.');
        }
      }
    } catch (e) { }

    if (cg.length) {
      parts.push('');
      parts.push('🔎 Common gaps & flags:');
      cg.forEach(m => parts.push(`• ${m}`));
    }
  } catch (e) { /* ignore, best-effort diagnostics */ }

  // Runway sanity (SUMMARY B5)
  try {
    const sum = getSheetByCandidates_(SHEET_CANDIDATES.SUMMARY);
    if (sum && sum.getLastRow() >= 5) {
      const v = Number(sum.getRange(5, 2).getValue() || 0);
      if (v > 120) {
        parts.push(`⚠️ Runway looks very high (B5 ≈ ${Math.round(v)} months).`);
        parts.push('   This likely means burn ~ 0. Check MODEL Net_Burn and ASSUMPTIONS costs.');
        next.push('Validate key costs in ASSUMPTIONS (COGS %, payroll, opex).');
      } else if (v <= 0) {
        parts.push('⚠️ Runway is not meaningful (≤ 0).');
        next.push('Check MODEL Net_Burn sign and Starting Cash in ASSUMPTIONS.');
      } else {
        parts.push('🟢 Runway looks reasonable.');
      }
    }
  } catch (e) { }

  // Balance sheet zeros / PP&E
  try {
    const bs = getSheetByCandidates_(SHEET_CANDIDATES.BS);
    if (bs && bs.getLastRow() >= 6 && bs.getLastColumn() >= 2) {
      const c = Math.min(3, bs.getLastColumn());
      const vals = bs.getRange(2, c, 5, 1).getValues().map(r => Number(r[0] || 0));
      const mostlyZero = vals.filter(x => Math.abs(x) < 0.0001).length >= 4;
      if (mostlyZero) {
        parts.push('⚠️ Balance Sheet shows zeros for Cash/PP&E.');
        parts.push('   Likely causes: no CapEx per month, zero Depreciation months, or MODEL Cum_Cash not linked.');
        next.push('Set “CapEx per month (CHF)” and “Depreciation months” in ASSUMPTIONS, then run ⑤.');
      } else {
        parts.push('🟢 Balance sheet core lines populated.');
      }
    }
  } catch (e) { }

  // Mapping guidance
  parts.push('');
  parts.push('🧭 Tips: If you changed key assumptions, re-run ③ then ⑤. For beginner guidance, run ⑪ Onboarding and ⑩ Enrich Notes & Guides.');
  if (next.length) { parts.push(''); parts.push('👉 Next actions:'); next.forEach(s => parts.push(`• ${s}`)); }

  SpreadsheetApp.getUi().alert(parts.join('\n'));
}

/* -------------------- Formula errors scan & report -------------------- */
function scanFormulaErrorsReport() {
  const ss = SpreadsheetApp.getActive();
  const report = getSheetByCandidates_(['FORMULA_ERRORS', 'Formula Errors', 'Errors']) || ss.insertSheet('FORMULA_ERRORS');
  report.clear();
  const header = ['Sheet', 'Cell', 'Formula', 'Error'];
  const out = [header];
  const errorRegex = /^#(N\/A|REF!|VALUE!|NAME\?|DIV\/0!|NUM!|NULL!)/i;

  ss.getSheets().forEach(sh => {
    try {
      const lr = Math.max(1, sh.getLastRow()), lc = Math.max(1, sh.getLastColumn());
      if (lr === 1 && lc === 1 && !sh.getRange(1, 1).getValue() && !sh.getRange(1, 1).getFormula()) return;
      const rng = sh.getRange(1, 1, lr, lc);
      const formulas = rng.getFormulas();
      const values = rng.getValues();
      for (let r = 0; r < lr; r++) {
        for (let c = 0; c < lc; c++) {
          const f = formulas[r][c];
          if (!f) continue;
          const v = values[r][c];
          const s = typeof v === 'string' ? v : '';
          if (errorRegex.test(s)) {
            const a1 = sh.getRange(r + 1, c + 1).getA1Notation();
            out.push([sh.getName(), a1, f, s]);
          }
        }
      }
    } catch (e) { /* ignore sheet errors to not break report */ }
  });

  report.getRange(1, 1, out.length, header.length).setValues(out);
  report.setFrozenRows(1);
  if (out.length > 1) {
    report.getRange(2, 1, out.length - 1, header.length).setBackground('#FDECEA');
  } else {
    report.getRange(2, 1).setValue('No formula errors found ✅');
  }
  SpreadsheetApp.getActive().toast('Formula errors scan complete ✅', 'Colibri', 5);
}

/* -------------------- Onboarding Walkthrough -------------------- */
function createOnboardingWalkthrough() {
  const ss = SpreadsheetApp.getActive();
  // Create or clear ONBOARDING sheet
  let onboard = getSheetByCandidates_(['ONBOARDING', 'Onboarding', 'Walkthrough']);
  if (!onboard) onboard = ss.insertSheet('ONBOARDING'); else onboard.clear();
  onboard.setTabColor('#BDE5F8');
  const rows = [];
  rows.push(['Welcome to Colibri 🐦', 'A quick guided path to set up your model in ~10 minutes.']);
  rows.push(['1) Fix inputs', 'Use Colibri menu > ① Normalize Dates & Sections, then ② Fix Assumptions: Types & Formats.']);
  rows.push(['2) Build the model', 'Run ③ Build/Refresh MODEL & SUMMARY to compute core metrics and the dashboard.']);
  rows.push(['3) Financial statements', 'Run ⑤ Build / Refresh Financial Statements to build IS / CF / BS.']);
  rows.push(['4) Make it readable', 'Run ⑦ Apply MODEL Color Scales and ⑧ EU Formatting & Clean Comments.']);
  rows.push(['5) Guides and notes', 'Run ⑩ Enrich Notes & Guides for helpful tips across tabs.']);
  rows.push(['6) Audit your assumptions', 'Run ⑨ Audit ASSUMPTIONS for Outliers (CH) and review the ASSUMPTIONS_AUDIT tab.']);
  rows.push(['7) Optional mapping', 'Run ④ Build/Refresh Mapping if you use category-to-bucket mapping.']);
  rows.push(['Pro tip', 'Edit values only in ASSUMPTIONS (Column A). Labels are in Column B. Avoid typing in MODEL/IS/CF/BS except where noted.']);
  rows.push(['Hiring', 'Fill Hire 1–8 with Role, Level (L1–L4), and Start Date. The model will add headcount and payroll automatically.']);
  rows.push(['Runway', 'SUMMARY uses average burn over last 3 months for stability. Aim for > 10 months.']);
  rows.push(['Re-run safely', 'All actions are idempotent: you can re-run them any time; they won’t duplicate headers or notes.']);
  rows.push(['Where to start?', 'Open ASSUMPTIONS and adjust top “Priority” rows first (highlighted). Then re-run steps ② and ③.']);

  const n = rows.length;
  onboard.getRange(1, 1, n, 2).setValues(rows);
  onboard.setFrozenRows(1);
  // Formatting
  onboard.getRange(1, 1, 1, 2).setFontWeight('bold').setFontSize(18).setBackground('#DFF2FD');
  onboard.getRange(2, 1, n - 1, 1).setFontWeight('bold');
  onboard.getRange(1, 1, n, 2).setWrap(true).setVerticalAlignment('top');
  onboard.autoResizeColumns(1, 2);
  // Add notes with quick command palette reminder
  appendNotesRect_(onboard.getRange(1, 1, n, 2), 'Find the Colibri menu under Extensions > Colibri. You can re-run steps safely.');
  SpreadsheetApp.getActive().toast('ONBOARDING created ✅', 'Colibri', 5);
}

/* -------------------- 10) Enrich Notes & Guides -------------------- */
function enrichNotesAndGuides() {
  applyGlobalUX_();
  enrichReadme_();
  enrichAssumptions_();
  enrichMapping_();
  enrichModel_();
  enrichIS_();
  enrichCF_();
  enrichBS_();
  enrichSummary_();
  enrichFormulaLib_();
  SpreadsheetApp.getActive().toast('Notes & Guides enriched ✅', 'Colibri', 6);
}

function applyGlobalUX_() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sh => {
    try { sh.getDataRange().setFontSize(16); } catch (e) { }
    try {
      const lc = Math.max(1, sh.getLastColumn());
      const color = sh.getTabColor() || '#E0E0E0';
      sh.getRange(1, 1, 1, lc).setBackground(color);
    } catch (e) { }
  });
}

function enrichReadme_() {
  const sh = getOrInsert_('README', 'READ ME');
  // Ensure header
  if (sh.getLastRow() < 1) sh.getRange('A1').setValue('Step');
  const lr = Math.max(sh.getLastRow(), 6);
  const steps = [
    ['Open Colibri menu', 'Use the Colibri menu in this Sheet to run tasks (normalize, build model, format, audit).'],
    ['Onboarding walkthrough', 'Run ⑪ Create Onboarding Walkthrough to get a step-by-step guide in the ONBOARDING sheet.'],
    ['Build the model', 'Run ③ Build/Refresh MODEL & SUMMARY, then ⑤ Build / Refresh Financial Statements. Re-run after input changes.'],
    ['Scan formulas for errors', 'Run ⑫ Scan Formula Errors & Report to generate a FORMULA_ERRORS sheet. Fix items shown, then rebuild ③/⑤.'],
    ['Diagnostics', 'Use the Diagnostics item in the Colibri menu to see a friendly health check, runway sanity, and next actions.'],
    ['Update via GitHub (recommended)', '1) Install Node + clasp\n2) Login: clasp login\n3) Check .clasp.json has your scriptId\n4) From repo folder google-sheets/calculator-apps-script run: clasp push -f\n5) Reload Sheet and run from Colibri menu.'],
    ['Manual update via Google Apps Script', '1) Extensions → Apps Script\n2) Replace Code.js with repository version\n3) Update appsscript.json scopes to include spreadsheets + drive\n4) Save → Run onOpen or use Colibri menu.'],
    ['Troubleshooting', 'If push fails: ensure advanced services set under dependencies.enabledAdvancedServices (Drive v2, Sheets v4). If menu missing: reload sheet.'],
    ['Formatting & locale', 'Run EU Formatting to normalize numerics and set locale. Numbers use space thousands and %/dates formatting as Swiss-friendly.'],
    ['Outlier audit', 'Use the Audit action to highlight unrealistic inputs and get a summary on ASSUMPTIONS_AUDIT.']
  ];
  sh.getRange(1, 1, 1, 2).setValues([['Step', 'Details']]);
  sh.getRange(2, 1, steps.length, 2).setValues(steps);
  // Also add notes in column B for longer guidance
  const notes = [
    'Click the Colibri menu (top) to find all actions.',
    'Creates an ONBOARDING sheet with the key steps and tips — great for new collaborators.',
    'Re-running these builders is safe and idempotent — they won\’t duplicate headers.',
    'FORMULA_ERRORS lists cells with #REF!, #N/A, etc. Tackle the first few, then re-run ③/⑤.',
    'Diagnostics summarizes health with emojis and provides actionable next steps.',
    'GitHub push uses clasp; you can also clone and PR.\nKeep a .claspignore to avoid pushing node_modules.',
    'Manual path is a fallback if clasp is blocked in your org.',
    'Invalid manifest error? Move advancedServices under dependencies.enabledAdvancedServices.',
    'EU Formatting removes apostrophes, normalizes commas/dots, and re-applies header notes.',
    'Audit highlights type issues (yellow), soft outliers (amber), hard outliers (red).'
  ];
  const rng = sh.getRange(2, 2, steps.length, 1);
  const cur = rng.getNotes();
  for (let i = 0; i < steps.length; i++) {
    const tag = `— Colibri: ${notes[i]}`;
    const c = cur[i][0] || '';
    if (c.indexOf(tag) === -1) cur[i][0] = c ? `${c}\n\n${tag}` : tag;
  }
  rng.setNotes(cur);
  sh.setFrozenRows(1);
}

function enrichAssumptions_() {
  const sh = getOrInsert_('ASSUMPTIONS', 'ASSUMPTIONS');
  const lr = sh.getLastRow(); if (lr < 2) return;
  // Column C: typical values, Column D: human comments
  const labels = sh.getRange(2, 2, lr - 1, 1).getValues().map(r => String(r[0] || ''));
  const colC = sh.getRange(2, 3, lr - 1, 1);
  const colD = sh.getRange(2, 4, lr - 1, 1);
  const cVals = colC.getValues();
  const dNotes = colD.getNotes();
  const dVals = colD.getValues();
  const tips = {
    'Salary L1 (CHF/yr)': 'CHF 140’000–220’000',
    'Salary L2 (CHF/yr)': 'CHF 110’000–160’000',
    'Salary L3 (CHF/yr)': 'CHF 90’000–130’000',
    'Salary L4 (CHF/yr)': 'CHF 60’000–95’000',
    'Payroll Overhead %': '~12–22%',
    'ARPU – CaaS monthly (CHF)': 'CHF 100–600',
    'Churn monthly (0–1)': '0.01–0.05',
    'Awareness→Conversion (0–1)': '0.01–0.05',
    'CAC blended (CHF)': 'CHF 1k–8k+',
    'LTV months': '12–60',
    'Culture & Learning per FTE / month (CHF)': 'CHF 50–200',
    'Tech stack per FTE / month (CHF)': 'CHF 50–300',
  };
  for (let i = 0; i < labels.length; i++) {
    const lab = labels[i];
    if (tips[lab] && !cVals[i][0]) cVals[i][0] = tips[lab];
    // Beginner-friendly, emoji-structured note
    const note = `
🔎 What: ${lab}
🧭 How to think: Set a simple, believable number. You can refine later.
💡 Typical: ${tips[lab] || 'See README ranges or your actuals.'}
⚠️ Watch: If this is 0 or extreme, MODEL and SUMMARY will look odd.
✅ Next: Change this → run ③ MODEL → run ⑤ IS/CF/BS → check SUMMARY.`;
    const tag = `— Colibri: ${note}`; const cur = dNotes[i][0] || '';
    if (cur.indexOf(tag) === -1) dNotes[i][0] = cur ? `${cur}\n\n${tag}` : tag;
    // Add ADHD-friendly section label in Column D cell value (not just note)
    const section = (() => {
      if (/start date|end date/i.test(lab)) return '🗓️ Timeline';
      if (/starting cash|cash yield|tax rate/i.test(lab)) return '🏦 Cash & Taxes';
      if (/arpu|leads|conversion|churn|price|customers|ltv|cac/i.test(lab)) return '📦 Revenue & Growth';
      if (/salary|headcount|hire|level|payroll/i.test(lab)) return '👥 Team & Payroll';
      if (/cogs|ai agents|runtime/i.test(lab)) return '🧪 COGS & Delivery';
      if (/capex|depreciation/i.test(lab)) return '🏗️ CapEx & Depreciation';
      if (/media|product dev|r&d|labs|tech stack|culture|purchased services/i.test(lab)) return '🧰 Operating Expenses';
      return '';
    })();
    const curD = String(dVals[i][0] || '').trim();
    const hasEmoji = /^(🗓️|🏦|📦|👥|🧪|🏗️|🧰)\b/.test(curD);
    if (section && (curD === '' || !hasEmoji)) dVals[i][0] = section;
  }
  colC.setValues(cVals);
  colD.setNotes(dNotes);
  colD.setValues(dVals);

  // Highlight priority assumptions for beginners (darker green background whole row)
  const priority = ['ARPU – CaaS monthly (CHF)', 'Monthly leads', 'Awareness→Conversion (0–1)', 'Churn monthly (0–1)', 'Cost of Revenue % (COGS)', 'Starting Cash (CHF)', 'CAC blended (CHF)'];
  const lastRow = sh.getLastRow();
  for (let i = 0; i < labels.length; i++) {
    const lab = labels[i];
    if (priority.indexOf(lab) > -1) { sh.getRange(2 + i, 1, 1, 4).setBackground('#C8E6C9'); }
  }

  // Headcount schedule: allow hire dates for up to 8 roles
  ensureAssumptionValueWithNote_(sh, 'Hire 1 – Role', '', { what: 'Role name for hire 1', source: 'Payroll plan', how: 'Set a role label', typical: 'e.g., Engineer', when: 'When planning hires', warn: '—' });
  ensureAssumptionValueWithNote_(sh, 'Hire 1 – Level (L1–L4)', 'L3', { what: 'Level mapping to salary reference', source: 'Payroll plan', how: 'Use L1..L4 to pick salary', typical: 'L3/L2', when: 'Adjust per candidate', warn: '—' });
  ensureAssumptionValueWithNote_(sh, 'Hire 1 – Start (yyyy-mm-dd)', new Date(), { what: 'Hire date', source: 'Payroll plan', how: 'Used to turn on HC at this month', typical: 'First of month', when: 'When you sign', warn: '—' });
  for (let i = 2; i <= 8; i++) {
    ensureAssumptionValueWithNote_(sh, `Hire ${i} – Role`, '', { what: `Role name for hire ${i}`, source: 'Payroll plan', how: 'Set a role label', typical: '', when: '', warn: '—' });
    ensureAssumptionValueWithNote_(sh, `Hire ${i} – Level (L1–L4)`, '', { what: 'Level mapping to salary reference', source: 'Payroll plan', how: 'Use L1..L4 to pick salary', typical: '', when: '', warn: '—' });
    ensureAssumptionValueWithNote_(sh, `Hire ${i} – Start (yyyy-mm-dd)`, '', { what: 'Hire date', source: 'Payroll plan', how: 'Used to turn on HC at this month', typical: '', when: '', warn: '—' });
  }

  // Data validation for hire fields
  const lvls = ['L1', 'L2', 'L3', 'L4'];
  for (let i = 1; i <= 8; i++) {
    const rLvl = findLabelRow_(sh, `Hire ${i} – Level (L1–L4)`);
    if (rLvl) sh.getRange(rLvl, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(lvls, true).build());
    const rStart = findLabelRow_(sh, `Hire ${i} – Start (yyyy-mm-dd)`);
    if (rStart) { sh.getRange(rStart, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build()); sh.getRange(rStart, 1).setNumberFormat('yyyy-mm-dd'); }
    const rRole = findLabelRow_(sh, `Hire ${i} – Role`);
    if (rRole) { try { sh.getRange(rRole, 1).clearDataValidations(); } catch (e) { } }
  }

  // Clear validation for specific cells A38 and A41 if they exist by position
  try { const a38 = sh.getRange(38, 1); a38.clearDataValidations(); } catch (e) { }
  try { const a41 = sh.getRange(41, 1); a41.clearDataValidations(); } catch (e) { }
}

function enrichMapping_() {
  const sh = getOrInsert_('MAP', 'MAPPING_CATEGORIES'); if (!sh) return;
  // Add long description at top-left cell
  sh.getRange('A1').setNote('🗺️ How to use (read me)\n\n• Each row links an ASSUMPTION label to a Category and a Bucket (COGS vs Opex).\n• Category = analysis grouping (infra, tools, services).\n• Bucket = accounting placement (COGS lowers GM%, Opex affects EBITDA).\n• Start with BIG drivers first: infra/AI runtime/tools/media.\n\nWhy this matters: Clear mapping → clean P&L split → faster decisions.');
  // Header notes for clear purposes
  const map = {
    'Source': '📍 Source tab (ASSUMPTIONS/Growth/Cost). Keep consistent to find labels quickly.',
    'Label': '🏷️ Exact ASSUMPTIONS label. Spelling matters (INDEX/MATCH lookup).',
    'Category': '🧩 Group for analysis. Best practice: group by theme (Infra, AI Runtime, Tools, Media, Services).',
    'Bucket': '⚖️ COGS vs Opex. Rule of thumb: Direct to serve revenue → COGS; the rest → Opex.',
    'Notes': '🗒️ Context or decision trail (e.g., vendor, contract length).'
  };
  const lc = sh.getLastColumn(); const hdr = sh.getRange(1, 1, 1, lc);
  const hvals = hdr.getValues()[0]; const notes = hdr.getNotes();
  for (let i = 0; i < hvals.length; i++) {
    const t = map[String(hvals[i] || '').trim()]; if (!t) continue;
    notes[0][i] = t; // overwrite to keep it simple and visible
  }
  hdr.setNotes(notes);

  // Column B row-level help (label meaning): what it is, when/why change, effect on MODEL
  const lr = sh.getLastRow(); if (lr >= 2) {
    const rng = sh.getRange(2, 2, lr - 1, 1); const rn = rng.getNotes(); const labels = rng.getValues();
    for (let r = 0; r < lr - 1; r++) {
      const lab = String(labels[r][0] || '');
      rn[r][0] = `🔎 ${lab}\n• What: Value in ASSUMPTIONS that drives cost/revenue.\n• Where: Feeds MODEL → IS/CF/BS.\n• When: Change if price/volume/vendor changes.\n• Why: Keeps P&L current and trustworthy.\n• Best practice: Map biggest CHF items first.✅`;
    }
    rng.setNotes(rn);

    // Column C (Category) and D (Bucket) implications with beginner wording
    const catR = sh.getRange(2, 3, lr - 1, 1); const buckR = sh.getRange(2, 4, lr - 1, 1);
    const catN = catR.getNotes(); const buckN = buckR.getNotes();
    for (let r = 0; r < lr - 1; r++) {
      const catTag = '🧩 Category (analysis)\n• What: Theme for spend.\n• Why: Clear breakdowns for decisions.\n• Best practice: Infra/AI Runtime/Tools/Media/Services.\n• Tip: Keep naming consistent.';
      const bcTag = '⚖️ Bucket (COGS vs Opex)\n• Rule: Direct to serve revenue → COGS; others → Opex.\n• Why: COGS lowers GM%; Opex affects EBITDA.\n• Best practice: Infra/AI runtime → COGS; Media/Tools → Opex.';
      catN[r][0] = catTag;
      buckN[r][0] = bcTag;
    }
    catR.setNotes(catN); buckR.setNotes(buckN);
  }
}

function enrichModel_() {
  const sh = getSheetByCandidates_(SHEET_CANDIDATES.MODEL); if (!sh) return;
  const lc = sh.getLastColumn(); const hdr = sh.getRange(1, 1, 1, lc);
  // Clear any existing header notes first
  hdr.setNotes([new Array(lc).fill('')]);
  // ADHD-friendly, CFO-style quick guides
  const map = {
    'Month': '📅 Timeline anchor. Tip: Edit Start/End Dates in ASSUMPTIONS only.',
    'In_Horizon': '🎯 1 = in forecast window. Used to pick “latest month”.',
    'Active_Customers': '👥 Paying customers now. Drives MRR.',
    'New_Customers': '➕ Leads × Conversion. Keep conversion realistic (1–5%).',
    'Churned_Customers': '➖ Customers lost (Churn × prior base). Lower is better.',
    'ARPU_CaaS': '💳 Avg revenue per customer. Higher → more MRR.',
    'MRR_Display': '💵 Monthly recurring revenue. Watch trend.',
    'ARR_Display': '📈 12 × MRR. Use for milestones, not cash.',
    'Headcount_Total': '👤 Total team. Align with runway; delay hires if needed.',
    'Payroll_CHF': '🧾 Salaries + overhead. Your biggest cost early.',
    'COGS_CHF': '🧪 Direct costs (infra/AI runtime/fees). Impacts GM%.',
    'Other_Opex_CHF': '🧰 Indirect costs (media/tools/services).',
    'Revenue_Display': '🏷️ Total revenue (MRR + other).',
    'Costs_Display': '💸 Total costs (Payroll + COGS + Opex).',
    'Net_Burn': '🔥 Costs − Revenue (positive = burn).',
    'Cum_Cash': '🏦 Cash balance over time. Aim runway > 10 months.'
  };
  const hvals = hdr.getValues()[0]; const notes = hdr.getNotes();
  for (let i = 0; i < hvals.length; i++) {
    const t = map[String(hvals[i] || '').trim()]; if (!t) continue; notes[0][i] = t;
  }
  hdr.setNotes(notes);
}

function enrichIS_() {
  const sh = getOrInsert_('IS', 'INCOME_STATEMENT'); if (!sh) return;
  const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2) return;
  const labels = sh.getRange(2, 1, lr - 1, 1).getValues(); const rng = sh.getRange(2, 1, lr - 1, 1);
  const notes = rng.getNotes();
  const guide = {
    'Revenue (Total) 💵': 'What: Total monthly revenue.\nMeans: Primary growth KPI.\nWatch: Grow while improving GM%.',
    'COGS (Direct Total) 🧪': 'What: Direct delivery costs.\nMeans: Infra, runtime, fees.\nWatch: Lower COGS → higher GM%.',
    'Gross Profit 💎': 'What: Revenue − COGS.\nMeans: Cash to pay for opex.\nWatch: Grow faster than opex.',
    'Gross Margin % 📈': 'What: GP ÷ Revenue.\nMeans: SaaS usually 60–90%.\nWatch: Improve via pricing/COGS.',
    'Opex (Total) 🧰': 'What: Indirect operating costs.\nMeans: Marketing, tools, services.\nWatch: Keep lean; track ROI.',
    'EBITDA 📊': 'What: Operating profit before non-cash.\nMeans: Early negative is common.\nWatch: Trend toward breakeven.',
    'Depreciation 🧱': 'What: Non-cash expense of assets.\nMeans: From CapEx & lifetime.\nWatch: Set Depreciation months.',
    'EBIT 🧮': 'What: EBITDA − Depreciation.\nMeans: Operating profit metric.\nWatch: A key milestone to positive.',
    'Interest (net) 💳': 'What: Yield on cash.\nMeans: No debt modeled.\nWatch: Small positive only.',
    'EBT 💼': 'What: Pre-tax profit.\nMeans: Before taxes.\nWatch: Should follow EBIT trend.',
    'Taxes 🧾': 'What: Corporate taxes.\nMeans: Uses Tax rate % assumption.\nWatch: Ensure rate is realistic.',
    'Net Income 🟢': 'What: Bottom line.\nMeans: Profit after taxes.\nWatch: Trend and volatility.'
  };
  for (let r = 0; r < labels.length; r++) {
    const lab = String(labels[r][0] || ''); const tag = `— Colibri: ${guide[lab] || 'Interpret this P&L line; ideal values depend on stage.'}`;
    const cur = notes[r][0] || ''; if (cur.indexOf(tag) === -1) notes[r][0] = cur ? `${cur}\n\n${tag}` : tag;
  }
  rng.setNotes(notes);

  // Fix #N/A on rows 21 (Taxes) and 22 (Net Income) by ensuring IFERROR guards
  for (let c = 2; c <= lc; c++) {
    const ebt = sh.getRange(20, c).getA1Notation();
    sh.getRange(21, c).setFormula(`=IFERROR(MAX(0, ${ebt} * IFERROR(INDEX(ASSUMPTIONS!A:A, MATCH("Tax rate %", ASSUMPTIONS!B:B,0))/100,0)),0)`);
    const tax = sh.getRange(21, c).getA1Notation();
    sh.getRange(22, c).setFormula(`=IFERROR(${ebt}-${tax},0)`);
  }
}

function enrichCF_() {
  const sh = getOrInsert_('CF', 'CASH_FLOW'); if (!sh) return; const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2) return;
  const rng = sh.getRange(2, 1, lr - 1, 1); const notes = rng.getNotes(); const labels = rng.getValues();
  const guide = {
    'Net Income 🧾': 'From P&L — starting point for cash.',
    'Non-cash: Depreciation 🧱': 'Add-back of non-cash expense.',
    'Working Capital Δ 🔄': 'Change in AR/AP/Inventory (kept 0 for simplicity).',
    'Operating Cash Flow 💧': 'NI + non-cash + WC.',
    'CapEx 🛠️': 'Cash spent on assets (ASSUMPTIONS → CapEx per month).',
    'Free Cash Flow 🟢': 'Operating CF − CapEx.',
    'Financing (Debt/Equity) 💳': 'External cash in/out (set to 0 placeholder).',
    'Net Cash Change 💱': 'FCF + Financing.',
    'Ending Cash 🏦': 'Cash after change — ties to BS Cash.'
  };
  for (let r = 0; r < labels.length; r++) {
    const lab = String(labels[r][0] || ''); const tag = `— Colibri: ${guide[lab] || 'Interpretation for cash flow line.'}`; const cur = notes[r][0] || ''; if (cur.indexOf(tag) === -1) notes[r][0] = cur ? `${cur}\n\n${tag}` : tag;
  }
  rng.setNotes(notes);

  // Harden key formulas (rows 2,3,10) if present
  try {
    for (let c = 2; c <= lc; c++) {
      // Row 2 Net Income from IS row 22
      const isSh = getOrInsert_('IS', 'INCOME_STATEMENT');
      sh.getRange(2, c).setFormula(`=IFERROR(${isSh.getRange(22, c).getA1Notation()},0)`);
      // Row 3 Depreciation from IS row 17
      sh.getRange(3, c).setFormula(`=IFERROR(${isSh.getRange(17, c).getA1Notation()},0)`);
      // Row 10 Ending Cash mirrors MODEL Cum_Cash at matching date
      const model = getSheetByCandidates_(SHEET_CANDIDATES.MODEL);
      if (model) {
        const cMonth = 1; // first column in CF/Model is Month
        const cCashCol = model.getRange(1, 1, 1, model.getLastColumn()).getValues()[0].indexOf('Cum_Cash') + 1;
        if (cCashCol > 0) {
          sh.getRange(10, c).setFormula(`=IFERROR(INDEX(${model.getSheetName()}!${columnLetter_(cCashCol)}:${columnLetter_(cCashCol)}, MATCH(${sh.getRange(1, c).getA1Notation()}, ${model.getSheetName()}!${columnLetter_(cMonth)}:${columnLetter_(cMonth)},0)),0)`);
        }
      }
    }
  } catch (e) { }
}

function enrichBS_() {
  const sh = getOrInsert_('BS', 'BALANCE_SHEET'); if (!sh) return; const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2) return;
  const rng = sh.getRange(2, 1, lr - 1, 1); const notes = rng.getNotes(); const labels = rng.getValues();
  const guide = {
    'Cash 🏦': 'From CF Ending Cash (MODEL Cum_Cash). If 0: set Starting Cash (CHF) in ASSUMPTIONS, then run ③ and ⑤.',
    'A/R 📬': 'Accounts receivable — 0 by default (DSO off). Enable WC if needed.',
    'Inventory 📦': '0 by default (software).',
    'Prepaids & Other 🗃️': '0 by default; add if prepayments.',
    'PP&E (net) 🏭': 'Prior PP&E + CapEx − Depreciation. If 0: set CapEx per month and Depreciation months in ASSUMPTIONS.',
    'Total Assets 💼': 'Sum of asset lines.',
    'A/P 🧾': '0 by default (DPO off).',
    'Other Liab 📑': '0 by default.',
    'Deferred Rev ⏳': '0 by default; add if pre-billing.',
    'Debt 💳': '0 by default; add if financing.',
    'Total Liabilities 🧮': 'Sum of liability lines.',
    'Equity 📈': 'Assets − Liabilities.',
    'Liabilities + Equity ⚖️': 'Should equal Total Assets.',
    'Balance Check ✅': '0 means balanced.'
  };
  for (let r = 0; r < labels.length; r++) {
    const lab = String(labels[r][0] || ''); const tag = `— Colibri: ${guide[lab] || 'Interpretation for balance sheet line.'}`; const cur = notes[r][0] || ''; if (cur.indexOf(tag) === -1) notes[r][0] = cur ? `${cur}\n\n${tag}` : tag;
  }
  rng.setNotes(notes);

  // Fix #N/A or propagation errors in rows 5,6,12,13,14
  for (let c = 2; c <= lc; c++) {
    const ppePrev = c > 2 ? sh.getRange(5, c - 1).getA1Notation() : '0';
    const dep = `${getOrInsert_('IS', 'INCOME_STATEMENT').getSheetName()}!${getOrInsert_('IS', 'INCOME_STATEMENT').getRange(17, c).getA1Notation()}`;
    const capex = `${getOrInsert_('CF', 'CASH_FLOW').getSheetName()}!${getOrInsert_('CF', 'CASH_FLOW').getRange(6, c).getA1Notation()}`;
    sh.getRange(5, c).setFormula(`=IFERROR(${ppePrev}+IFERROR(${capex},0)-IFERROR(${dep},0),0)`);
    const totalAssets = `SUM(${sh.getRange(2, c).getA1Notation()}:${sh.getRange(5, c).getA1Notation()})`;
    sh.getRange(6, c).setFormula(`=IFERROR(${totalAssets},0)`);
    const totalLiab = `SUM(${sh.getRange(7, c).getA1Notation()}:${sh.getRange(10, c).getA1Notation()})`;
    sh.getRange(11, c).setFormula(`=IFERROR(${totalLiab},0)`);
    sh.getRange(12, c).setFormula(`=IFERROR(${sh.getRange(6, c).getA1Notation()}-${sh.getRange(11, c).getA1Notation()},0)`);
    sh.getRange(13, c).setFormula(`=IFERROR(${sh.getRange(11, c).getA1Notation()}+${sh.getRange(12, c).getA1Notation()},0)`);
    sh.getRange(14, c).setFormula(`=IFERROR(${sh.getRange(6, c).getA1Notation()}-${sh.getRange(13, c).getA1Notation()},0)`);
  }
}

function enrichSummary_() {
  const sh = getOrInsert_('SUMMARY', 'SUMMARY'); if (!sh) return; const lr = sh.getLastRow(); if (lr < 2) return;
  const hdr = sh.getRange(1, 1, 1, 3); const hnotes = hdr.getNotes();
  const map = { 'Metric': 'What this KPI represents', 'Value': 'Latest value — look at trend', 'Description': 'Interpretation and threshold (A6 > 10 months means healthy runway)' };
  const hvals = hdr.getValues()[0]; for (let i = 0; i < hvals.length; i++) { const t = map[String(hvals[i] || '').trim()]; if (!t) continue; const tag = `— Colibri: ${t}`; const cur = hnotes[0][i] || ''; if (cur.indexOf(tag) === -1) hnotes[0][i] = cur ? `${cur}\n\n${tag}` : tag; }
  hdr.setNotes(hnotes);
  // Add cell-level notes only on Column A for non-empty KPI labels. Clear B & C notes.
  const labels = sh.getRange(2, 1, lr - 1, 1).getValues();
  const notesA = sh.getRange(2, 1, lr - 1, 1).getNotes();
  for (let r = 0; r < labels.length; r++) {
    const lab = String(labels[r][0] || '').trim();
    if (!lab) continue; // Skip empty rows to avoid notes down the sheet
    let tip = 'This KPI helps you track financial health. Focus on trend over time.';
    if (/latest month/i.test(lab)) tip = 'The most recent month in your forecast horizon.';
    else if (/mrr/i.test(lab)) tip = 'Monthly Recurring Revenue: recurring part of revenue. Aim to grow MRR steadily.';
    else if (/net burn/i.test(lab)) tip = 'Costs minus Revenue. Positive = burning cash. Work to reduce burn or increase revenue.';
    else if (/runway/i.test(lab)) tip = 'How many months of cash you have left at current burn. Aim for > 10 months. Extend by cutting costs, raising revenue, or new funding.';
    const cur = notesA[r][0] || ''; const tag = `— Colibri: ${tip}`;
    if (cur.indexOf(tag) === -1) notesA[r][0] = cur ? `${cur}\n\n${tag}` : tag;
  }
  sh.getRange(2, 1, lr - 1, 1).setNotes(notesA);
  // Clear any legacy notes in columns B..last for all rows (avoid black popups down the sheet)
  try {
    const maxRows = Math.max(sh.getMaxRows ? sh.getMaxRows() : lr, lr);
    const lastCol = Math.max(sh.getLastColumn(), 2);
    const colsToClear = Math.max(lastCol - 1, 1);
    sh.getRange(2, 2, Math.max(maxRows - 1, 1), colsToClear).clearNote();
  } catch (e) { }
  // Remove broken charts
  try { const model = getSheetByCandidates_(SHEET_CANDIDATES.MODEL); if (model) { model.getCharts().forEach(c => { const t = (c.getOptions() && c.getOptions().get('title')) || ''; if (String(t).toLowerCase().indexOf('revenue vs costs') > -1) model.removeChart(c); }); } } catch (e) { }

  // Ensure B4 (Net Burn) and B5 (Runway) are robust
  try {
    const model = getSheetByCandidates_(SHEET_CANDIDATES.MODEL); if (model) {
      const headers = model.getRange(1, 1, 1, model.getLastColumn()).getValues()[0];
      const cMonth = headers.indexOf('Month') + 1; const cBurn = headers.indexOf('Net_Burn') + 1; const cCash = headers.indexOf('Cum_Cash') + 1;
      if (cMonth > 0 && cBurn > 0 && cCash > 0) {
        sh.getRange(4, 2).setFormula(`=IFERROR(INDEX(${model.getSheetName()}!${columnLetter_(cBurn)}:${columnLetter_(cBurn)}, MATCH(${sh.getRange(2, 2).getA1Notation()}, ${model.getSheetName()}!${columnLetter_(cMonth)}:${columnLetter_(cMonth)},0)),0)`);
        sh.getRange(5, 2).setFormula(`=IFERROR( INDEX(${model.getSheetName()}!${columnLetter_(cCash)}:${columnLetter_(cCash)}, MATCH(${sh.getRange(2, 2).getA1Notation()}, ${model.getSheetName()}!${columnLetter_(cMonth)}:${columnLetter_(cMonth)},0)) / MAX(0.0001, INDEX(${model.getSheetName()}!${columnLetter_(cBurn)}:${columnLetter_(cBurn)}, MATCH(${sh.getRange(2, 2).getA1Notation()}, ${model.getSheetName()}!${columnLetter_(cMonth)}:${columnLetter_(cMonth)},0))), 0)`);
      }
    }
  } catch (e) { }

  // Add a small chart on SUMMARY for MRR vs Costs
  try {
    const model = getSheetByCandidates_(SHEET_CANDIDATES.MODEL); if (!model) return;
    const mHdr = model.getRange(1, 1, 1, model.getLastColumn()).getValues()[0];
    const cMRR = mHdr.indexOf('MRR_Display') + 1; const cCosts = mHdr.indexOf('Costs_Display') + 1;
    if (cMRR > 0 && cCosts > 0) {
      const charts = sh.getCharts(); charts.forEach(c => { const t = (c.getOptions() && c.getOptions().get('title')) || ''; if (String(t).toLowerCase().indexOf('mrr vs costs') > -1) sh.removeChart(c); });
      const lrM = model.getLastRow();
      const chart = sh.newChart()
        .asLineChart()
        .setOption('title', 'MRR vs Costs')
        .addRange(model.getRange(1, 1, lrM, 1))
        .addRange(model.getRange(1, cMRR, lrM, 1))
        .addRange(model.getRange(1, cCosts, lrM, 1))
        .setPosition(1, 5, 0, 0)
        .build();
      sh.insertChart(chart);
      // Add ARR chart next to it
      const charts2 = sh.getCharts(); charts2.forEach(c => { const t = (c.getOptions() && c.getOptions().get('title')) || ''; if (String(t).toLowerCase().indexOf('arr') > -1) sh.removeChart(c); });
      const cARR = mHdr.indexOf('ARR_Display') + 1;
      if (cARR > 0) {
        // Create a local copy range for ARR to avoid chart instabilities
        try {
          const tmpStartCol = 20; // T column region in SUMMARY for temp data
          const lrCopy = lrM;
          // Ensure SUMMARY has enough rows to host the copied data
          const needRows = Math.max(lrCopy, 2);
          const curMax = sh.getMaxRows ? sh.getMaxRows() : sh.getLastRow();
          if (curMax < needRows) sh.insertRowsAfter(curMax, needRows - curMax);
          sh.getRange(1, tmpStartCol, lrCopy, 1).setValues(model.getRange(1, 1, lrCopy, 1).getValues());
          sh.getRange(1, tmpStartCol + 1, lrCopy, 1).setValues(model.getRange(1, cARR, lrCopy, 1).getValues());
          sh.getRange(1, tmpStartCol, 1, 2).setValues([["Month", "ARR"]]);
          const chart2 = sh.newChart()
            .asLineChart()
            .setOption('title', 'ARR')
            .addRange(sh.getRange(1, tmpStartCol, lrCopy, 1))
            .addRange(sh.getRange(1, tmpStartCol + 1, lrCopy, 1))
            .setPosition(15, 5, 0, 0)
            .build();
          sh.insertChart(chart2);
        } catch (e) { }
      }
    }
  } catch (e) { }

  // Clean stray empty labels with notes in A6–A10
  try {
    for (let r = 6; r <= Math.min(10, sh.getLastRow()); r++) {
      const cell = sh.getRange(r, 1);
      const txt = String(cell.getValue() || '').trim();
      if (!txt) cell.setNote('');
    }
  } catch (e) { }

  // Add dynamic explanation note on B5 (Runway) to explain causes of extreme values
  try {
    const b5 = sh.getRange(5, 2);
    const v = Number(b5.getValue() || 0);
    let msg = 'Runway = Cash / Avg Burn (last 3 months). Aim > 10 months.';
    if (v > 120) msg += '\n⚠️ Very high: likely near-zero burn or huge cash. Check MODEL Net_Burn and ASSUMPTIONS costs (COGS %, payroll, opex).';
    if (v <= 0) msg += '\n⚠️ Non-positive: burn negative or cash ≤ 0. Check Starting Cash, Net_Burn sign.';
    b5.setNote(`🧠 ${msg}`);
  } catch (e) { }
}

/* -------------------- Assumptions Coach & Quick Fix -------------------- */
function showAssumptionsCoach() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(`
  <div style="font-family: Inter, system-ui, -apple-system, Segoe UI, Roboto; line-height:1.4;">
    <h2>🟡 ASSUMPTIONS — Friendly Coach</h2>
    <p>We’ll set the few numbers that drive everything. Do it in this order, and you’ll feel in control:</p>
    <p>1) 💰 Starting Cash — type what’s in the bank today. This seeds your cash line and runway.</p>
    <p>2) 🗓️ Start & End Dates — pick a start month and plan 24–60 months. This creates the timeline.</p>
    <p>3) 📦 Revenue — use simple, believable inputs:
    <br>• 💳 ARPU (e.g., 100–600)
    <br>• 🧲 Monthly leads (starts small)
    <br>• 🎯 Conversion 0.01–0.05 (1–5%)
    <br>• 🔁 Churn 0.01–0.05 (1–5%)</p>
    <p>4) 👥 Headcount — set “Starting Headcount” and a sensible 3‑year target. Add dated hires with level (L1–L4). Levels map to salary bands automatically.</p>
    <p>5) 🧰 Opex — tools per FTE, media/PR, product/R&D, services. Change the big numbers first; small ones can wait.</p>
    <p>6) 🧪 COGS — set “Cost of Revenue %” and any AI runtime cost per active customer. This controls Gross Margin.</p>
    <p>7) 🏗️ CapEx & 🧱 Depreciation — if you invest monthly, set CapEx and a lifetime (e.g., 36 months).</p>
    <p>How to iterate: change one group → run ③ (MODEL) then ⑤ (IS/CF/BS) → check SUMMARY and Diagnostics. That’s it.</p>
    <p>Goal: runway > 10 months, MRR trending up, GM% realistic (not 99%).</p>
  </div>
  `).setTitle('Assumptions Coach').setWidth(520).setHeight(580);
  ui.showModelessDialog(html, 'Assumptions Coach');
}

function newCoach_(title, bulletsHtml) {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(`
  <div style="font-family: Inter, system-ui, -apple-system, Segoe UI, Roboto; line-height:1.4; max-width:640px;">
    <h2>🟡 ${title}</h2>
    ${bulletsHtml}
    <p style="margin-top:12px;">Tip: Change one group → run ③ then ⑤ → check SUMMARY runway and Diagnostics.</p>
  </div>`).setTitle(title).setWidth(560).setHeight(600);
  ui.showModelessDialog(html, `${title}`);
}

function showModelCoach() {
  newCoach_('MODEL — Friendly Coach', `
  <p>Think of MODEL as a calculator. It turns your assumptions into monthly numbers.</p>
  <p>• 📅 Month and In_Horizon come from your Start/End Dates — change dates only in ASSUMPTIONS.</p>
  <p>• 👥 Customers + 💳 ARPU create 💵 MRR and 📈 ARR. Keep conversion/churn realistic so the curve feels honest.</p>
  <p>• 👤 Headcount grows with your plan. If runway looks tight, delay hires and re-run.</p>
  <p>• 💸 Costs = Payroll + COGS + Opex → this gives 🔥 Net Burn and 🏦 Cum_Cash.</p>
  <p>If Cum_Cash ever dips negative, you’ll run out of cash by that month — adjust and try again.</p>`);
}

function showSummaryCoach() {
  newCoach_('SUMMARY — Friendly Coach', `
  <p>SUMMARY is your dashboard.</p>
  <p>• 🗓️ Latest Month — the last month inside your horizon.</p>
  <p>• 💵 MRR — the recurring revenue today. We want an up‑and‑to‑the‑right line.</p>
  <p>• 🔥 Net Burn — costs minus revenue. Positive burn means you’re spending more than you make.</p>
  <p>• 🛟 Runway — cash divided by the average burn of the last 3 months. If it’s huge, burn is probably near 0. If it’s tiny, either low cash or high burn.</p>
  <p>Use this page to sanity‑check changes fast.</p>`);
}

function showISCoach() {
  newCoach_('Income Statement — Friendly Coach', `
  <p>This is your monthly profit story.</p>
  <p>• 💵 Revenue comes from MODEL (MRR + services).<br>
     • 🧪 COGS are the direct costs to deliver (infra, AI runtime, fees). Lower COGS → better margins.<br>
     • 💎 Gross Profit and 📈 Gross Margin % tell you quality — a healthy SaaS lands 60–90% depending on model.<br>
     • 🧰 Opex is everything else to run the company. Keep it focused and justified.</p>
  <p>Below that: 📊 EBITDA, 🧱 Depreciation, 🧮 EBIT, 💼 EBT, 🧾 Taxes, and finally 🟢 Net Income.</p>`);
}

function showCFCoach() {
  newCoach_('Cash Flow — Friendly Coach', `
  <p>Cash is survival. Start with 🧾 Net Income, add back 🧱 Depreciation (it’s non‑cash), keep 🔄 Working Capital simple (0 unless you add timing), subtract 🛠️ CapEx, and you get 💱 Net change in cash. 🏦 Ending Cash feeds the Balance Sheet.</p>`);
}

function showBSCoach() {
  newCoach_('Balance Sheet — Friendly Coach', `
  <p>Think of this as a snapshot of what you own and owe.</p>
  <p>• 🏦 Cash comes straight from Cash Flow (MODEL Cum_Cash).<br>
     • 🏭 PP&E grows with CapEx and shrinks with Depreciation.<br>
     • ⚖️ Assets must equal Liabilities + Equity. Balance Check should be 0.</p>
  <p>If it’s not balancing, rebuild ⑤ and check CapEx/Depreciation inputs.</p>`);
}

function showMappingCoach() {
  newCoach_('Mapping — Friendly Coach', `
  <p>Mapping tells the model where a cost should land.</p>
  <p>• 🧩 Category is for analysis (Infra, AI Runtime, Tools, Media, Services).<br>
     • ⚖️ Bucket decides P&L placement: direct delivery → COGS; everything else → Opex.</p>
  <p>Start with the biggest numbers first; that’s where clarity matters most. Add a note with vendor and contract details so future‑you remembers why.</p>`);
}

function showReadmeCoach() {
  newCoach_('README — Friendly Coach', `
  <p>Use the Colibri menu as your control panel. Actions are safe to re‑run. When something looks odd, open Diagnostics and the Formula Errors report — they’ll point you to the next fix.</p>`);
}

function showFormulaLibCoach() {
  newCoach_('Formula Library — Friendly Coach', `
  <p>Here you’ll find small examples (like INDEX/MATCH by label) that make the model robust when rows move. If a cell errors, just paste a text example and add a short note.</p>`);
}

function applySensibleDefaults() {
  const ass = getOrInsert_('ASSUMPTIONS', 'ASSUMPTIONS');
  const setIfEmpty = (label, val) => { const r = findLabelRow_(ass, label); if (r) { const cell = ass.getRange(r, 1); const v = cell.getValue(); if (v === '' || v === null || (typeof v === 'number' && v === 0)) cell.setValue(val); } };
  setIfEmpty('ARPU – CaaS monthly (CHF)', 250);
  setIfEmpty('Monthly leads', 200);
  setIfEmpty('Awareness→Conversion (0–1)', 0.03);
  setIfEmpty('Churn monthly (0–1)', 0.03);
  setIfEmpty('Starting Headcount', 3);
  setIfEmpty('Target Headcount (by year 3)', 12);
  setIfEmpty('Payroll Overhead %', 18);
  setIfEmpty('Tech stack per FTE / month (CHF)', 150);
  setIfEmpty('Culture & Learning per FTE / month (CHF)', 100);
  setIfEmpty('Media / PR monthly (CHF)', 1500);
  setIfEmpty('R&D AIX monthly (CHF)', 5000);
  setIfEmpty('Product dev (CaaS) monthly (CHF)', 5000);
  setIfEmpty('Purchased Services – Micro monthly (CHF)', 1000);
  setIfEmpty('Starting Cash (CHF)', 150000);
  Browser.msgBox('Applied sensible defaults where values were blank/zero. Re-run ③ then ⑤.');
}

function enrichFormulaLib_() {
  const sh = getOrInsert_('FORMULA_LIB', 'FORMULA_LIBRARY'); if (!sh) return;
  const lr = sh.getLastRow(); if (lr < 1) { sh.getRange('A1').setValue('Formula'); sh.getRange('B1').setValue('Example'); sh.getRange('C1').setValue('Explanation'); }
  const lc = sh.getLastColumn(); const hdr = sh.getRange(1, 1, 1, Math.max(3, lc));
  const map = { 'Formula': 'Named function or construct', 'Example': 'How to use in sheet', 'Explanation': 'Beginner-friendly explanation' };
  const hvals = hdr.getValues()[0]; const notes = hdr.getNotes();
  for (let i = 0; i < hvals.length; i++) { const t = map[String(hvals[i] || '').trim()]; if (!t) continue; const tag = `— Colibri: ${t}`; const cur = notes[0][i] || ''; if (cur.indexOf(tag) === -1) notes[0][i] = cur ? `${cur}\n\n${tag}` : tag; }
  hdr.setNotes(notes);
  // Fix B5 UX: show example as note or as visible text based on toggle
  const lr2 = sh.getLastRow(); if (lr2 >= 5) {
    const cell = sh.getRange(5, 2);
    const mode = getDocProp_('FORMULA_LIB_SHOW_EXAMPLES', 'note'); // 'note' | 'text'
    const example = 'INDEX(Sheet!A:A, MATCH("Label", Sheet!B:B, 0))';
    try { if (cell.getFormula()) cell.setFormula(''); } catch (e) { }
    try { cell.clearDataValidations(); } catch (e) { }
    cell.setNumberFormat('@');
    if (mode === 'text') {
      // Show as visible text (still not a formula)
      cell.setValue(`Example: ${example}`);
      cell.setNote('📚 Example shown as text. To try it, copy to a test cell and add = at the start. In EU locales use ; instead of ,');
    } else {
      // Note-only
      cell.setValue('');
      const note = '📚 Example (kept as a note to prevent parse issues):\\n' + example +
        '\\n\\nTip: To try it, paste into a safe cell and add = at the start. If your locale uses ;, replace commas with semicolons.';
      cell.setNote(note);
    }
  }
}

/* -------------------- 8) EU formatting + comment cleanup -------------------- */
function applyEUFormattingAndCleanComments() {
  const ss = SpreadsheetApp.getActive();
  // 1) Set spreadsheet locale to a European locale with decimal comma & dot thousands
  try { setSpreadsheetLocaleEU_('fr-FR'); } catch (e) { /* optional */ }

  // 2) Normalize numeric strings and apply formats on key sheets
  const sheets = [
    getOrInsert_('ASSUMPTIONS', 'ASSUMPTIONS'),
    getSheetByCandidates_(SHEET_CANDIDATES.MODEL),
    getOrInsert_('SUMMARY', 'SUMMARY'),
    getOrInsert_('IS', 'INCOME_STATEMENT'),
    getOrInsert_('CF', 'CASH_FLOW'),
    getOrInsert_('BS', 'BALANCE_SHEET'),
    getOrInsert_('MAP', 'MAPPING_CATEGORIES')
  ].filter(Boolean);

  sheets.forEach(sh => normalizeNumericStringsRange_(sh.getDataRange()));
  formatAssumptionsEU_();
  formatModelEU_();
  formatSummaryEU_();
  formatISEU_();
  formatCFEU_();
  formatBSEU_();

  // 3) Remove all Google Drive comments (yellow triangles) from this file
  try { removeAllComments_(); } catch (e) { /* may require scopes */ }

  // 4) Re-add helpful header notes (black triangle) and long notes optionally
  headerTitleNotes_();
  enrichModel_(); // ensure MODEL header notes are reapplied
  try { enrichFormulaLib_(); } catch (e) { }
  // appendLongNotesEverywhere(); // optional if you prefer minimal notes

  // 5) Apply consistent 16pt font size everywhere for readability
  try { ss.getSheets().forEach(sh => { try { sh.getDataRange().setFontSize(16); } catch (e) { } }); } catch (e) { }

  // 6) Make headers readable: color row 1 with tab color, freeze it, wrap text, and align
  try {
    ss.getSheets().forEach(sh => {
      const lc = Math.max(1, sh.getLastColumn());
      const header = sh.getRange(1, 1, 1, lc);
      const tabColor = sh.getTabColor() || '#E0E0E0';
      header.setBackground(tabColor).setFontWeight('bold');
      sh.setFrozenRows(1);
      // Wrap all cells and align headers center, data top-left
      sh.getDataRange().setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('left');
      header.setVerticalAlignment('middle').setHorizontalAlignment('center');
      // Auto-resize columns to fit wrapped content better (safe attempt)
      try { sh.autoResizeColumns(1, lc); } catch (e) { /* ignore */ }
    });
  } catch (e) { /* ignore formatting issues */ }

  ss.toast('EU formatting applied, numeric strings normalized, comments removed ✅', 'Colibri', 6);
}

// Optional utility to repair formatting quickly if an older script changed appearances
function repairFormatting() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sh => {
    try {
      const lc = Math.max(1, sh.getLastColumn());
      sh.getDataRange().setFontSize(16);
      const color = sh.getTabColor() || '#E0E0E0';
      sh.getRange(1, 1, 1, lc).setBackground(color).setFontWeight('bold');
      sh.setFrozenRows(1);
    } catch (e) { }
  });
  headerTitleNotes_();
  enrichModel_();
  SpreadsheetApp.getActive().toast('Formatting repaired ✅', 'Colibri', 4);
}

function setSpreadsheetLocaleEU_(locale) {
  // Requires Advanced Sheets service
  const id = SpreadsheetApp.getActive().getId();
  Sheets.Spreadsheets.batchUpdate({
    requests: [{ updateSpreadsheetProperties: { properties: { locale }, fields: 'locale' } }]
  }, id);
}

function normalizeNumericStringsRange_(rng) {
  const vals = rng.getValues();
  let changed = false;
  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[0].length; c++) {
      const v = vals[r][c];
      if (typeof v === 'string') {
        const s = v.replace(/\u00A0/g, ' ').trim(); // normalize nbsp
        // Accept numbers with thousand apostrophes or spaces and comma/dot decimals
        const s2 = s.replace(/'/g, '').replace(/\s+/g, '');
        // If looks like a number, parse with a tolerant rule: prefer comma as decimal if both present
        if (/^[+-]?[0-9]+([\.,][0-9]+)?$/.test(s2)) {
          let numStr = s2;
          if (numStr.indexOf(',') > -1 && numStr.indexOf('.') === -1) { numStr = numStr.replace(',', '.'); }
          const num = parseFloat(numStr);
          if (!isNaN(num)) { vals[r][c] = num; changed = true; }
        }
      }
    }
  }
  if (changed) rng.setValues(vals);
}

function formatAssumptionsEU_() {
  const ass = getOrInsert_('ASSUMPTIONS', 'ASSUMPTIONS');
  const lr = ass.getLastRow(); if (lr < 2) return;
  const labels = ass.getRange(2, 2, lr - 1, 1).getValues().map(r => String(r[0] || ''));
  const valsR = ass.getRange(2, 1, lr - 1, 1);
  const vals = valsR.getValues();
  for (let i = 0; i < labels.length; i++) {
    const lab = labels[i]; const row = i + 2; const cell = ass.getRange(row, 1);
    const isDate = /start date|end date/i.test(lab);
    const isRatio = /(0\s*[–-]\s*1|0-1|churn|conversion|growth)/i.test(lab);
    const isPercent = /%|tax rate|yield/i.test(lab);
    const isCurrency = /(CHF|\$|price|capex|budget|monthly|yearly|salary|cost|fees)/i.test(lab) && !isPercent && !isRatio;
    const isCount = /(headcount|customers|days|months)/i.test(lab) && !isCurrency && !isPercent;
    if (isDate) cell.setNumberFormat('yyyy-mm-dd');
    else if (isRatio) cell.setNumberFormat('0.000');
    else if (isPercent) cell.setNumberFormat('0.0%');
    else if (isCurrency) cell.setNumberFormat('# ##0');
    else if (isCount) cell.setNumberFormat('0');
  }
}

function formatModelEU_() {
  const sh = getSheetByCandidates_(SHEET_CANDIDATES.MODEL); if (!sh) return;
  const lc = sh.getLastColumn(), lr = sh.getLastRow(); if (lr < 2) return;
  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const idx = h => headers.indexOf(h) + 1;
  const setColFmt = (name, fmt) => { const c = idx(name); if (c > 0) sh.getRange(2, c, Math.max(lr - 1, 1), 1).setNumberFormat(fmt); };
  setColFmt('Month', 'yyyy-mm-dd');
  ['In_Horizon', 'Active_Customers', 'New_Customers', 'Churned_Customers', 'Headcount_Total'].forEach(n => setColFmt(n, '0'));
  ['ARPU_CaaS', 'MRR_Display', 'ARR_Display', 'Payroll_CHF', 'COGS_CHF', 'Other_Opex_CHF', 'Revenue_Display', 'Costs_Display', 'Net_Burn', 'Cum_Cash'].forEach(n => setColFmt(n, '# ##0'));
}

function formatSummaryEU_() {
  const sh = getOrInsert_('SUMMARY', 'SUMMARY'); const lr = sh.getLastRow(); if (lr < 2) return;
  // Values in column B: default integer, runway months keep one decimal
  sh.getRange(2, 2, lr - 1, 1).setNumberFormat('# ##0');
  // Find Runway row (label in col A)
  const labels = sh.getRange(2, 1, lr - 1, 1).getValues();
  for (let i = 0; i < labels.length; i++) if (String(labels[i][0] || '').toLowerCase().indexOf('runway') > -1) sh.getRange(2 + i, 2).setNumberFormat('0.0');
}

function formatISEU_() {
  const sh = getOrInsert_('IS', 'INCOME_STATEMENT'); const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2) return;
  // All amount rows except GM% row
  // Row indices relative to row 2: 0-based lines array in builder; GM% is row 8 overall -> index 7 ; but here, we’ll detect by label in column A
  const labels = sh.getRange(2, 1, lr - 1, 1).getValues();
  for (let r = 0; r < labels.length; r++) {
    const lab = String(labels[r][0] || '');
    const rng = sh.getRange(2 + r, 2, 1, Math.max(lc - 1, 1));
    if (/Gross Margin %/i.test(lab)) rng.setNumberFormat('0.0%'); else rng.setNumberFormat('# ##0');
  }
}

function formatCFEU_() {
  const sh = getOrInsert_('CF', 'CASH_FLOW'); const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2) return;
  sh.getRange(2, 2, Math.max(lr - 1, 1), Math.max(lc - 1, 1)).setNumberFormat('# ##0');
}

function formatBSEU_() {
  const sh = getOrInsert_('BS', 'BALANCE_SHEET'); const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2) return;
  sh.getRange(2, 2, Math.max(lr - 1, 1), Math.max(lc - 1, 1)).setNumberFormat('# ##0');
}

function removeAllComments_() {
  // Requires Advanced Drive service
  const fileId = SpreadsheetApp.getActive().getId();
  const res = Drive.Comments.list(fileId);
  const comments = res && res.items ? res.items : (res.comments || []);
  if (comments && comments.forEach) { comments.forEach(c => { try { Drive.Comments.remove(fileId, c.commentId || c.id); } catch (e) { } }); }
}
