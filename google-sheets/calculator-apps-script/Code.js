/********* Colibri — Bugfix Pack: Long Notes + Safe Color Scales + Mapping + Sections (CH) *********/
const SHEET_CANDIDATES = {
  MODEL: ['MODEL','Financial Model','FINANCIAL MODEL','FINANCIAL_MODEL'],
  ASSUMPTIONS: ['ASSUMPTIONS','Assumptions'],
  SUMMARY: ['SUMMARY','Summary'],
  README: ['README','Readme','READ ME'],
  FORMULA_LIB: ['FORMULA_LIBRARY','Formula Library','FORMULA LIBRARY','Formula_Library'],
  GROWTH: ['Growth Hypothesis','GROWTH HYPOTHESIS','Growth_Hypothesis'],
  COST: ['cost hypothesis','Cost Hypothesis','COST HYPOTHESIS','cost_hypothesis'],
  MAP: ['MAPPING_CATEGORIES','Mapping','MAP'],
  IS: ['INCOME_STATEMENT','Income Statement','P&L'],
  CF: ['CASH_FLOW','Cash Flow','Cashflow'],
  BS: ['BALANCE_SHEET','Balance Sheet','BS']
};

/* -------------------- utilities -------------------- */
function normalize_(s){ return String(s||'').toLowerCase().replace(/\s|[_-]+/g,''); }
function getSheetByCandidates_(cands){
  const ss=SpreadsheetApp.getActive(), all=ss.getSheets();
  for (const c of cands){ const t=normalize_(c); for (const s of all){ if (normalize_(s.getName())===t) return s; } }
  for (const s of all){ const n=normalize_(s.getName()); for (const c of cands){ if (n.includes(normalize_(c))) return s; } }
  return null;
}
function getOrInsert_(key, fallback){ return getSheetByCandidates_(SHEET_CANDIDATES[key]) || SpreadsheetApp.getActive().insertSheet(fallback); }
function columnLetter_(col){ let t=''; while(col>0){ let r=(col-1)%26; t=String.fromCharCode(65+r)+t; col=Math.floor((col-1)/26);} return t; }

// Non-destructive note append (single cell)
function appendNoteCell_(cell, text){
  const tag = `— Colibri: ${text}`;
  const cur = cell.getNote() || '';
  if (cur.indexOf(tag) !== -1) return;
  cell.setNote(cur ? `${cur}\n\n${tag}` : tag);
}
// Non-destructive for a rectangle
function appendNotesRect_(rng, text, onlyIfHasValue=false){
  const tag = `— Colibri: ${text}`;
  const vals=rng.getValues(), notes=rng.getNotes();
  for (let r=0;r<rng.getNumRows();r++){
    for (let c=0;c<rng.getNumColumns();c++){
      if (onlyIfHasValue && (vals[r][c]===null || vals[r][c]==='')) continue;
      const cur=notes[r][c]||'';
      if (cur.indexOf(tag)===-1) notes[r][c]=cur?`${cur}\n\n${tag}`:tag;
    }
  }
  rng.setNotes(notes);
}
function blockNote_(o){ return [
  `🧩 What: ${o.what}`,
  `📄 Source: ${o.source}`,
  `🧮 How: ${o.how}`,
  `🇨🇭 Typical in CH: ${o.typical}`,
  `🕹️ When to change: ${o.when}`,
  `⚠️ Warnings: ${o.warn}`
].join('\n'); }

/* -------------------- menu -------------------- */
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('Colibri')
    .addItem('① Normalize Dates & Sections', 'normalizeDatesAndSections')
    .addItem('② Build/Refresh Mapping', 'buildOrRefreshMapping')
    .addItem('③ Build / Refresh Financial Statements', 'buildFinancialStatements')
    .addItem('④ Append Long Notes (everywhere)', 'appendLongNotesEverywhere')
    .addItem('⑤ Apply MODEL Color Scales (blue)', 'applyModelColorScales')
    .addSeparator()
    .addItem('Diagnostics', 'diagnostics')
    .addToUi();
  SpreadsheetApp.getActive().toast('Colibri menu ready ✅','Colibri',5);
}

/* -------------------- 1) Dates + Sections on ASSUMPTIONS -------------------- */
function normalizeDatesAndSections(){
  const ass=getOrInsert_('ASSUMPTIONS','ASSUMPTIONS');
  // ensure columns A/B/C exist
  ass.getRange('A1').setValue('Value');
  ass.getRange('B1').setValue('Assumption');
  ass.getRange('C1').setValue('Source / Notes');
  // Ensure Section column (D)
  ass.getRange('D1').setValue('Section');

  // Start/End base
  ensureAssumptionValueWithNote_(ass,'Start Date (yyyy-mm-dd)', new Date(2025,10,1), {
    what:'Start month of forecast.',
    source:'Timeline for MODEL.',
    how:'Month_1 = Start; next months use EDATE(prev,1).',
    typical:'First of any month.',
    when:'Change to shift the entire horizon.',
    warn:'Blank or invalid date breaks the timeline.'
  });
  const startRow = findLabelRow_(ass,'Start Date (yyyy-mm-dd)');
  const start = ass.getRange(startRow,1).getValue();
  const end = new Date(start); end.setMonth(end.getMonth()+72);
  ensureAssumptionValueWithNote_(ass,'End Date (yyyy-mm-dd)', end, {
    what:'Last forecast month.',
    source:'Timeline for MODEL.',
    how:'In_Horizon = 1 between Start and End.',
    typical:'5–7 years.',
    when:'Extend for longer planning.',
    warn:'End < Start truncates series.'
  });

  // Section tagging based on your flow
  const sectionRules = buildSectionRules_();
  const last=ass.getLastRow(); if (last<2) return;
  const labels=ass.getRange(2,2,last-1,1).getValues().map(r=>String(r[0]||'').trim());
  const sectCol=ass.getRange(2,4,last-1,1);
  const sectVals=sectCol.getValues();
  for (let i=0;i<labels.length;i++){
    sectVals[i][0] = sectionForLabel_(labels[i], sectionRules);
  }
  sectCol.setValues(sectVals);

  // Colour band sections for scannability (pastel backgrounds)
  const colors = {
    'Customer Journey':'#FFF7E6',
    'Revenue & Customer Metrics':'#E8F4FD',
    'CAC & Efficiency':'#FFF9C4',
    'Revenue → Profit':'#E8F5E9',
    'Margins':'#F3E5F5',
    'Cash Flow & Balance Sheet':'#F1F8E9',
    'Investments':'#EDE7F6',
    'Other':'#F5F5F5'
  };
  for (let i=0;i<labels.length;i++){
    const clr = colors[sectVals[i][0]] || '#FFFFFF';
    ass.getRange(2+i,1,1,4).setBackground(clr);
  }
  ass.setFrozenRows(1);
  appendNoteCell_(ass.getRange('D1'),'📚 Logical grouping that mirrors your flowchart.');
  SpreadsheetApp.getActive().toast('Dates normalised + Sections applied ✅','Colibri',5);
}
function buildSectionRules_(){
  return [
    {section:'Customer Journey', keys:['Monthly leads','Awareness→Conversion']},
    {section:'Revenue & Customer Metrics', keys:['ARPU','LTV months','Gross Margin %','Churn','Initial CaaS customers','AIX Monthly price','AIX Yearly price']},
    {section:'CAC & Efficiency', keys:['CAC blended']},
    {section:'Revenue → Profit', keys:['Cost of Revenue % (COGS)','AI Agents cost per active customer per month (CHF)','Day rate – training','Training days','Other digital']},
    {section:'Margins', keys:['Gross Margin %']},
    {section:'Cash Flow & Balance Sheet', keys:['Starting Cash','DSO','DPO','DIO']},
    {section:'Investments', keys:['CapEx','Depreciation']},
  ];
}
function sectionForLabel_(label, rules){
  const s = String(label||'');
  for (const r of rules){
    if (r.keys.some(k=>s.indexOf(k)>-1)) return r.section;
  }
  return 'Other';
}
function ensureAssumptionValueWithNote_(ass, label, value, note){
  const r=findOrCreateLabel_(ass,label);
  if (!ass.getRange(r,1).getValue()) ass.getRange(r,1).setValue(value);
  appendNoteCell_(ass.getRange(r,2), blockNote_(note));
  appendNoteCell_(ass.getRange(r,1), '🟢 Edit here. This value drives the model.');
}
function findLabelRow_(ass,label){
  const last=ass.getLastRow(); if (last<2) return null;
  const labels=ass.getRange(2,2,last-1,1).getValues().map(r=>String(r[0]||''));
  const idx=labels.findIndex(x=>x===label);
  return idx===-1?null:(2+idx);
}
function findOrCreateLabel_(ass, label){
  let r=findLabelRow_(ass,label);
  if (r) return r;
  const nr=ass.getLastRow()+1; ass.getRange(nr,2).setValue(label); return nr;
}

/* -------------------- 2) Mapping (drivers → categories & bucket) -------------------- */
function buildOrRefreshMapping(){
  const ass=getOrInsert_('ASSUMPTIONS','ASSUMPTIONS');
  const mapSh=getOrInsert_('MAP','MAPPING_CATEGORIES'); mapSh.clear();
  const header=['Source','Label','Category','Bucket','Notes'];
  const rows=[header];

  // Pull from driver tabs if they exist
  const growth = readKV_(getSheetByCandidates_(SHEET_CANDIDATES.GROWTH));
  const cost   = readKV_(getSheetByCandidates_(SHEET_CANDIDATES.COST));
  const haveDriverTabs = growth.length || cost.length;

  // If missing, seed from ASSUMPTIONS cost-related labels
  const fallbackLabels = [
    'Cost of Revenue % (COGS)','AI Agents cost per active customer per month (CHF)',
    'Media / PR monthly (CHF)','R&D AIX monthly (CHF)','Product dev (CaaS) monthly (CHF)',
    'Tech stack per FTE / month (CHF)','Culture & Learning per FTE / month (CHF)',
    'Labs & Universities monthly (CHF)',
    'Purchased Services – Micro monthly (CHF)','Purchased Services – Meso monthly (CHF)','Purchased Services – Macro monthly (CHF)','Purchased Services – Mundo monthly (CHF)'
  ];
  const fromAss = labelsFromAssumptions_(ass, fallbackLabels);

  const all = haveDriverTabs
    ? [...growth.map(([k])=>['Growth Hypothesis',k]), ...cost.map(([k])=>['cost hypothesis',k])]
    : fromAss.map(k=>['ASSUMPTIONS',k]);

  const defaultCatBucket = labelToDefaultCatBucket_;
  all.forEach(([src,label])=>{
    const {cat,bucket} = defaultCatBucket(label);
    rows.push([src,label,cat,bucket,'']);
  });

  mapSh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
  const cats=['Hosting/Infra','AI Runtime','Payment Fees','Marketing/Media','R&D / Product','Tech & Tools','Services & Labs','Culture & Learning','Other'];
  const buckets=['COGS','Opex'];
  mapSh.getRange(2,3,rows.length-1,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(cats,true).build());
  mapSh.getRange(2,4,rows.length-1,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(buckets,true).build());
  mapSh.setFrozenRows(1);

  appendNoteCell_(mapSh.getRange('A1'),'✍️ Edit Category/Bucket per line. This drives IS split across COGS & Opex buckets.');
  SpreadsheetApp.getActive().toast('MAPPING_CATEGORIES filled ✅','Colibri',5);
}
function readKV_(sheet){
  if (!sheet) return [];
  const lr=sheet.getLastRow(), lc=sheet.getLastColumn(); if (lr<2) return [];
  const hdrs=sheet.getRange(1,1,1,Math.min(lc,6)).getValues()[0].map(x=>String(x||'').toLowerCase());
  let colKey=1, colVal=2; const keyH=['label','metric','assumption','name','kpi']; const valH=['value','amount','chf','val','number'];
  for (let c=1;c<=Math.min(lc,6);c++){ const h=hdrs[c-1]; if (keyH.some(k=>h.indexOf(k)>-1)) colKey=c; if (valH.some(k=>h.indexOf(k)>-1)) colVal=c; }
  const data=sheet.getRange(2,1,lr-1,Math.max(colKey,colVal)).getValues();
  const out=[]; for (const row of data){ const k=String(row[colKey-1]||'').trim(); const v=row[colVal-1]; if (k) out.push([k,v]); }
  return out;
}
function labelsFromAssumptions_(ass, want){
  const last=ass.getLastRow(); if (last<2) return [];
  const labels=ass.getRange(2,2,last-1,1).getValues().map(r=>String(r[0]||'').trim());
  return labels.filter(l=> want.some(w=>l.indexOf(w)>-1) || /(monthly|per FTE|Purchased Services|COGS|AI Agents|Media|R&D|Product dev|Tech stack|Labs)/i.test(l));
}
function labelToDefaultCatBucket_(label){
  const l=String(label||'').toLowerCase();
  if (/(hosting|cloud|infra|server|cdn)/.test(l)) return {cat:'Hosting/Infra', bucket:'COGS'};
  if (/ai agent|ai runtime|ai cost|ai agents/.test(l)) return {cat:'AI Runtime', bucket:'COGS'};
  if (/payment|stripe|fees/.test(l)) return {cat:'Payment Fees', bucket:'COGS'};
  if (/media|marketing|ads|advert|pr/.test(l)) return {cat:'Marketing/Media', bucket:'Opex'};
  if (/r&d|research|aix|product dev/.test(l)) return {cat:'R&D / Product', bucket:'Opex'};
  if (/tech stack|tools|saas|license|software|cloud/.test(l)) return {cat:'Tech & Tools', bucket:'Opex'};
  if (/labs|universit/.test(l)) return {cat:'Services & Labs', bucket:'Opex'};
  if (/culture .*learning/.test(l)) return {cat:'Culture & Learning', bucket:'Opex'};
  if (/purchased services/.test(l)) return {cat:'Services & Labs', bucket:'Opex'};
  return {cat:'Other', bucket:'Opex'};
}

/* -------------------- 3) Build IS/CF/BS (reads mapping) -------------------- */
function buildFinancialStatements(){
  const model=getSheetByCandidates_(SHEET_CANDIDATES.MODEL);
  if (!model) throw new Error('MODEL tab not found.');

  // MODEL headers
  const headers=model.getRange(1,1,1,model.getLastColumn()).getValues()[0];
  const idx=n=>{ const i=headers.indexOf(n); if(i<0) throw new Error(`MODEL header not found: ${n}`); return i+1; };
  const cMonth=idx('Month'), cInHor=idx('In_Horizon'), cRev=idx('Revenue_Display'), cCosts=idx('Costs_Display'), cCash=idx('Cum_Cash');

  // Horizon rows
  const hv=model.getRange(2,cInHor,model.getLastRow()-1,1).getValues();
  let last=1; for (let i=0;i<hv.length;i++) if (hv[i][0]===1||hv[i][0]==='1') last=i+1;
  const rows=Math.max(last,1), lastCol=1+rows;
  const dates=model.getRange(2,cMonth,rows,1).getValues();

  // Mapping
  const mapSh=getOrInsert_('MAP','MAPPING_CATEGORIES');
  const mLR=mapSh.getLastRow();
  const mapRows=mLR>=2? mapSh.getRange(2,1,mLR-1,4).getValues():[];
  const catAgg = { 'Hosting/Infra':0,'AI Runtime':0,'Payment Fees':0,'Marketing/Media':0,'R&D / Product':0,'Tech & Tools':0,'Services & Labs':0,'Culture & Learning':0,'Other':0 };
  mapRows.forEach(r=>{
    const label=String(r[1]||'').trim(), cat=r[2]||'Other', bucket=r[3]||'Opex';
    // Pull amounts from ASSUMPTIONS if present
    const ass=getOrInsert_('ASSUMPTIONS','ASSUMPTIONS');
    const row=findLabelRow_(ass,label);
    const val=(row? ass.getRange(row,1).getValue() : 0) || 0;
    // Only aggregate numeric amounts (ignores % labels like COGS%)
    if (typeof val==='number') catAgg[cat]=(catAgg[cat]||0)+val;
  });

  /* ----- INCOME STATEMENT ----- */
  const is=getOrInsert_('IS','INCOME_STATEMENT'); is.clear();
  const hdr=['Metric', ...dates.map(d=>d[0])];
  const lines=[
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
  is.getRange(1,1,1,hdr.length).setValues([hdr]);
  for (let r=0;r<lines.length;r++) is.getRange(2+r,1).setValue(lines[r]);
  for (let c=2;c<=lastCol;c++){
    const REV = `INDEX(MODEL!${columnLetter_(cRev)}:${columnLetter_(cRev)}, MATCH(${is.getRange(1,c).getA1Notation()}, MODEL!${columnLetter_(cMonth)}:${columnLetter_(cMonth)}, 0))`;
    is.getRange(2,c).setFormula(`=${REV}`);
    // COGS buckets (flat monthly, mapping-driven)
    is.getRange(3,c).setValue(catAgg['Hosting/Infra']||0);
    is.getRange(4,c).setValue(catAgg['AI Runtime']||0);
    is.getRange(5,c).setValue(catAgg['Payment Fees']||0);
    is.getRange(6,c).setFormula(`=SUM(${is.getRange(3,c).getA1Notation()}:${is.getRange(5,c).getA1Notation()})`);
    // GP / GM%
    is.getRange(7,c).setFormula(`=${is.getRange(2,c).getA1Notation()}-${is.getRange(6,c).getA1Notation()}`);
    is.getRange(8,c).setFormula(`=IFERROR(${is.getRange(7,c).getA1Notation()}/${is.getRange(2,c).getA1Notation()},)`);
    // Opex buckets
    is.getRange(9,c).setValue(catAgg['Marketing/Media']||0);
    is.getRange(10,c).setValue(catAgg['R&D / Product']||0);
    is.getRange(11,c).setValue(catAgg['Tech & Tools']||0);
    is.getRange(12,c).setValue(catAgg['Services & Labs']||0);
    is.getRange(13,c).setValue(catAgg['Culture & Learning']||0);
    is.getRange(14,c).setValue(catAgg['Other']||0);
    is.getRange(15,c).setFormula(`=SUM(${is.getRange(9,c).getA1Notation()}:${is.getRange(14,c).getA1Notation()})`);
    // EBITDA / Dep / EBIT
    is.getRange(16,c).setFormula(`=${is.getRange(7,c).getA1Notation()}-${is.getRange(15,c).getA1Notation()}`);
    is.getRange(17,c).setFormula(`=IFERROR(INDEX(ASSUMPTIONS!A:A, MATCH("CapEx per month (CHF)", ASSUMPTIONS!B:B,0))/MAX(1, INDEX(ASSUMPTIONS!A:A, MATCH("Depreciation months", ASSUMPTIONS!B:B,0))), 0)`);
    is.getRange(18,c).setFormula(`=${is.getRange(16,c).getA1Notation()}-${is.getRange(17,c).getA1Notation()}`);
    // Interest (cash yield only, no debt schedule)
    const CF=getOrInsert_('CF','CASH_FLOW');
    const CASH_t=CF.getRange(10,c).getA1Notation();
    const CASH_prev=c>2? CF.getRange(10,c-1).getA1Notation() : CASH_t;
    const AVG_CASH=`(${CASH_t}+${CASH_prev})/2`;
    const CASH_INT=`IFERROR(${AVG_CASH} * INDEX(ASSUMPTIONS!A:A, MATCH("Cash yield %", ASSUMPTIONS!B:B,0))/100/12, 0)`;
    is.getRange(19,c).setFormula(`=${CASH_INT}`);
    is.getRange(20,c).setFormula(`=${is.getRange(18,c).getA1Notation()}+${is.getRange(19,c).getA1Notation()}`);
    is.getRange(21,c).setFormula(`=MAX(0, ${is.getRange(20,c).getA1Notation()} * INDEX(ASSUMPTIONS!A:A, MATCH("Tax rate %", ASSUMPTIONS!B:B,0))/100)`);
    is.getRange(22,c).setFormula(`=${is.getRange(20,c).getA1Notation()}-${is.getRange(21,c).getA1Notation()}`);
  }

  /* ----- CASH FLOW ----- */
  const cf=getOrInsert_('CF','CASH_FLOW'); cf.clear();
  const cfHdr=['Metric', ...dates.map(d=>d[0])];
  const cfLines=['Net Income 🧾','Non-cash: Depreciation 🧱','Working Capital Δ 🔄','Operating Cash Flow 💧','CapEx 🛠️','Free Cash Flow 🟢','Financing (Debt/Equity) 💳','Net Cash Change 💱','Ending Cash 🏦'];
  cf.getRange(1,1,1,cfHdr.length).setValues([cfHdr]);
  cfLines.forEach((L,i)=>cf.getRange(2+i,1).setValue(L));
  for (let c=2;c<=lastCol;c++){
    cf.getRange(2,c).setFormula(`=${is.getRange(22,c).getA1Notation()}`);
    cf.getRange(3,c).setFormula(`=${is.getRange(17,c).getA1Notation()}`);
    // keep WC simple (0) unless you enable DSO/DPO/DIO later
    cf.getRange(4,c).setValue(0);
    cf.getRange(5,c).setFormula(`=${cf.getRange(2,c).getA1Notation()}+${cf.getRange(3,c).getA1Notation()}+${cf.getRange(4,c).getA1Notation()}`);
    cf.getRange(6,c).setFormula(`=IFERROR(INDEX(ASSUMPTIONS!A:A, MATCH("CapEx per month (CHF)", ASSUMPTIONS!B:B,0)),0)`);
    cf.getRange(7,c).setFormula(`=${cf.getRange(5,c).getA1Notation()}-${cf.getRange(6,c).getA1Notation()}`);
    cf.getRange(8,c).setValue(0);
    cf.getRange(9,c).setFormula(`=${cf.getRange(7,c).getA1Notation()}+${cf.getRange(8,c).getA1Notation()}`);
    const CASH_MODEL=`INDEX(MODEL!${columnLetter_(cCash)}:${columnLetter_(cCash)}, MATCH(${cf.getRange(1,c).getA1Notation()}, MODEL!${columnLetter_(cMonth)}:${columnLetter_(cMonth)},0))`;
    cf.getRange(10,c).setFormula(`=${CASH_MODEL}`);
  }

  /* ----- BALANCE SHEET ----- */
  const bs=getOrInsert_('BS','BALANCE_SHEET'); bs.clear();
  const bsHdr=['Metric', ...dates.map(d=>d[0])];
  const bsLines=['Cash 🏦','A/R 📬','Inventory 📦','Prepaids & Other 🗃️','PP&E (net) 🏭','Total Assets 💼','A/P 🧾','Other Liab 📑','Deferred Rev ⏳','Debt 💳','Total Liabilities 🧮','Equity 📈','Liabilities + Equity ⚖️','Balance Check ✅'];
  bs.getRange(1,1,1,bsHdr.length).setValues([bsHdr]);
  bsLines.forEach((L,i)=>bs.getRange(2+i,1).setValue(L));
  for (let c=2;c<=lastCol;c++){
    bs.getRange(2,c).setFormula(`=${cf.getRange(10,c).getA1Notation()}`);
    bs.getRange(3,c).setValue(0);
    bs.getRange(4,c).setValue(0);
    const prev=c>2? bs.getRange(6,c-1).getA1Notation() : '0';
    bs.getRange(5,c).setFormula(`=${prev}+${cf.getRange(6,c).getA1Notation()}-${cf.getRange(3,c).getA1Notation()}`);
    bs.getRange(6,c).setFormula(`=SUM(${bs.getRange(2,c).getA1Notation()}:${bs.getRange(5,c).getA1Notation()})`);
    bs.getRange(7,c).setValue(0);
    bs.getRange(8,c).setValue(0);
    bs.getRange(9,c).setValue(0);
    bs.getRange(10,c).setValue(0);
    bs.getRange(11,c).setFormula(`=SUM(${bs.getRange(7,c).getA1Notation()}:${bs.getRange(10,c).getA1Notation()})`);
    bs.getRange(12,c).setFormula(`=${bs.getRange(6,c).getA1Notation()}-${bs.getRange(11,c).getA1Notation()}`);
    bs.getRange(13,c).setFormula(`=${bs.getRange(11,c).getA1Notation()}+${bs.getRange(12,c).getA1Notation()}`);
    bs.getRange(14,c).setFormula(`=${bs.getRange(6,c).getA1Notation()}-${bs.getRange(13,c).getA1Notation()}`);
  }

  SpreadsheetApp.getActive().toast('IS/CF/BS rebuilt ✅','Colibri',5);
}

/* -------------------- 4) Long notes (titles + all editable values) -------------------- */
function appendLongNotesEverywhere(){
  // ASSUMPTIONS — headers
  const ass=getOrInsert_('ASSUMPTIONS','ASSUMPTIONS');
  appendNoteCell_(ass.getRange('A1'),'🔢 Enter numbers/dates. This is the only column you usually edit.');
  appendNoteCell_(ass.getRange('B1'),'🏷️ Labels. Formulas look up these exact names.');
  appendNoteCell_(ass.getRange('C1'),'🔗 Sources / rationale. Paste links here.');
  appendNoteCell_(ass.getRange('D1'),'📚 Section for scannability (matches your flowchart).');

  // ASSUMPTIONS — per-value notes (full 6-part). Known labels map below; others get a generic helper.
  const notes = assumptionNotesCH_();
  const last=ass.getLastRow(); if (last>=2){
    const labR=ass.getRange(2,2,last-1,1).getValues(), valR=ass.getRange(2,1,last-1,1);
    const valNotes=valR.getNotes();
    for (let r=0;r<labR.length;r++){
      const lab=String(labR[r][0]||'').trim();
      const n = notes[lab] || notes.__generic(lab);
      const tag = `— Colibri: ${blockNote_(n)}`;
      const cur = valNotes[r][0]||'';
      if (cur.indexOf(tag)===-1) valNotes[r][0]=cur?`${cur}\n\n${tag}`:tag;
    }
    valR.setNotes(valNotes);
  }

  // MODEL / SUMMARY / IS / CF / BS — header title notes
  headerTitleNotes_();

  // Formula library short helper notes
  const fl=getOrInsert_('FORMULA_LIB','FORMULA_LIBRARY');
  appendNoteCell_(fl.getRange('A1'),'📚 Handy functions used in this model.');
  SpreadsheetApp.getActive().toast('Long notes appended across tabs ✅','Colibri',5);
}
function assumptionNotesCH_(){
  const A = (t)=>`INDEX(ASSUMPTIONS!A:A, MATCH("${t}", ASSUMPTIONS!B:B, 0))`;
  return {
    'Starting Cash (CHF)': {
      what:'Money in bank at the start month.',
      source:'Cash Flow (begin) & Balance Sheet (Cash).',
      how:'Used to seed MODEL Cum_Cash;\nRunway = Starting Cash ÷ Latest Net Burn.',
      typical:'CHF 50k–500k seed stage (varies).',
      when:'Change when you raise/spend before start.',
      warn:'0 with positive burn → runway 0.'
    },
    'Payroll Overhead %': {
      what:'Employer on-top costs (AHV/ALV/BVG/accident).',
      source:'Income Statement (Opex).',
      how:'Payroll/mo = (Σ HC×Salary ÷12)×(1+Overhead%).',
      typical:'~12–22% depending on benefits.',
      when:'Adjust per benefits set.',
      warn:'Too low inflates profits.'
    },
    'Salary L1 (CHF/yr)': { what:'Yearly salary Level 1 (Exec/Architect).', source:'Opex (Payroll).', how:'Part of Σ HC×Salary ÷12.', typical:'CHF 140k–220k+', when:'Adjust to actual contracts.', warn:'Outliers skew payroll.' },
    'Salary L2 (CHF/yr)': { what:'Yearly salary Level 2 (Lead/Owner).', source:'Opex (Payroll).', how:'Part of Σ HC×Salary ÷12.', typical:'CHF 110k–160k', when:'Adjust to actual.', warn:'—' },
    'Salary L3 (CHF/yr)': { what:'Yearly salary Level 3 (Senior).', source:'Opex (Payroll).', how:'Part of Σ HC×Salary ÷12.', typical:'CHF 90k–130k', when:'Adjust to actual.', warn:'—' },
    'Salary L4 (CHF/yr)': { what:'Yearly salary Level 4 (Associate).', source:'Opex (Payroll).', how:'Part of Σ HC×Salary ÷12.', typical:'CHF 60k–95k', when:'Adjust to actual.', warn:'—' },
    'Starting Headcount': { what:'People on payroll at start.', source:'Opex (Payroll).', how:'Drives Σ HC in MODEL.', typical:'Founders 2–4 + early hires.', when:'Update as you hire.', warn:'—' },
    'Target Headcount (by year 3)': { what:'Goal headcount ~36 months.', source:'Opex (Payroll).', how:'MODEL ramps towards this.', typical:'Context-specific.', when:'Update as plans evolve.', warn:'—' },
    'Cost of Revenue % (COGS)': {
      what:'Direct costs share of revenue.',
      source:'Income Statement (COGS).',
      how:`COGS = Revenue × ${A('Cost of Revenue % (COGS)')}/100.`,
      typical:'SaaS net GM 60–90%.',
      when:'Update with infra/support data.',
      warn:'0% likely unrealistic.'
    },
    'AI Agents cost per active customer per month (CHF)': {
      what:'AI runtime CHF per active customer.',
      source:'COGS (direct).',
      how:'COGS += Active_Customers × AI_cost_per_customer.',
      typical:'Highly variable (usage).',
      when:'Update with observed usage.',
      warn:'If 0, GM% can look too high.'
    },
    'ARPU – CaaS monthly (CHF)': {
      what:'Average revenue per customer per month.',
      source:'Income Statement (Revenue).',
      how:'MRR = Active_Customers × ARPU.',
      typical:'CHF 100–600 B2B (varies).',
      when:'Change when repricing.',
      warn:'0 → zero MRR.'
    },
    'Churn monthly (0–1)': {
      what:'Monthly % of customers that cancel.',
      source:'Revenue dynamics.',
      how:'Active_t = Active_{t-1}×(1−churn)+New.',
      typical:'1–5% B2B (enterprise <1%).',
      when:'Update with real data.',
      warn:'High churn kills LTV.'
    },
    'Awareness→Conversion (0–1)': {
      what:'% of leads that become paying customers.',
      source:'Top-of-funnel.',
      how:'New = Leads × Conversion.',
      typical:'1–5% cold; higher warm.',
      when:'Update from funnel metrics.',
      warn:'0 → no growth.'
    },
    'CAC blended (CHF)': {
      what:'Average cost to acquire one customer.',
      source:'Unit economics.',
      how:'CAC Payback = CAC ÷ (ARPU × GM%).',
      typical:'CHF 1k–8k+ (B2B).',
      when:'Update from real campaigns.',
      warn:'Huge CAC + low ARPU is bad.'
    },
    'LTV months': {
      what:'Typical months a customer stays.',
      source:'Unit economics.',
      how:'LTV = ARPU × months × GM%.',
      typical:'12–60 based on segment.',
      when:'Update as retention matures.',
      warn:'Too low harms LTV/CAC.'
    },
    'Gross Margin %': {
      what:'% kept after direct costs.',
      source:'P&L.',
      how:'GM% = GrossProfit ÷ Revenue.',
      typical:'SaaS 60–90%.',
      when:'Derived; no manual change.',
      warn:'If forcing manual, beware.'
    },
    'AIX Monthly price (CHF)': {
      what:'Price per month for AIX content.',
      source:'Revenue.',
      how:'AIX_Monthly = subs × price.',
      typical:'Context-specific.',
      when:'Change when repricing.',
      warn:'—'
    },
    'AIX Yearly price (CHF)': {
      what:'Yearly price (recognized monthly /12).',
      source:'Revenue + Deferred Rev.',
      how:'Monthly recog = yearly/12.',
      typical:'Context-specific.',
      when:'Change when repricing.',
      warn:'—'
    },
    'Initial CaaS customers': {
      what:'Customers at month 1.',
      source:'Revenue dynamics.',
      how:'Seed Active_Customers.',
      typical:'0–20 early stage.',
      when:'Set actual count.',
      warn:'—'
    },
    'Monthly leads': {
      what:'Leads entering funnel per month.',
      source:'Funnel.',
      how:'New customers = leads × conversion.',
      typical:'Depends on channel.',
      when:'Update with marketing plan.',
      warn:'—'
    },
    'Day rate – training (CHF)': {
      what:'Training/consulting fee per day.',
      source:'Revenue (services).',
      how:'Revenue = rate × days.',
      typical:'CHF 1.5k–5k+.',
      when:'Update your offer.',
      warn:'—'
    },
    'Training days / month': {
      what:'Billable training/consulting days.',
      source:'Revenue (services).',
      how:'Revenue = rate × days.',
      typical:'0–10 early stage.',
      when:'Update capacity.',
      warn:'—'
    },
    'Other digital products – start (CHF)': {
      what:'Starting revenue for experiments.',
      source:'Revenue (other).',
      how:'Growth compounding per month.',
      typical:'Small pilot values.',
      when:'Set when testing ideas.',
      warn:'—'
    },
    'Other digital monthly growth (0–1)': {
      what:'Monthly growth rate for other digital.',
      source:'Revenue (other).',
      how:'Value_t = start×(1+g)^t.',
      typical:'0–10%/mo.',
      when:'Tune to learning pace.',
      warn:'>20%/mo may be optimistic.'
    },
    'Culture & Learning per FTE / month (CHF)': {
      what:'Monthly budget per employee for learning.',
      source:'Opex.',
      how:'= Headcount × budget/FTE.',
      typical:'CHF 50–200.',
      when:'Change with policy.',
      warn:'—'
    },
    'Tech stack per FTE / month (CHF)': {
      what:'SaaS tools/cloud per employee.',
      source:'Opex.',
      how:'= Headcount × tools/FTE.',
      typical:'CHF 50–300.',
      when:'Change with stack.',
      warn:'—'
    },
    'Labs & Universities monthly (CHF)': {
      what:'Collaboration budget (Aalto etc.).',
      source:'Opex.',
      how:'Flat monthly.',
      typical:'CHF 500–5k+.',
      when:'Change with contracts.',
      warn:'—'
    },
    'R&D AIX monthly (CHF)': {
      what:'AIX R&D budget.',
      source:'Opex.',
      how:'Flat monthly.',
      typical:'CHF 1k–10k+.',
      when:'Change by roadmap.',
      warn:'—'
    },
    'Product dev (CaaS) monthly (CHF)': {
      what:'CaaS product build budget.',
      source:'Opex.',
      how:'Flat monthly.',
      typical:'CHF 2k–20k+.',
      when:'Change by roadmap.',
      warn:'—'
    },
    'Media / PR monthly (CHF)': {
      what:'Paid marketing/PR.',
      source:'Opex.',
      how:'Flat monthly.',
      typical:'CHF 500–20k+.',
      when:'Change by plan.',
      warn:'—'
    },
    'Purchased Services – Micro monthly (CHF)': {
      what:'Coaching/individual services.',
      source:'Opex.',
      how:'Flat monthly.',
      typical:'Varies.',
      when:'Change by contracts.',
      warn:'—'
    },
    'Purchased Services – Meso monthly (CHF)': {
      what:'Team-level external services.',
      source:'Opex.',
      how:'Flat monthly.',
      typical:'Varies.',
      when:'Change by contracts.',
      warn:'—'
    },
    'Purchased Services – Macro monthly (CHF)': {
      what:'Org-level consulting.',
      source:'Opex.',
      how:'Flat monthly.',
      typical:'Varies.',
      when:'Change by contracts.',
      warn:'—'
    },
    'Purchased Services – Mundo monthly (CHF)': {
      what:'Specialist expertise (e.g., MLOps).',
      source:'Opex.',
      how:'Flat monthly.',
      typical:'Varies.',
      when:'Change by contracts.',
      warn:'—'
    },
    // Programmatic fallback:
    __generic: (lab)=>({
      what:`Input for ${lab}.`,
      source:'Feeds MODEL/IS as appropriate.',
      how:'Referenced via INDEX/MATCH by label.',
      typical:'Set to your reality.',
      when:'Edit as your plan updates.',
      warn:'Blank may propagate zeros.'
    })
  };
}
function headerTitleNotes_(){
  const addHead = (sh, map)=>{
    if (!sh) return;
    const lc=sh.getLastColumn(); const hdr=sh.getRange(1,1,1,lc);
    const hvals=hdr.getValues()[0]; const notes=hdr.getNotes();
    for (let i=0;i<hvals.length;i++){
      const k=String(hvals[i]||'').trim();
      const t=map[k]; if (!t) continue;
      const tag=`— Colibri: ${t}`; const cur=notes[0][i]||'';
      if (cur.indexOf(tag)===-1) notes[0][i]=cur?`${cur}\n\n${tag}`:tag;
    }
    hdr.setNotes(notes);
  };
  addHead(getSheetByCandidates_(SHEET_CANDIDATES.MODEL), {
    'Month':'📅 First of each month (timeline).',
    'Revenue_Display':'💵 Total revenue (for charts).',
    'Costs_Display':'💸 Total costs (for charts).',
    'Net_Burn':'🔥 Costs − Revenue (positive = burn).',
    'Cum_Cash':'🏦 Running cash balance.'
  });
  addHead(getOrInsert_('SUMMARY','SUMMARY'), {
    'Metric':'🏷️ KPI name',
    'Value':'🔢 Value (CHF)',
    'Description':'ℹ️ Explanation'
  });
  addHead(getOrInsert_('IS','INCOME_STATEMENT'), {'Metric':'🏷️ P&L line item'});
  addHead(getOrInsert_('CF','CASH_FLOW'), {'Metric':'🏷️ Cash flow line item'});
  addHead(getOrInsert_('BS','BALANCE_SHEET'), {'Metric':'🏷️ Balance sheet line item'});
}

/* -------------------- 5) MODEL blue color scales (safe) -------------------- */
function applyModelColorScales(){
  const sh=getSheetByCandidates_(SHEET_CANDIDATES.MODEL);
  if (!sh){ SpreadsheetApp.getActive().toast('MODEL not found','Colibri',5); return; }
  const lr=sh.getLastRow(), lc=sh.getLastColumn();
  const headers=sh.getRange(1,1,1,lc).getValues()[0];
  const targets=['MRR_Display','ARR_Display','Revenue_Display','Costs_Display','Net_Burn','Active_Customers','Headcount_Total']
    .map(h=>headers.indexOf(h)+1).filter(i=>i>0);

  // Keep all existing rules, just ADD our per-column gradients (no deletes)
  const rules=sh.getConditionalFormatRules();
  targets.forEach(col=>{
    const rule=SpreadsheetApp.newConditionalFormatRule()
      .setRanges([sh.getRange(2,col,Math.max(lr-1,1),1)])
      .setGradientMinpoint('#E3F2FD')
      .setGradientMaxpoint('#0D47A1')
      .build();
    rules.push(rule);
  });
  sh.setConditionalFormatRules(rules);
  SpreadsheetApp.getActive().toast('Blue gradients applied (MODEL) ✅','Colibri',5);
}

/* -------------------- helpers -------------------- */
function diagnostics(){
  const ss=SpreadsheetApp.getActive();
  const names=ss.getSheets().map(s=>s.getName()).join(', ');
  const found=key=>!!getSheetByCandidates_(SHEET_CANDIDATES[key]);
  SpreadsheetApp.getUi().alert([
    `File: ${ss.getName()}`,
    `Sheets: ${names}`,
    `MODEL: ${found('MODEL')}`,
    `ASSUMPTIONS: ${found('ASSUMPTIONS')}`,
    `MAPPING_CATEGORIES: ${found('MAP')}`,
    `IS/CF/BS: ${found('IS')}/${found('CF')}/${found('BS')}`,
    `Growth/Cost tabs present: ${found('GROWTH')}/${found('COST')}`
  ].join('\n'));
}
