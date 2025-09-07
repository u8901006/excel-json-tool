/*********************
 * 共同：選單聚合器
 *********************/
function onOpen() {
  buildOrderMenu_();       // 叫藥自動化
  buildInventoryMenu_();   // 庫存工具（條件格式/批次上色/觸發器）
  buildDoubleStockMenu_(); // 2倍庫存（由指定工作表產生【結果】）
}

/* ---------- 叫藥自動化 ---------- */
function buildOrderMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('叫藥自動化')
    .addItem('生成訂單文字', 'generateVendorOrderLines')
    .addSeparator()
    .addItem('從本機 JSON 填入耗量（完全一致）', 'WF_showUploadDialog') // ⬅ 新增
    .addToUi();
}

/* ---------- 庫存工具 ---------- */
function buildInventoryMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('庫存工具')
    .addItem('套用藍色警示（1倍：結餘 < 安全庫存）', 'applyConditionalBlueForConfiguredSheet')
    .addItem('套用粉紅警示（2倍：M ≤ 結餘 < 2M）', 'applyConditionalPinkForConfiguredSheet')
    .addSeparator()
    .addItem('清除此表的藍色「條件格式」', 'clearConditionalBlueForConfiguredSheet')
    .addItem('清除此表的粉紅「條件格式」', 'clearConditionalPinkForConfiguredSheet')
    .addSeparator()
    .addItem('重新掃描並填色（靜態：2倍）', 'scanAndPaintPinkForConfiguredSheet')
    .addSeparator()
    .addItem('建立每天 07:30 觸發器', 'createDailyTrigger0730')
    .addItem('刪除本專案所有觸發器', 'deleteAllTriggers')
    .addToUi();
}

/* ---------- 2倍庫存 ---------- */
function buildDoubleStockMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('2倍庫存')
    .addItem('從指定工作表產生【結果】（依整列底色）', 'runDoubleSafetyFromSheetPrompt')
    .addToUi();
}

/********************************************
 * 庫存工具（結餘 vs 安全庫存）
 ********************************************/
const CONFIG = {
  spreadsheetId: '',
  sheetName: null,          // 若指定名稱就用名稱；否則用 sheetIndex1Based
  sheetIndex1Based: 8,      // 預設第 8 張工作表
  headerRow: 1,
  colBalance: 'L',
  colSafety: 'M',
  pink: '#FFC0CB',          // 2倍：M ≤ L < 2M
  blue: '#CFE2F3'           // 1倍：L < M
};

function getSpreadsheet_() {
  return CONFIG.spreadsheetId
    ? SpreadsheetApp.openById(CONFIG.spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();
}
function getTargetSheet_() {
  const ss = getSpreadsheet_();
  if (CONFIG.sheetName && String(CONFIG.sheetName).trim() !== '') {
    const sh = ss.getSheetByName(String(CONFIG.sheetName));
    if (!sh) throw new Error('找不到名為『' + CONFIG.sheetName + '』的工作表');
    return sh;
  }
  const idx = Math.max(1, Number(CONFIG.sheetIndex1Based) || 1) - 1;
  const sheets = ss.getSheets();
  if (idx < 0 || idx >= sheets.length) throw new Error('第 ' + CONFIG.sheetIndex1Based + ' 張工作表不存在');
  return sheets[idx];
}
function colA1ToIndex_(a1) {
  let n = 0, s = a1.toUpperCase().replace(/[^A-Z]/g, '');
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

/** ✅ 1倍：結餘 < 安全庫存 → 藍色 */
function applyConditionalBlueForConfiguredSheet() {
  const sheet = getTargetSheet_();
  const headerRow = CONFIG.headerRow || 1;
  const startRow = headerRow + 1;

  const lastRow = Math.max(startRow, sheet.getLastRow());
  const lastCol = sheet.getLastColumn();
  if (lastRow < startRow || lastCol < 1) return;

  const range = sheet.getRange(startRow, 1, lastRow - headerRow, lastCol);

  const L = CONFIG.colBalance, M = CONFIG.colSafety;
  const formulaBlue = `=$${L}${startRow}<$${M}${startRow}`; // L < M

  const blueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formulaBlue)
    .setBackground(CONFIG.blue)
    .setRanges([range])
    .build();

  const rules = sheet.getConditionalFormatRules();
  const kept = rules.filter(r => {
    const bc = r.getBooleanCondition && r.getBooleanCondition();
    if (!bc) return true;
    const type = bc.getCriteriaType && bc.getCriteriaType();
    const vals = bc.getCriteriaValues && bc.getCriteriaValues();
    return !(type === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA &&
             vals && vals[0] === formulaBlue);
  });
  kept.push(blueRule);
  sheet.setConditionalFormatRules(kept);
}

/** ✅ 2倍：M ≤ 結餘 < 2M → 粉紅（與藍色互斥） */
function applyConditionalPinkForConfiguredSheet() {
  const sheet = getTargetSheet_();
  const headerRow = CONFIG.headerRow || 1;
  const startRow = headerRow + 1;

  const lastRow = Math.max(startRow, sheet.getLastRow());
  const lastCol = sheet.getLastColumn();
  if (lastRow < startRow || lastCol < 1) return;

  const range = sheet.getRange(startRow, 1, lastRow - headerRow, lastCol);

  const L = CONFIG.colBalance, M = CONFIG.colSafety;
  const formulaPink = `=AND($${L}${startRow}>=${'$'}${M}${startRow},$${L}${startRow}<2*$${M}${startRow})`;

  const pinkRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formulaPink)
    .setBackground(CONFIG.pink)
    .setRanges([range])
    .build();

  const rules = sheet.getConditionalFormatRules();
  const kept = rules.filter(r => {
    const bc = r.getBooleanCondition && r.getBooleanCondition();
    if (!bc) return true;
    const type = bc.getCriteriaType && bc.getCriteriaType();
    const vals = bc.getCriteriaValues && bc.getCriteriaValues();
    return !(type === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA &&
             vals && vals[0] === formulaPink);
  });
  kept.push(pinkRule);
  sheet.setConditionalFormatRules(kept);
}

/** ✅ 清除此表的藍色條件格式（只移除我們那條） */
function clearConditionalBlueForConfiguredSheet() {
  const sheet = getTargetSheet_();
  const headerRow = CONFIG.headerRow || 1;
  const startRow = headerRow + 1;
  const L = CONFIG.colBalance, M = CONFIG.colSafety;
  const formulaBlue = `=$${L}${startRow}<$${M}${startRow}`;

  const rules = sheet.getConditionalFormatRules();
  const kept = rules.filter(r => {
    const bc = r.getBooleanCondition && r.getBooleanCondition();
    if (!bc) return true;
    const type = bc.getCriteriaType && bc.getCriteriaType();
    const vals = bc.getCriteriaValues && bc.getCriteriaValues();
    const isOur = (type === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA &&
                   vals && vals[0] === formulaBlue);
    return !isOur;
  });
  sheet.setConditionalFormatRules(kept);
}

/** ✅ 清除此表的粉紅條件格式（只移除我們那條） */
function clearConditionalPinkForConfiguredSheet() {
  const sheet = getTargetSheet_();
  const headerRow = CONFIG.headerRow || 1;
  const startRow = headerRow + 1;
  const L = CONFIG.colBalance, M = CONFIG.colSafety;
  const formulaPink = `=AND($${L}${startRow}>=${'$'}${M}${startRow},$${L}${startRow}<2*$${M}${startRow})`;

  const rules = sheet.getConditionalFormatRules();
  const kept = rules.filter(r => {
    const bc = r.getBooleanCondition && r.getBooleanCondition();
    if (!bc) return true;
    const type = bc.getCriteriaType && bc.getCriteriaType();
    const vals = bc.getCriteriaValues && bc.getCriteriaValues();
    const isOur = (type === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA &&
                   vals && vals[0] === formulaPink);
    return !isOur;
  });
  sheet.setConditionalFormatRules(kept);
}

/** （可選）一次批次把符合 2 倍條件的列塗粉紅（靜態，不會自動更新） */
function scanAndPaintPinkForConfiguredSheet() {
  const sheet = getTargetSheet_();
  const headerRow = CONFIG.headerRow || 1;
  const startRow = headerRow + 1;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < startRow) return;
  const colL = colA1ToIndex_(CONFIG.colBalance);
  const colM = colA1ToIndex_(CONFIG.colSafety);
  const valuesL = sheet.getRange(startRow, colL, lastRow - headerRow, 1).getValues();
  const valuesM = sheet.getRange(startRow, colM, lastRow - headerRow, 1).getValues();
  const bgs = sheet.getRange(startRow, 1, lastRow - headerRow, lastCol).getBackgrounds();
  for (let i = 0; i < valuesL.length; i++) {
    const L = parseFloat(valuesL[i][0]);
    const M = parseFloat(valuesM[i][0]);
    const cond = (isFinite(L) && isFinite(M)) ? (L >= M && L < 2 * M) : false;
    if (cond) for (let c = 0; c < lastCol; c++) bgs[i][c] = CONFIG.pink;
  }
  sheet.getRange(startRow, 1, lastRow - headerRow, lastCol).setBackgrounds(bgs);
}

function createDailyTrigger0730() {
  deleteAllTriggers();
  ScriptApp.newTrigger('applyConditionalPinkForConfiguredSheet')
    .timeBased().atHour(7).nearMinute(30).everyDays(1).create();
}
function deleteAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}

/********************************************
 * 2倍庫存（由指定工作表 → 彙整到【結果】）
 * 依整列底色過濾；B=藥名, L=結餘, M=安全庫存, P=廠商
 ********************************************/
const DOUBLE_STOCK_CONFIG = {
  headerRow: 1,
  targetColors: ['#ead1dc', '#ffc0cb'].map(s => s.toLowerCase()), // 淺洋紅色3 & 粉紅
  colIdx: { // 0-based
    name: 1,      // B
    balance: 11,  // L
    safety: 12,   // M
    vendor: 15    // P
  },
  outputSheet: '結果'
};

function runDoubleSafetyFromSheetPrompt() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('2倍庫存', '請輸入要處理的工作表名稱（例如：8）', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const sheetName = String(resp.getResponseText() || '').trim();
  if (!sheetName) { ui.alert('未輸入工作表名稱'); return; }

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) { ui.alert('找不到工作表：' + sheetName); return; }

  const results = buildDoubleSafetyResultsFromSheet_(sheet, DOUBLE_STOCK_CONFIG);
  writeDoubleSafetyResults_(ss, results, DOUBLE_STOCK_CONFIG.outputSheet);
  ui.alert('完成：已輸出到【' + DOUBLE_STOCK_CONFIG.outputSheet + '】');
}

function buildDoubleSafetyResultsFromSheet_(sheet, conf) {
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const bgColors = dataRange.getBackgrounds().map(row => row.map(c => String(c || '').toLowerCase()));

  const startRow = (conf.headerRow || 1);
  const vendorGroups = {};

  for (let r = startRow; r < data.length; r++) {
    const rowColors = bgColors[r];
    const hasTarget = rowColors && rowColors.some(c => conf.targetColors.includes(c));
    if (!hasTarget) continue;

    const row = data[r];
    const drugName = row[conf.colIdx.name];
    const balance  = parseFloat(row[conf.colIdx.balance]);
    const safety   = parseFloat(row[conf.colIdx.safety]);
    const vendor   = row[conf.colIdx.vendor];

    if (!isFinite(balance) || !isFinite(safety)) {
      Logger.log(`第 ${r + 1} 行的安全庫存或目前結餘資料有誤`);
      continue;
    }
    const doubleSafe = safety * 2;
    const entry = `${drugName} (兩倍安全庫存: ${doubleSafe})`;

    const key = String(vendor || '').trim() || '（未指定廠商）';
    (vendorGroups[key] = vendorGroups[key] || []).push(entry);
  }

  const results = [['廠商', '藥品資訊']];
  Object.keys(vendorGroups).sort((a,b)=>a.localeCompare(b,'zh-Hant'))
    .forEach(vendor => {
      results.push([vendor, vendorGroups[vendor].join('\n')]);
    });
  return results;
}

function writeDoubleSafetyResults_(ss, results, sheetName) {
  const outSheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  outSheet.clear();
  outSheet.getRange(1, 1, results.length, results[0].length).setValues(results);
  outSheet.getRange(1, 2, results.length, 1).setWrap(true);
}

/********************************************
 * 叫藥自動化（讀【結果】→ 寫【訂單文字】）
 ********************************************/
const RESULT_SHEET = '結果';
const MAP_SHEET = '常見量對照';
const OUTPUT_SHEET = '訂單文字';

const INCLUDE_SPEC = false;
const INCLUDE_SPEC_SMART = true;
const SHOW_IF_QTY_MISSING = true;
const QTY_PLACEHOLDER = '（未設常見量）';
const DEFAULT_UNIT = '盒';
const FUZZY_VENDOR = true;

const BRAND_PREFIX_RE = /^(apo|sandoz|teva|mylan|stada|actavis|sun|krka|aurobindo|cipla|gsk|pfizer|roche|merck|bayer|abbvie|janssen|novartis|sanofi|lilly|astrazeneca|msd|takeda|kyowa|otsuka|macleods)[\s\-_\.]*/i;

function generateVendorOrderLines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheet = ss.getSheetByName(RESULT_SHEET);
  const mapSheet = ss.getSheetByName(MAP_SHEET);
  if (!resultSheet) throw new Error(`找不到分頁：${RESULT_SHEET}`);
  if (!mapSheet) throw new Error(`找不到分頁：${MAP_SHEET}（請建立並放欄位：商品｜規格｜常見叫藥數量｜廠商）`);
  const resultRows = readResultSheet_(resultSheet);
  const mapRows = readMapSheet_(mapSheet);
  const vendorText = buildVendorBigText_(resultRows);
  const lines = makeLinesFromMap_(vendorText, mapRows);
  writeOutput_(ss, lines);
}

function readResultSheet_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];
  const headers = values[0].map(h => String(h).trim());
  const idx = colIndexFinder_(headers, {
    vendor: ['廠商','供應商','製造商','廠牌'],
    name: ['商品名','商品名稱','品名','藥品名稱','藥品名','名稱','品項'],
    spec: ['規格','含量','劑量','規格含量','包裝','Strength'],
    info: ['藥品資訊','藥品資訊(商品+規格)','資訊']
  });
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const vendor = normText_(pickCell_(row, headers, idx.vendor));
    const name   = normText_(pickCell_(row, headers, idx.name));
    const spec   = normText_(pickCell_(row, headers, idx.spec));
    const info   = normText_(pickCell_(row, headers, idx.info));
    const rawText = [name, spec, info].filter(Boolean).join(' ').trim();
    if (!vendor && !rawText) continue;
    rows.push({ vendor, name, spec, info, rawText });
  }
  return rows;
}

function readMapSheet_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error(`對照分頁 ${MAP_SHEET} 是空的`);
  const headers = values[0].map(h => String(h).trim());
  ['商品','常見叫藥數量'].forEach(k => { if (!headers.includes(k)) throw new Error(`${MAP_SHEET} 缺少欄位：${k}`); });
  const get = name => headers.indexOf(name);
  const iName = get('商品');
  const iSpec = headers.indexOf('規格');
  const iQty  = get('常見叫藥數量');
  const iVen  = headers.indexOf('廠商');
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const v = values[r];
    const name = normText_(v[iName]);
    const spec = normText_(iSpec >= 0 ? v[iSpec] : '');
    const qty  = normText_(v[iQty]);
    const vendor = normText_(iVen >= 0 ? v[iVen] : '');
    if (!name) continue;

    const nameKey = normKey_(name);
    const nameNoBrand = dropBrand_(name);
    const nameNoBrandKey = normKey_(nameNoBrand);
    const specMinimal = specMinimalKey_(spec);
    const nameSpecKey = normKey_(name + ' ' + (spec || ''));
    const nameSpecNoBrandKey = normKey_(nameNoBrand + ' ' + (spec || ''));

    rows.push({
      name, nameNorm: name.toLowerCase(),
      spec, specNorm: (spec || '').toLowerCase(),
      qty,
      vendor, vendorNorm: (vendor || '').toLowerCase(),
      nameKey, nameNoBrandKey, nameSpecKey, nameSpecNoBrandKey, specMinimal
    });
  }
  return rows;
}

function buildVendorBigText_(rows) {
  const map = {};
  rows.forEach(({ vendor, rawText }) => {
    const key = vendor || '';
    map[key] = (map[key] || '') + ' ' + (rawText || '');
  });
  Object.keys(map).forEach(k => { map[k] = normKey_(normText_(map[k]).toLowerCase()); });
  return map;
}

function normKey_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[‐-‒–—―−\-]+/g, '')
    .replace(/[\s\u00A0\u3000]+/g, '')
    .replace(/（/g,'(').replace(/）/g,')')
    .replace(/．/g,'.')
    .replace(/µg|μg/gi,'ug');
}
function dropBrand_(s) { return String(s || '').replace(BRAND_PREFIX_RE, '').trim(); }
function specMinimalKey_(spec) {
  const t = String(spec || '').toLowerCase();
  const m = t.match(/(\d+(?:\.\d+)?)\s*(mg|mcg|ug|g|ml|iu|%)/i);
  if (!m) return '';
  return (m[1] + m[2]).replace(/\s+/g, '');
}
function cleanVendor_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/\s+/g,'')
    .replace(/(股份有限|有限|有限公司|股份|公司|實業|藥品|企業社|商行|藥局|藥品行)$/g,'');
}
function vendorMatch_(vendorFromMap, vendorFromResult) {
  if (!vendorFromMap) return true;
  if (!FUZZY_VENDOR) return vendorFromMap === vendorFromResult;
  const a = cleanVendor_(vendorFromMap);
  const b = cleanVendor_(vendorFromResult);
  return a === b || a.includes(b) || b.includes(a);
}
function qtyOrPlaceholder_(s) {
  const x = (s == null) ? '' : String(s).trim();
  if (!x || x.toLowerCase() === 'nan' || x.toLowerCase() === 'none') {
    return SHOW_IF_QTY_MISSING ? QTY_PLACEHOLDER : '';
  }
  if (/(盒|粒|顆|錠|瓶|支|包|條|罐)\b/.test(x)) return x;
  if (/\d$/.test(x)) return x + DEFAULT_UNIT;
  return x;
}
function makeLinesFromMap_(vendorText, mapRows) {
  const lines = [];
  const vendors = Object.keys(vendorText);

  vendors.forEach(vendorKey => {
    const bigKey = vendorText[vendorKey];
    const vendorKeyLower = (vendorKey || '').toLowerCase();

    const candidates = mapRows.filter(r => vendorMatch_(r.vendorNorm, vendorKeyLower));
    if (!candidates.length) return;

    const byNameCount = {};
    candidates.forEach(r => { byNameCount[r.nameNorm] = (byNameCount[r.nameNorm] || 0) + 1; });

    const matched = [];
    candidates.forEach(r => {
      const multiSpec = byNameCount[r.nameNorm] > 1;

      const tryKeys = [];
      if (multiSpec) {
        if (r.nameSpecKey) tryKeys.push(r.nameSpecKey);
        if (r.specMinimal) tryKeys.push(r.nameKey + r.specMinimal);
        if (r.nameSpecNoBrandKey) tryKeys.push(r.nameSpecNoBrandKey);
        if (r.specMinimal) tryKeys.push(r.nameNoBrandKey + r.specMinimal);
      } else {
        tryKeys.push(r.nameKey, r.nameNoBrandKey);
      }

      const hit = tryKeys.some(k => k && bigKey.includes(k));
      if (!hit) return;

      const qty = qtyOrPlaceholder_(r.qty);
      if (!qty && !SHOW_IF_QTY_MISSING) return;

      matched.push({
        name: r.name,
        nameKey: r.nameKey,
        nameNoBrandKey: r.nameNoBrandKey,
        spec: r.spec,
        specMinimal: r.specMinimal,
        qty,
        multiSpec
      });
    });

    if (!matched.length) return;

    const items = [];
    const seen = new Set();
    matched.forEach(m => {
      const uniq = (m.nameKey || m.nameNoBrandKey) + '||' + (m.specMinimal || '');
      if (seen.has(uniq)) return;
      seen.add(uniq);

      const needSpec = INCLUDE_SPEC || (INCLUDE_SPEC_SMART && (m.multiSpec && m.spec));
      const disp = needSpec && m.spec ? `${m.name} ${m.spec} ${m.qty}` : `${m.name} ${m.qty}`;
      items.push(disp);
    });

    if (items.length) {
      const vendorDisp = vendorKey || '（未指定廠商）';
      lines.push([vendorDisp, `${vendorDisp}想訂` + items.join('、')]);
    }
  });

  return lines.sort((a, b) => a[0].localeCompare(b[0], 'zh-Hant'));
}

function writeOutput_(ss, lines) {
  const sheet = ss.getSheetByName(OUTPUT_SHEET) || ss.insertSheet(OUTPUT_SHEET);
  sheet.clear();
  const out = [['廠商', '訂單文字']].concat(lines);
  sheet.getRange(1, 1, out.length, out[0].length).setValues(out);
}

/* -------- 共用小工具 -------- */
function colIndexFinder_(headers, dict) {
  const norm = s => String(s || '').replace(/\s+/g, '').toLowerCase();
  const findOne = keys => {
    const set = new Set(headers.map(h => norm(h)));
    for (const k of keys) {
      const key = norm(k);
      if (set.has(key)) return headers.findIndex(h => norm(h) === key);
    }
    for (let i = 0; i < headers.length; i++) {
      const h = norm(headers[i]);
      if (keys.some(k => h.includes(norm(k)))) return i;
    }
    return -1;
  };
  return {
    vendor: findOne(dict.vendor || []),
    name:   findOne(dict.name || []),
    spec:   findOne(dict.spec || []),
    info:   findOne(dict.info || []),
  };
}
function pickCell_(row, headers, idx) {
  if (idx < 0) return '';
  const v = row[idx];
  return v == null ? '' : v;
}
function normText_(s) {
  let t = (s == null) ? '' : String(s);
  t = t.replace(/\s*[（(]兩倍安全庫存[:：][^)）]*[)）]\s*/g, ' ');
  t = t.replace(/\s*[（(]安全庫存[:：][^)）]*[)）]\s*/g, ' ');
  return t.trim().replace(/\s+/g, ' ');
}

/* =============================================================
 *   新增：從本機 JSON 填入耗量（完全一致）＋「第N週…」彈性標題
 *   - 只在 JSON.drug.trim() 與「9」分頁【藥名成】完全一致時寫入
 *   - 週欄允許：第一週 / 第1週 / 第一週01-06 / 第四週22-27 …（只要以「第N週」開頭）
 *   - 功能入口：叫藥自動化 → 從本機 JSON 填入耗量（完全一致）
 * ============================================================= */
const WF_CFG = {
  TARGET_SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
  TARGET_SHEET_NAME: '9',     // 分頁名稱
  COL_TARGET_NAME: '藥名成',   // 以此欄位做「完全一致」比對
  WEEK_LABELS: { 1:'第一週', 2:'第二週', 3:'第三週', 4:'第四週', 5:'第五週' }
};

// 顯示上傳對話框（需要 Upload.html）
function WF_showUploadDialog(){
  const html = HtmlService.createHtmlOutputFromFile('Upload')
    .setTitle('從本機 JSON 填入耗量（完全一致）')
    .setWidth(520).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, '從本機 JSON 填入耗量（完全一致）');
}

// 後端：接收 JSON 字串 + 週別 → 僅完全一致時寫入
function WF_processUploadedJson(jsonText, weekNumber){
  if (!(weekNumber >=1 && weekNumber <=5)) throw new Error('週別需為 1~5');

  // 解析 JSON（格式：[{"drug":"...", "total":123}, ...]）
  let stats = JSON.parse(jsonText);
  if (!Array.isArray(stats)) throw new Error('JSON 根節點必須是陣列');
  stats = stats.map(r => ({
    drug: String((r && r.drug) || '').trim(),
    total: Number(r && r.total)
  })).filter(r => r.drug && !isNaN(r.total));
  if (!stats.length) throw new Error('JSON 無有效資料（需包含 drug 與 total）');

  // 讀目標表
  const ss = SpreadsheetApp.openById(WF_CFG.TARGET_SPREADSHEET_ID);
  const sh = ss.getSheetByName(WF_CFG.TARGET_SHEET_NAME);
  if (!sh) throw new Error(`找不到分頁：${WF_CFG.TARGET_SHEET_NAME}`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('目標分頁沒有資料或沒有表頭');

  const headers = values[0].map(v => (v||'').toString().trim());

  // 找「藥名成」欄（等值）
  const cName = WF_findHeaderIndexEqual_(headers, WF_CFG.COl_TARGET_NAME || WF_CFG.COL_TARGET_NAME);
  // 上行防呆（有人曾拼錯 key）
  if (cName === -1) {
    const i = headers.indexOf(WF_CFG.COL_TARGET_NAME);
    if (i === -1) throw new Error(`找不到欄位：${WF_CFG.COL_TARGET_NAME}`);
    else {} // 用 i
  }

  const nameColIndex = headers.indexOf(WF_CFG.COL_TARGET_NAME);
  const cWeek = WF_findWeekColumnIndex_(headers, weekNumber);
  if (cWeek === -1) throw new Error(`找不到週欄：以「第${weekNumber}週」或「第${toZhNum_(weekNumber)}週」開頭的任一欄位`);

  // 建立「完全一致」索引：藥名成（trim 後）→ 多列 row 索引（1-based, 含表頭）
  const nameIndex = new Map();
  for (let r = 1; r < values.length; r++) {
    const name = String(values[r][nameColIndex] || '').trim();
    if (!name) continue;
    if (!nameIndex.has(name)) nameIndex.set(name, []);
    nameIndex.get(name).push(r); // r 是資料列（從 1 開始，0=表頭）
  }

  // 只更新週欄
  const headerRow = 1;
  const rowCount = values.length - headerRow;
  const weekColRange = sh.getRange(headerRow+1, cWeek+1, rowCount, 1);
  const weekColVals = weekColRange.getValues();

  let writtenCells = 0;
  let matchedDrugs = 0;
  const miss = [];
  const dupTouched = [];

  // 逐筆 JSON → 完全一致才寫
  for (const {drug, total} of stats) {
    const rows = nameIndex.get(drug);
    if (rows && rows.length) {
      matchedDrugs++;
      rows.forEach(r1 => {
        const i = r1 - headerRow;
        if (i >= 0 && i < weekColVals.length) {
          weekColVals[i][0] = total;
          writtenCells++;
        }
      });
      if (rows.length > 1) dupTouched.push({ name: drug, rows: rows.map(x=>x+1) });
    } else {
      miss.push(drug);
    }
  }

  weekColRange.setValues(weekColVals);

  return {
    weekLabel: headers[cWeek] || WF_CFG.WEEK_LABELS[weekNumber],
    writtenCells,
    matchedDrugs,
    missedCount: miss.length,
    missedExamples: miss.slice(0, 10),
    duplicatesUpdated: dupTouched.slice(0, 10)
  };
}

/* ==== 小工具（等值查找/週欄查找/中文數字） ==== */
function WF_findHeaderIndexEqual_(headers, key){
  for (let i=0;i<headers.length;i++) if (headers[i] === key) return i;
  return -1;
}
function toZhNum_(n){
  const zh = ['零','一','二','三','四','五','六','七','八','九','十'];
  return zh[n] || String(n);
}
function WF_findWeekColumnIndex_(headers, number){
  const zh = toZhNum_(number);           // 一/二/三/四/五
  const prefixA = `第${zh}週`;           // 第一週
  const prefixB = `第${number}週`;       // 第1週
  // 規範：移空白，將「周」→「週」，然後檢查是否「以 prefix 開頭」
  for (let i=0; i<headers.length; i++){
    const h = String(headers[i] || '')
      .replace(/\s+/g,'')
      .replace(/周/g,'週');
    if (h.startsWith(prefixA) || h.startsWith(prefixB)) return i;
  }
  // 次要：寬鬆（原字串做開頭比對，允許中間有空白）
  const re = new RegExp(`^\\s*第\\s*(${zh}|${number})\\s*週`, 'i');
  for (let i=0; i<headers.length; i++){
    const h = String(headers[i] || '').replace(/周/g,'週');
    if (re.test(h)) return i;
  }
  return -1;
}
