/********************************************************
 * Goobike PAS 自動入力（Z列見出し版）
 * - アクティブ行の「Z列から」必要項目を読み取り
 * - 欠損はデフォルト補完
 * - ブラウザ側で実行するブックマークレットを生成
 * - 321〜400行のうち掲載区分が「掲載（グーのみ）」の行も一括生成
 ********************************************************/

const PAS_TARGET_SHEET = '新・在庫台帳';

/** Z列からの見出し順（セルの表示値を使います） */
const PAS_HEADERS_Z = [
  'メーカー\ngoobike', '排気量\ngoobike', '車種', '支払総額', '区分', 'モデル年式', '初度登録年',
  '車検/自賠責', '製造国', '排気量', '修復歴', 'タイプ', 'メーカー認定', 'メーカー保証', '販売店保証',
  '整備', '構造変更済み', 'ABS', '品質評価書', 'ワンオーナー', 'ノーマル車', '逆輸入車', '通信販売可能車',
  '社外マフラー', '社外メーター', 'オーディオ', 'セキュリティ', 'セル付', 'ナビ', 'フルカスタム', 'FI車',
  '４スト', 'LED/HID付', 'ETC', 'ボアアップ車', 'MT'
];

/** 欠損時のデフォルト（必須相当のもの中心） */
const PAS_DEFAULTS = {
  'メーカー\ngoobike': 'ホンダ',
  '排気量\ngoobike': '～125cc',
  '車種': 'スーパーカブ110',
  '区分': '中古車',
  '修復歴': 'なし',
  '製造国': '日本',
  '排気量': '110',
  '支払総額': '29.8',
  'モデル年式': '未記入',
  '初度登録年': '不明',
  '車検/自賠責': 'なし'
};

/** チェック系（「空でなければON」扱い） */
const PAS_CHECK_KEYS = [
  'メーカー認定','メーカー保証','販売店保証','整備','構造変更済み','ABS','品質評価書','ワンオーナー',
  'ノーマル車','逆輸入車','通信販売可能車','社外マフラー','社外メーター','オーディオ','セキュリティ',
  'セル付','ナビ','フルカスタム','FI車','４スト','LED/HID付','ETC','ボアアップ車','MT'
];

const PAS_TEST_RANGE_START = 321;
const PAS_TEST_RANGE_END   = 400;
const PAS_TEST_COLUMN = 3; // C 列
const PAS_TEST_VALUE  = '掲載（グーのみ）';

function addBookmarkletMenu_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Goobike連携')
    .addItem('この行のブックマークレット生成（Z列見出し版）', 'generateBookmarkletFromRow_Z')
    .addItem('321〜400行（掲載グーのみ）一覧', 'generateBookmarkletsForTestRange_Z')
    .addSeparator()
    .addItem('（任意）Z列ヘッダを書き込む', 'writeZHeaders')
    .addToUi();
}

/** Z列に見出し行を出す（任意） */
function writeZHeaders() {
  const sh = SpreadsheetApp.getActive().getSheetByName(PAS_TARGET_SHEET);
  if (!sh) return;
  const startCol = colToIndex_('Z');
  PAS_HEADERS_Z.forEach(function(h, i) { sh.getRange(1, startCol + i).setValue(h); });
  SpreadsheetApp.getUi().alert('Z列にヘッダを書き込みました。');
}

/** 列記号→番号 */
function colToIndex_(col) {
  var n = 0;
  for (var i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
  return n;
}

/** アクティブ行→データ抽出→ブックマークレット生成 */
function generateBookmarkletFromRow_Z() {
  const sh = SpreadsheetApp.getActive().getSheetByName(PAS_TARGET_SHEET);
  if (!sh) { SpreadsheetApp.getUi().alert('シートが見つかりません: ' + PAS_TARGET_SHEET); return; }

  const cell = sh.getActiveCell();
  if (!cell) { SpreadsheetApp.getUi().alert('アクティブセルが見つかりません。'); return; }

  const row = cell.getRow();
  const data = buildBookmarkletPayload_(sh, row);
  if (!data) { SpreadsheetApp.getUi().alert('ブックマークレットの生成に失敗しました。'); return; }

  showBookmarkletDialog_([{ row: row, bookmarklet: data.bookmarklet }], 'ブックマークレット生成（Z列見出し版）');
}

/** 321〜400行（掲載グーのみ）を対象にブックマークレットを生成 */
function generateBookmarkletsForTestRange_Z() {
  const sh = SpreadsheetApp.getActive().getSheetByName(PAS_TARGET_SHEET);
  if (!sh) { SpreadsheetApp.getUi().alert('シートが見つかりません: ' + PAS_TARGET_SHEET); return; }

  const results = [];
  for (var row = PAS_TEST_RANGE_START; row <= PAS_TEST_RANGE_END; row++) {
    const display = sh.getRange(row, PAS_TEST_COLUMN).getDisplayValue();
    if (display !== PAS_TEST_VALUE) continue;
    const payload = buildBookmarkletPayload_(sh, row);
    if (payload) {
      results.push({ row: row, bookmarklet: payload.bookmarklet });
    }
  }

  if (results.length === 0) {
    SpreadsheetApp.getUi().alert('条件に該当する行がありませんでした。');
    return;
  }

  showBookmarkletDialog_(results, '掲載（グーのみ）対象行のブックマークレット');
}

function buildBookmarkletPayload_(sheet, row) {
  try {
    const startCol = colToIndex_('Z');
    const values = sheet.getRange(row, startCol, 1, PAS_HEADERS_Z.length).getDisplayValues()[0];
    const data = {};
    PAS_HEADERS_Z.forEach(function(header, i) {
      const raw = values[i] ? values[i].toString().trim() : '';
      const fallback = PAS_DEFAULTS.hasOwnProperty(header) ? PAS_DEFAULTS[header] : '';
      data[header] = raw || fallback || '';
    });

    PAS_CHECK_KEYS.forEach(function(key) {
      const v = data[key] ? data[key].toString().trim() : '';
      data[key] = v ? '1' : '';
    });

    const filler = makePASFiller_(data);
    const bookmarklet = 'javascript:' + encodeURIComponent(filler);
    return { data: data, bookmarklet: bookmarklet };
  } catch (e) {
    Logger.log('ブックマークレット生成に失敗しました: row=' + row + ', error=' + e);
    return null;
  }
}

function showBookmarkletDialog_(items, title) {
  const rowsHtml = items.map(function(item, index) {
    const textareaId = 'bookmarklet-' + index;
    const escaped = escapeHtml_(item.bookmarklet);
    return '<section class="item">' +
      '<h4>' + item.row + ' 行目</h4>' +
      '<textarea id="' + textareaId + '" class="code">' + escaped + '</textarea>' +
      '<button onclick="copyBookmarklet(\'' + textareaId + '\')">コピー</button>' +
      '</section>';
  }).join('');

  const html = HtmlService.createHtmlOutput('\
    <style>\
      body { font-family: system-ui, -apple-system, \"Segoe UI\", sans-serif; margin: 0; padding: 16px; background: #f9fafb; }\
      h3 { margin: 0 0 12px; font-size: 18px; }\
      .item { background: #fff; border: 1px solid #e5e7eb; border-radius: 8px; padding: 12px; margin-bottom: 12px; }\
      .item h4 { margin: 0 0 8px; font-size: 15px; }\
      .code { width: 100%; height: 140px; font-family: ui-monospace, SFMono-Regular, SFMono-Regular, Menlo, Monaco, Consolas, \"Liberation Mono\", monospace; font-size: 12px; border: 1px solid #d1d5db; border-radius: 6px; padding: 8px; box-sizing: border-box; background: #f3f4f6; }\
      button { margin-top: 8px; padding: 6px 12px; font-size: 12px; border-radius: 6px; border: none; background: #2563eb; color: #fff; cursor: pointer; }\
      button:hover { background: #1d4ed8; }\
      .note { margin-top: 12px; color: #6b7280; font-size: 12px; }\
    </style>\
    <h3>' + escapeHtml_(title) + '</h3>\
    <p>下のテキストを<strong>まるごと</strong>ブックマークのURLに貼り付けて保存し、対象ページでクリックしてください。</p>\
    ' + rowsHtml + '\
    <p class="note">必須項目が空欄の場合はデフォルト値を補完しています。保存前に内容をご確認ください。</p>\
    <script>\
      function copyBookmarklet(id) {\
        var el = document.getElementById(id);\
        if (!el) return;\
        el.focus();\
        el.select();\
        document.execCommand(\"copy\");\
        alert(\"コピーしました\");\
      }\
    </script>\
  ').setWidth(760).setHeight(Math.min(520, 200 + items.length * 170));

  SpreadsheetApp.getUi().showModalDialog(html, title);
}

function escapeHtml_(value) {
  return (value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function makePASFiller_(DATA) {
  return "\
(() => {\
  const D = " + JSON.stringify(DATA) + ";\
  const norm = s => (s||'').toString().replace(/\s+/g,'').toLowerCase();\
  const $  = (s,r=document)=>r.querySelector(s);\
  const $$ = (s,r=document)=>Array.from(r.querySelectorAll(s));\
\
  function findFieldByLabel(labelText){\
    if(!labelText) return null;\
    const target = norm(labelText);\
    for(const lb of $$('label')){\
      const txt = norm(lb.textContent||'');\
      if(txt.includes(target)){\
        const forId = lb.getAttribute('for');\
        if(forId && document.getElementById(forId)) return document.getElementById(forId);\
        const near = lb.parentElement?.querySelector('select, input, textarea');\
        if(near) return near;\
      }\
    }\
    for(const h of $$('h2,h3,dt,div,p,span')){\
      const txt = norm(h.textContent||'');\
      if(txt.includes(target)){\
        const near = h.parentElement?.querySelector('select, input, textarea');\
        if(near) return near;\
      }\
    }\
    return null;\
  }\
\
  function setSelectByText(labelOrEl, want){\
    if(!want) return;\
    const el = (labelOrEl instanceof Element) ? labelOrEl : findFieldByLabel(labelOrEl);\
    if(!el) return;\
    const w = norm(want);\
    let hit = false;\
    for(const opt of el.options){\
      const t = norm(opt.textContent||opt.label||'');\
      if(t.includes(w) || w.includes(t)){ opt.selected = true; hit = true; break; }\
    }\
    el.dispatchEvent(new Event('change',{bubbles:true}));\
    if(!hit){ el.value = want; el.dispatchEvent(new Event('change',{bubbles:true})); }\
  }\
\
  function setRadioByText(label, wantText){\
    if(!wantText) return;\
    const anchor = findFieldByLabel(label);\
    const root = anchor ? anchor.closest('section, fieldset, div, form') || document : document;\
    const w = norm(wantText);\
    for(const r of $$('input[type=\"radio\"]', root)){\
      const lab = r.closest('label') || r.parentElement;\
      const txt = norm(lab ? lab.textContent : '');\
      if(txt.includes(w)){ r.click(); return; }\
    }\
  }\
\
  function setCheckboxByText(label, textList){\
    if(!textList || !textList.length) return;\
    const anchor = findFieldByLabel(label);\
    const root = anchor ? anchor.closest('section, fieldset, div, form') || document : document;\
    const wants = textList.map(t=>norm(t));\
    for(const c of $$('input[type=\"checkbox\"]', root)){\
      const lab = c.closest('label') || c.parentElement;\
      const txt = norm(lab ? lab.textContent : '');\
      if(wants.some(w=>txt.includes(w))){ if(!c.checked) c.click(); }\
    }\
  }\
\
  function setTextByLabel(label,val){\
    if(!val) return;\
    const el = findFieldByLabel(label);\
    if(!el) return;\
    el.focus(); el.value = val;\
    el.dispatchEvent(new Event('input',{bubbles:true}));\
    el.dispatchEvent(new Event('change',{bubbles:true}));\
  }\
\
  setSelectByText('メーカー', D['メーカー\\ngoobike']);\
  setSelectByText('排気区分', D['排気量\\ngoobike']);\
  setSelectByText('車種', D['車種']);\
  setSelectByText('排気量', D['排気量']);\
  setTextByLabel('支払総額', D['支払総額']);\
  setRadioByText('区分', D['区分']);\
  setRadioByText('モデル年式', D['モデル年式']);\
  setSelectByText('モデル年式', D['モデル年式']);\
  setRadioByText('初度登録年', D['初度登録年']);\
  setSelectByText('初度登録年', D['初度登録年']);\
  setRadioByText('車検・自賠責保険', D['車検/自賠責']);\
  setSelectByText('製造国', D['製造国']);\
  setRadioByText('修復歴', D['修復歴']);\
  setSelectByText('タイプ', D['タイプ']);\
\
  const ON = k => (D[k]||'') === '1';\
  const marks = [];\
  if(ON('メーカー認定'))   marks.push('メーカー認定');\
  if(ON('メーカー保証'))   marks.push('メーカー保証');\
  if(ON('販売店保証'))     marks.push('保証');\
  if(ON('整備'))           marks.push('整備');\
  const opts = [];\
  if(ON('構造変更済み'))   opts.push('構造変更済み');\
  if(ON('ABS'))            opts.push('ABS');\
  if(ON('品質評価書'))     opts.push('品質評価書');\
  if(ON('ワンオーナー'))   opts.push('ワンオーナー');\
  if(ON('ノーマル車'))     opts.push('ノーマル車');\
  if(ON('逆輸入車'))       opts.push('逆輸入車');\
  if(ON('通信販売可能車')) opts.push('通信販売可能車');\
  if(ON('社外マフラー'))   opts.push('社外マフラー');\
  if(ON('社外メーター'))   opts.push('社外メーター');\
  if(ON('オーディオ'))     opts.push('オーディオ');\
  if(ON('セキュリティ'))   opts.push('セキュリティ');\
  if(ON('セル付'))         opts.push('セル付');\
  if(ON('ナビ'))           opts.push('ナビ');\
  if(ON('フルカスタム'))   opts.push('フルカスタム');\
  if(ON('FI車'))           opts.push('FI車');\
  if(ON('４スト'))         opts.push('４スト');\
  if(ON('LED/HID付'))      opts.push('LED/HID付');\
  if(ON('ETC'))            opts.push('ETC');\
  if(ON('ボアアップ車'))   opts.push('ボアアップ車');\
  if(ON('MT'))             opts.push('MT');\
\
  if(marks.length) setCheckboxByText('マーク', marks);\
  if(opts.length)  setCheckboxByText('オプション', opts);\
\
  const buttons = $$('button, input[type=\"button\"], input[type=\"submit\"]');\
  const cand = buttons.find(b => {\
    const t = (b.value || b.textContent || '').trim();\
    return /一時保存/.test(t);\
  });\
  if (cand) {\
    cand.click();\
    setTimeout(()=>alert('✅ 入力し「一時保存」をクリックしました。画面の結果をご確認ください。'), 300);\
  } else {\
    alert('✅ 入力を反映しました（「一時保存」ボタンは見つけられませんでした）。このページは開いたままでOKです。');\
  }\
})();\
";
}
