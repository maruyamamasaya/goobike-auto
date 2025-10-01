/********************************************************
 * Goobike PAS 自動入力（Z列見出し版・321〜400行バッチ／進捗ログつき）
 * - シート: 新・在庫台帳
 * - 対象: 321〜400行 かつ C列が「掲載（グーのみ）」
 * - Z列ヘッダで読み取り、欠損はデフォルト補完
 * - サイドバーでリアルタイム進捗ログ＆行ごとのブックマークレット提示
 ********************************************************/

var TARGET_SHEET = '新・在庫台帳';
var START_ROW = 321;
var END_ROW   = 400;
var CHUNK_SIZE = 10; // 一度に処理する行数（進捗更新単位）

/** Z列からの見出し順（セルの表示値と一致させる） */
var HEADERS_Z = [
  'メーカー\ngoobike','排気量\ngoobike','車種','支払総額','区分','モデル年式','初度登録年','車検/自賠責',
  '製造国','排気量','修復歴','タイプ','メーカー認定','メーカー保証','販売店保証','整備','構造変更済み',
  'ABS','品質評価書','ワンオーナー','ノーマル車','逆輸入車','通信販売可能車','社外マフラー','社外メーター',
  'オーディオ','セキュリティ','セル付','ナビ','フルカスタム','FI車','４スト','LED/HID付','ETC','ボアアップ車','MT'
];

/** 必須相当のデフォルト補完値（必要に応じて調整可） */
var DEFAULTS = {
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

/** チェック群（空でなければON扱いにするキー） */
var CHECK_KEYS = [
  'メーカー認定','メーカー保証','販売店保証','整備','構造変更済み','ABS','品質評価書','ワンオーナー',
  'ノーマル車','逆輸入車','通信販売可能車','社外マフラー','社外メーター','オーディオ','セキュリティ',
  'セル付','ナビ','フルカスタム','FI車','４スト','LED/HID付','ETC','ボアアップ車','MT'
];

/** メニュー追加 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Goobike連携')
    .addItem('進捗ログつきサイドバー（321–400行）', 'openSidebar')
    .addToUi();
}

/** サイドバーを開く */
function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Goobike連携（進捗ログ）');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** 列記号 → 列番号（A=1） */
function colToIndex_(col) {
  var n = 0;
  for (var i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}

/** ブラウザ側で実行する：ページ入力＆一時保存ロジック（文字列として返す） */
function makePASFiller_(DATA) {
  return "" +
"(function(){\n" +
"  var D = " + JSON.stringify(DATA) + ";\n" +
"  var norm = function(s){ return (s||'').toString().replace(/\\s+/g,'').toLowerCase(); };\n" +
"  var $  = function(s,r){ return (r||document).querySelector(s); };\n" +
"  var $$ = function(s,r){ return Array.prototype.slice.call((r||document).querySelectorAll(s)); };\n" +

"  function findFieldByLabel(labelText){\n" +
"    if(!labelText) return null;\n" +
"    var target = norm(labelText);\n" +
"    var labels = $$('label');\n" +
"    for(var i=0;i<labels.length;i++){\n" +
"      var lb = labels[i];\n" +
"      var txt = norm(lb.textContent||'');\n" +
"      if(txt.indexOf(target) !== -1){\n" +
"        var forId = lb.getAttribute('for');\n" +
"        if(forId && document.getElementById(forId)) return document.getElementById(forId);\n" +
"        var near = lb.parentElement && lb.parentElement.querySelector('select, input, textarea');\n" +
"        if(near) return near;\n" +
"      }\n" +
"    }\n" +
"    var heads = $$('h2,h3,dt,div,p,span');\n" +
"    for(var j=0;j<heads.length;j++){\n" +
"      var h = heads[j];\n" +
"      var t = norm(h.textContent||'');\n" +
"      if(t.indexOf(target) !== -1){\n" +
"        var near2 = h.parentElement && h.parentElement.querySelector('select, input, textarea');\n" +
"        if(near2) return near2;\n" +
"      }\n" +
"    }\n" +
"    return null;\n" +
"  }\n" +

"  function setSelectByText(labelOrEl, want){\n" +
"    if(!want) return;\n" +
"    var el = (labelOrEl && labelOrEl.nodeType===1) ? labelOrEl : findFieldByLabel(labelOrEl);\n" +
"    if(!el || !el.options) return;\n" +
"    var w = norm(want);\n" +
"    var hit = false;\n" +
"    for(var i=0;i<el.options.length;i++){\n" +
"      var opt = el.options[i];\n" +
"      var t = norm(opt.textContent||opt.label||'');\n" +
"      if(t.indexOf(w)!==-1 || w.indexOf(t)!==-1){ opt.selected = true; hit = true; break; }\n" +
"    }\n" +
"    el.dispatchEvent(new Event('change',{bubbles:true}));\n" +
"    if(!hit){ el.value = want; el.dispatchEvent(new Event('change',{bubbles:true})); }\n" +
"  }\n" +

"  function setRadioByText(label, wantText){\n" +
"    if(!wantText) return;\n" +
"    var anchor = findFieldByLabel(label);\n" +
"    var root = anchor ? (anchor.closest && anchor.closest('section, fieldset, div, form')) || document : document;\n" +
"    var w = norm(wantText);\n" +
"    var rads = $$('input[type=\"radio\"]', root);\n" +
"    for(var i=0;i<rads.length;i++){\n" +
"      var r = rads[i];\n" +
"      var lab = r.closest ? r.closest('label') : (r.parentElement||null);\n" +
"      var txt = norm(lab ? lab.textContent : '');\n" +
"      if(txt.indexOf(w)!==-1){ r.click(); return; }\n" +
"    }\n" +
"  }\n" +

"  function setCheckboxByText(label, textList){\n" +
"    if(!textList || !textList.length) return;\n" +
"    var anchor = findFieldByLabel(label);\n" +
"    var root = anchor ? (anchor.closest && anchor.closest('section, fieldset, div, form')) || document : document;\n" +
"    var wants = textList.map(function(t){ return norm(t); });\n" +
"    var cks = $$('input[type=\"checkbox\"]', root);\n" +
"    for(var i=0;i<cks.length;i++){\n" +
"      var c = cks[i];\n" +
"      var lab = c.closest ? c.closest('label') : (c.parentElement||null);\n" +
"      var txt = norm(lab ? lab.textContent : '');\n" +
"      for(var k=0;k<wants.length;k++){\n" +
"        if(txt.indexOf(wants[k])!==-1){ if(!c.checked) c.click(); break; }\n" +
"      }\n" +
"    }\n" +
"  }\n" +

"  function setTextByLabel(label,val){\n" +
"    if(val===null || val===undefined || val==='') return;\n" +
"    var el = findFieldByLabel(label);\n" +
"    if(!el) return;\n" +
"    el.focus(); el.value = val;\n" +
"    el.dispatchEvent(new Event('input',{bubbles:true}));\n" +
"    el.dispatchEvent(new Event('change',{bubbles:true}));\n" +
"  }\n" +

"  /* ====== 入力マッピング ====== */\n" +
"  setSelectByText('メーカー',      D['メーカー\\ngoobike']);\n" +
"  setSelectByText('排気区分',      D['排気量\\ngoobike']);\n" +
"  setSelectByText('車種',          D['車種']);\n" +
"  setSelectByText('排気量',        D['排気量']);\n" +
"  setTextByLabel('支払総額',       D['支払総額']);\n" +
"  setRadioByText('区分',           D['区分']);\n" +
"  setRadioByText('モデル年式',     D['モデル年式']);\n" +
"  setSelectByText('モデル年式',    D['モデル年式']);\n" +
"  setRadioByText('初度登録年',     D['初度登録年']);\n" +
"  setSelectByText('初度登録年',    D['初度登録年']);\n" +
"  setRadioByText('車検・自賠責保険', D['車検/自賠責']);\n" +
"  setSelectByText('製造国',        D['製造国']);\n" +
"  setRadioByText('修復歴',         D['修復歴']);\n" +
"  setSelectByText('タイプ',        D['タイプ']);\n" +

"  function ON(k){ return (D[k]||'') === '1'; }\n" +
"  var marks = [];\n" +
"  if(ON('メーカー認定')) marks.push('メーカー認定');\n" +
"  if(ON('メーカー保証')) marks.push('メーカー保証');\n" +
"  if(ON('販売店保証'))   marks.push('保証');\n" +
"  if(ON('整備'))         marks.push('整備');\n" +

"  var opts = [];\n" +
"  if(ON('構造変更済み'))   opts.push('構造変更済み');\n" +
"  if(ON('ABS'))            opts.push('ABS');\n" +
"  if(ON('品質評価書'))     opts.push('品質評価書');\n" +
"  if(ON('ワンオーナー'))   opts.push('ワンオーナー');\n" +
"  if(ON('ノーマル車'))     opts.push('ノーマル車');\n" +
"  if(ON('逆輸入車'))       opts.push('逆輸入車');\n" +
"  if(ON('通信販売可能車')) opts.push('通信販売可能車');\n" +
"  if(ON('社外マフラー'))   opts.push('社外マフラー');\n" +
"  if(ON('社外メーター'))   opts.push('社外メーター');\n" +
"  if(ON('オーディオ'))     opts.push('オーディオ');\n" +
"  if(ON('セキュリティ'))   opts.push('セキュリティ');\n" +
"  if(ON('セル付'))         opts.push('セル付');\n" +
"  if(ON('ナビ'))           opts.push('ナビ');\n" +
"  if(ON('フルカスタム'))   opts.push('フルカスタム');\n" +
"  if(ON('FI車'))           opts.push('FI車');\n" +
"  if(ON('４スト'))         opts.push('４スト');\n" +
"  if(ON('LED/HID付'))      opts.push('LED/HID付');\n" +
"  if(ON('ETC'))            opts.push('ETC');\n" +
"  if(ON('ボアアップ車'))   opts.push('ボアアップ車');\n" +
"  if(ON('MT'))             opts.push('MT');\n" +

"  if(marks.length) setCheckboxByText('マーク', marks);\n" +
"  if(opts.length)  setCheckboxByText('オプション', opts);\n" +

"  var btns = $$('button, input[type=\"button\"], input[type=\"submit\"]');\n" +
"  var save = null;\n" +
"  for(var i=0;i<btns.length;i++){\n" +
"    var t = (btns[i].value || btns[i].textContent || '').trim();\n" +
"    if(/一時保存/.test(t)){ save = btns[i]; break; }\n" +
"  }\n" +
"  if (save) {\n" +
"    save.click();\n" +
"    setTimeout(function(){ alert('✅ 入力して「一時保存」をクリックしました。結果を確認してください。'); }, 300);\n" +
"  } else {\n" +
"    alert('✅ 入力を反映しました（「一時保存」ボタンは見つかりませんでした）。ページは開いたままでOKです。');\n" +
"  }\n" +
"})();\n";
}

/** 1チャンク処理：startIndex（0基点）から size 件分スキャンして結果返す */
function processChunk(startIndex, size) {
  var sh = SpreadsheetApp.getActive().getSheetByName(TARGET_SHEET);
  if (!sh) throw new Error('シートが見つかりません: ' + TARGET_SHEET);

  // パラメータの防御
  var totalRows = Math.max(0, END_ROW - START_ROW + 1);
  var idx = Math.max(0, parseInt(startIndex, 10) || 0);
  var step = Math.max(1, parseInt(size, 10) || CHUNK_SIZE);

  if (totalRows === 0) {
    return { done: true, items: [], logs: ['対象範囲に行がありません'], nextIndex: idx, totalRows: 0 };
  }
  if (idx >= totalRows) {
    return { done: true, items: [], logs: ['done'], nextIndex: totalRows, totalRows: totalRows };
  }

  // from は必ず 1 以上になるように
  var from = START_ROW + idx;              // シート上の開始行
  if (from < 1) from = 1;

  // 残り行数（from が範囲外に出ていないかチェック）
  var remaining = END_ROW - from + 1;
  if (remaining <= 0) {
    return { done: true, items: [], logs: ['done'], nextIndex: totalRows, totalRows: totalRows };
  }

  var count = Math.min(step, remaining);

  var logs = [];
  logs.push('🔎 範囲読み込み: Row ' + from + '〜' + (from + count - 1));

  var colCIdx   = colToIndex_('C');
  var startColZ = colToIndex_('Z');

  // getRange の行・列・件数は必ず 1 以上に
  var rngC = sh.getRange(from, colCIdx, count, 1).getDisplayValues();
  var rngZ = sh.getRange(from, startColZ, count, HEADERS_Z.length).getDisplayValues();

  var items = []; // { row, bookmarklet, summary }
  var hit = 0, skipped = 0, errors = 0;

  for (var i = 0; i < count; i++) {
    var rowNum = from + i;
    var cVal = (rngC[i][0] || '').toString().trim();

    if (cVal !== '掲載（グーのみ）') { skipped++; continue; }

    try {
      var vals = rngZ[i];
      var rec = {};
      for (var j = 0; j < HEADERS_Z.length; j++) {
        var header = HEADERS_Z[j];
        var raw = (vals[j] || '').toString().trim();
        rec[header] = raw || DEFAULTS[header] || '';
      }
      // チェック群は空でなければ '1'
      for (var k = 0; k < CHECK_KEYS.length; k++) {
        var ck = CHECK_KEYS[k];
        rec[ck] = rec[ck] ? '1' : '';
      }

      var filler = makePASFiller_(rec);
      var bm = 'javascript:' + encodeURIComponent(filler);

      var summary = (rec['メーカー\ngoobike'] || '') + ' / ' +
                    (rec['車種'] || '') + ' / 総額:' +
                    (rec['支払総額'] || '');

      items.push({ row: rowNum, bookmarklet: bm, summary: summary });
      hit++;
    } catch (e) {
      errors++;
      logs.push('❌ 失敗: Row ' + rowNum + ' / ' + (e && e.message ? e.message : e));
    }
  }

  logs.push('✅ チャンク完了: ヒット ' + hit + '件, スキップ ' + skipped + '件, エラー ' + errors + '件');

  var nextIndex = idx + count;
  var done = nextIndex >= totalRows;

  return {
    done: done,
    nextIndex: done ? totalRows : nextIndex, // 常に数値を返す
    items: items,
    logs: logs,
    totalRows: totalRows
  };
}
