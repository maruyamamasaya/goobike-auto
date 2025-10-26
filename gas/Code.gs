/*******************************************************
 * 新・在庫台帳 用 スクリプト（ガード一括ON/OFF＋スナップ個別ON/OFF 完全版）
 * - 行/列追加の禁止（差し戻し）
 * - 範囲保護の設定/解除
 * - 自動更新パネル（サイドバー）表示
 * - ガード機能の一括ON/OFF（差し戻し/トースト/保護）
 * - スナップショット記録（行数/列数）だけ個別にON/OFF
 *******************************************************/

/** 対象シート名 */
const TARGET_SHEETS = ['新・在庫台帳'];
const TARGET_SHEET  = '新・在庫台帳';

/** プロパティキー */
const PROP_PREFIX         = 'banInsert:';        // スナップショット保存用（シートIDごと）
const PROP_GUARDS_KEY     = 'guardsEnabled';     // "true" | "false"（デフォルトtrue）
const PROP_SNAPSHOT_KEY   = 'snapshotEnabled';   // "true" | "false"（デフォルトtrue）


function resetGuardAndSnapshotSettings() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('guardsEnabled');
  props.deleteProperty('snapshotEnabled');
  SpreadsheetApp.getUi().alert('✅ ガード・スナップショット設定をリセットしました');
}

/* =========================
 *  ガード（全体）ON/OFF
 * ========================= */
function isGuardsEnabled_() {
  const v = PropertiesService.getScriptProperties().getProperty(PROP_GUARDS_KEY);
  return v === null ? true : v === 'true';
}
function enableGuards() {
  PropertiesService.getScriptProperties().setProperty(PROP_GUARDS_KEY, 'true');
  SpreadsheetApp.getActive().toast('✅ ガード機能を有効化しました', '状態', 5);
}
function disableGuards() {
  PropertiesService.getScriptProperties().setProperty(PROP_GUARDS_KEY, 'false');
  try { clearProtectionsForSheet(); } catch (_) {}
  SpreadsheetApp.getActive().toast('⛔ ガード機能を無効化・保護を解除しました', '状態', 6);
}

/* =========================
 *  スナップショット（サイズ記録）ON/OFF
 * ========================= */
function isSnapshotEnabled_() {
  const v = PropertiesService.getScriptProperties().getProperty(PROP_SNAPSHOT_KEY);
  return v === null ? true : v === 'true';
}
function enableSnapshot() {
  PropertiesService.getScriptProperties().setProperty(PROP_SNAPSHOT_KEY, 'true');
  SpreadsheetApp.getActive().toast('📸 スナップショット記録を再開しました', '状態', 5);
}
function disableSnapshot() {
  PropertiesService.getScriptProperties().setProperty(PROP_SNAPSHOT_KEY, 'false');
  SpreadsheetApp.getActive().toast('📸 スナップショット記録を停止しました', '状態', 5);
}

/* =========================
 *  メニュー
 * ========================= */
function addGuardMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('ガード切替')
    .addItem('ガードを有効化', 'enableGuards')
    .addItem('ガードを無効化（保護解除）', 'disableGuards')
    .addSeparator()
    .addItem('スナップショットを有効化', 'enableSnapshot')
    .addItem('スナップショットを無効化', 'disableSnapshot')
    .addToUi();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('管理者用')
    .addItem('自動更新パネルを開く', 'showSidebar')
    .addToUi();

  addGuardMenu_();
  if (typeof addBookmarkletMenu_ === 'function') addBookmarkletMenu_();
}

/* =========================
 *  onChange（警告表示のみ）
 * ========================= */
function onChange(e) {
  if (!isGuardsEnabled_()) return;              // ガードOFFなら何もしない
  if (!e || !e.changeType || !e.source) return;

  const ss = e.source;
  const sheet = ss.getActiveSheet();
  if (!sheet) return;
  if (!TARGET_SHEETS.includes(sheet.getName())) return;

  const type   = e.changeType;
  const isRow  = type === 'INSERT_ROW';
  const isCol  = type === 'INSERT_COLUMN';
  const isGrid = type === 'INSERT_GRID';

  if (isRow || isCol || isGrid) {
    const msg = isRow
      ? '🚫【新・在庫台帳】行の追加は禁止です！\n関数がズレる原因になります。'
      : isCol
        ? '🚫【新・在庫台帳】列の追加は禁止です！\n関数がズレる原因になります。'
        : '🚫【新・在庫台帳】行/列数の変更（グリッド拡張）は禁止です！\n関数がズレる原因になります。';

    ss.toast(msg, '禁止アラート', 7);
    SpreadsheetApp.getUi().alert(msg);
  }
}

/* =========================
 *  範囲保護 設定
 * ========================= */
function setupRangeProtections_BM() {
  if (!isGuardsEnabled_()) return; // ガードOFF時は保護を張らない

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(TARGET_SHEET);
  if (!sheet) throw new Error('対象シートが見つかりません: ' + TARGET_SHEET);

  // 既存保護を全削除
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  const maxR = sheet.getMaxRows();
  const maxC = sheet.getMaxColumns();

  const okLastRow = Math.min(3000, maxR);
  const bmColIdx  = Math.min(sheet.getRange('BM1').getColumn(), maxC); // BM=65列

  const ngRanges = [];

  // 1) ヘッダ 1行 全列
  if (maxR >= 1 && maxC >= 1) {
    ngRanges.push(sheet.getRange(1, 1, 1, maxC));
  }
  // 2) 3001行目以降
  if (maxR > okLastRow) ngRanges.push(sheet.getRange(okLastRow + 1, 1, maxR - okLastRow, maxC));
  // 3) BM右側
  if (maxC > bmColIdx) ngRanges.push(sheet.getRange(1, bmColIdx + 1, maxR, maxC - bmColIdx));

  const me = Session.getEffectiveUser();
  ngRanges.forEach(r => {
    const prot = r.protect();
    prot.setDescription('編集禁止ゾーン（関数ズレ防止）');
    prot.addEditor(me);
    prot.removeEditors(prot.getEditors().filter(u => u.getEmail() !== me.getEmail()));
    // prot.setWarningOnly(true); // 警告のみ許可にしたい場合
  });
}

/* =========================
 *  スナップショット 初期保存（任意）
 * ========================= */
function banInsertInit() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TARGET_SHEET);
  if (!sheet) return;
  _saveSize(sheet); // スナップショットがOFFの場合は内部で無視
}

/* =========================
 *  差し戻し（インストール型トリガー用）
 * ========================= */
function banInsertOnChange(e) {
  if (!isGuardsEnabled_()) return;              // ガードOFFなら差し戻しも無効
  if (!e || !e.changeType || !e.source) return;

  const sheet = e.source.getActiveSheet();
  if (!sheet || sheet.getName() !== TARGET_SHEET) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const type = e.changeType;

    if (type === 'INSERT_ROW') {
      const rng = sheet.getActiveRange();
      if (rng) {
        sheet.deleteRows(rng.getRow(), rng.getNumRows());
        _toast('🚫 行の追加は禁止です（元に戻しました）');
      } else {
        _toast('🚫 行の追加は禁止です');
      }
    } else if (type === 'INSERT_COLUMN') {
      const rng = sheet.getActiveRange();
      if (rng) {
        sheet.deleteColumns(rng.getColumn(), rng.getNumColumns());
        _toast('🚫 列の追加は禁止です（元に戻しました）');
      } else {
        _toast('🚫 列の追加は禁止です');
      }
    } else if (type === 'INSERT_GRID') {
      // 右端/下端の「行/列を追加」系はサイズ比較で差し戻し
      const props = PropertiesService.getScriptProperties();
      const key = PROP_PREFIX + sheet.getSheetId();
      let prev = {};
      try { prev = JSON.parse(props.getProperty(key) || '{}'); } catch (_){ }

      if (!prev.rows || !prev.cols) { _saveSize(sheet); return; }

      const curRows = sheet.getMaxRows();
      const curCols = sheet.getMaxColumns();

      if (curRows > prev.rows) {
        sheet.deleteRows(prev.rows + 1, curRows - prev.rows);
        _toast('🚫 行の追加は禁止です（元に戻しました）');
      }
      if (curCols > prev.cols) {
        sheet.deleteColumns(prev.cols + 1, curCols - prev.cols);
        _toast('🚫 列の追加は禁止です（元に戻しました）');
      }
    }

    _saveSize(sheet); // ← スナップOFF時は内部で拒否
  } finally {
    lock.releaseLock();
  }
}

/* =========================
 *  ユーティリティ
 * ========================= */
function _saveSize(sheet) {
  // ガードOFF時 or スナップOFF時は記録しない
  if (!isGuardsEnabled_())   return;
  if (!isSnapshotEnabled_()) return;

  const props = PropertiesService.getScriptProperties();
  const key = PROP_PREFIX + sheet.getSheetId();
  props.setProperty(key, JSON.stringify({
    rows: sheet.getMaxRows(),
    cols: sheet.getMaxColumns()
  }));
}

function _toast(msg) {
  SpreadsheetApp.getActive().toast(msg, '禁止', 6);
  // SpreadsheetApp.getUi().alert(msg); // 必要なら
}

/* =========================
 *  保護解除（手動）
 * ========================= */
function clearProtectionsForSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET);
  if (!sheet) throw new Error('シートが見つかりません: ' + TARGET_SHEET);

  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => { if (p.canEdit()) p.remove(); });
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => { if (p.canEdit()) p.remove(); });

  SpreadsheetApp.getActive().toast("✅ シート '" + TARGET_SHEET + "' の保護を解除しました", "完了", 5);
}

/* =========================
 *  サイドバー（ダミー表示）
 * ========================= */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('自動更新');
  SpreadsheetApp.getUi().showSidebar(html);
}

function simulateRun() {
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  const steps = [];
  const menuUrl = 'https://pas.goobike.com/php/client/menu.php';
  const regUrl  = 'https://pas.goobike.com/sa/bike_registration/main';

  let modelText = '';
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(TARGET_SHEET);
    modelText = sh ? sh.getRange('Y2').getDisplayValue() : '';
  } catch (e) {}

  const missing = [];
  if (!modelText) missing.push('車種');

  steps.push({type: 'info',    text: `【${now}】管理者パネルを開きました。`});
  steps.push({type: 'success', text: 'ログイン成功しました。'});
  steps.push({type: 'link',    text: 'グーバイクセールスアシスタント【クリックすると開きます】', href: menuUrl});
  steps.push({type: 'link',    text: 'バイク登録 | グーバイクセールスアシスタント【クリックすると開きます】', href: regUrl});
  steps.push({type: 'success', text: 'メーカー・車種情報の登録フローを確認しました。'});
  steps.push({type: 'info',    text: 'スプレッドシートの各項目の値を取得中…'});
  steps.push({type: 'info',    text: `「新・在庫台帳」シートの Y2（車種）を登録します：${modelText || '（未入力）'}`});
  steps.push({type: 'warn',    text: '必須項目が読み込めませんでした。'});

  if (missing.length > 0) {
    steps.push({type: 'error', text: `入力項目に未入力の項目があります：${missing.join('、')}`});
  } else {
    steps.push({type: 'success', text: 'すべての必須項目を読み込みました。'});
  }

  steps.push({type: 'note', text: '更新は実施しませんでした。'});

  return { startedAt: now, steps, caution: '実行中のため、再度クリックしないでください。' };
}
