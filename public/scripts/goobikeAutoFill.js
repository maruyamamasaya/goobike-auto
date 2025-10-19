(function () {
  // ===== 1) 入力受け取り（ヘッダー行＋データ行推奨。データ行のみでもOK） =====
  const raw = window.prompt(
    "スプレッドシートのヘッダー行＋データ行（2行）を貼り付けてください。\n" +
    "※データ行だけでも可（今回提示のヘッダー順にて解釈）"
  );
  if (!raw) return;

  // ===== 2) TSV/CSV を行ごとに分解 =====
  // 先頭にヘッダーがあれば header[] と data[] を作る。なければ既知順の header を採用。
  function splitLines(str) {
    return str.replace(/\r\n?/g, "\n").split("\n").filter(l => l.trim() !== "");
  }
  const lines = splitLines(raw);

  // セル分割（タブ or カンマ）。カンマが小数/金額の区切りでもだいたい大丈夫なように調整。
  function splitCells(line) {
    // まずタブ優先。なければCSV風（ダブルクオート対応の簡易版）
    if (line.includes("\t")) return line.split("\t");
    const cells = [];
    let cur = "", inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
        else inQ = !inQ;
      } else if (ch === "," && !inQ) {
        cells.push(cur); cur = "";
      } else {
        cur += ch;
      }
    }
    cells.push(cur);
    return cells;
  }

  // ===== 3) 既知のヘッダー一覧（今回いただいたものに一致） =====
  const KNOWN_HEADERS = [
    "No.","ステータス\n困った時はコチラ","メモ","管理","車種","車体番号","顧客番号","仕入\n価格","メモ","小売\n価格",
    "年式","車検","仕入日※","搬入先A\n（仕入先）","特徴・修理指示etc.","色","走行距離","売約日※","売り先","列の追加を\nしないよう\nお願いします。",
    "画像登録","画像登録URL","メーカー","排気量","メーカー","排気区分","車種","排気量","型式","系統色","色",
    "区分","初度登録年","登録年","モデル年式","年","走行距離","実数値","本体価格\n小売価格","支払総額","車台番号","修復歴",
    "メーカー保証","メーカー認定","セキュリティ","キャブ車","FI","ETC車載器付","逆輸入","フルカスタム","ノーマル車","セル付き",
    "ボアアップ社","社外マフラー","ナビ","2スト","4スト","MT","AT","ワンオーナー","社外メーター","ABS","通販可","オーディオ","品質評価書","LED／HID付","構造変更済",
    "投稿結果","記事ID","記事URL"
  ];

  let header = null, data = null;
  if (lines.length >= 2) {
    const maybeHeader = splitCells(lines[0]);
    const maybeData   = splitCells(lines[1]);
    // ヘッダーらしさ判定：既知ヘッダーのうち一致が複数あるか
    const hit = maybeHeader.filter(h => KNOWN_HEADERS.includes(h)).length;
    if (hit >= 10) {
      header = maybeHeader;
      data   = maybeData;
    }
  }
  if (!header) {
    // データ行だけが貼られた場合：既知ヘッダー順を採用
    header = KNOWN_HEADERS.slice();
    data   = splitCells(lines[0] || "");
  }

  // 空セル補完
  while (data.length < header.length) data.push("");

  // ===== 4) ユーティリティ =====
  const norm = (s) => (s ?? "").toString().trim();
  const yn   = (s) => /^(1|true|yes|y|on|はい|有|あり|TRUE)$/i.test(norm(s));
  const numDigits = (s, max) => (s || "").replace(/[^\d]/g, "").slice(0, max);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const $  = (sel, root = document) => root.querySelector(sel);
  const fire = (el, type) => el && el.dispatchEvent(new Event(type, { bubbles: true }));

  // 和暦・略号 → 西暦（ざっくり対応）
  function toYear(value) {
    const v = norm(value);
    if (!v) return "";
    // 例: "令和07年｜2025年" → 2025 を優先
    const yyyy = v.match(/(\d{4})年/);
    if (yyyy) return yyyy[1];

    // 和暦フル
    const era = [
      { re: /(令和|R)/i, base: 2018 }, // 令和元年=2019 => +2018
      { re: /(平成|H)/i, base: 1988 }, // 平成元年=1989 => +1988
      { re: /(昭和|S)/i, base: 1925 }, // 昭和元年=1926 => +1925
    ];
    // 例: "令和07年｜2025年" / "令和05年" / "R7.10.13"
    for (const e of era) {
      if (e.re.test(v)) {
        const n = v.match(/(\d{1,2})/);
        if (n) return String(e.base + parseInt(n[1], 10));
      }
    }
    // 単独の西暦数値
    const just = v.match(/(19|20)\d{2}/);
    return just ? just[0] : "";
  }

  // 値取得（ヘッダー名で拾う。無ければ ""）
  function get(hname, occurrence = 1) {
    const wanted = Math.max(occurrence, 1);
    let count = 0;
    for (let i = 0; i < header.length; i++) {
      if (norm(header[i]) === norm(hname)) {
        count++;
        if (count === wanted) return norm(data[i]);
      }
    }
    return "";
  }

  // セレクトヘルパ
  function selectByTextOrValue(selectEl, target) {
    if (!selectEl) return false;
    const t = norm(target);
    if (!t) return false;
    const opts = Array.from(selectEl.options);
    let hit = opts.find(o => norm(o.value) === t)
            || opts.find(o => norm(o.textContent) === t)
            || opts.find(o => norm(o.textContent).startsWith(t))
            || opts.find(o => norm(o.textContent).includes(t));
    if (hit) { selectEl.value = hit.value; fire(selectEl, 'change'); return true; }
    return false;
  }

  // ラジオヘルパ
  function clickRadioByNameValueOrLabel(name, wanted, root) {
    const target = norm(wanted);
    if (!target) return false;
    const radios = $$(`input[type="radio"][name="${name}"]`, root || document);
    const labelOf = (inp) => inp.closest('label')?.querySelector('.c-radio-btn__value span')?.textContent?.trim() || inp.closest('label')?.querySelector('.c-radio-btn__value')?.textContent?.trim() || "";
    let hit = radios.find(r => norm(r.value) === target)
           || radios.find(r => norm(labelOf(r)) === target)
           || radios.find(r => norm(labelOf(r)).includes(target));
    if (hit) { hit.click(); return true; }
    return false;
  }

  // チェック（IDで）
  function setCheckById(id, checked) {
    const el = document.getElementById(id);
    if (!el) return false;
    el.checked = !!checked;
    // true/false-value があれば value 同期
    const tv = el.getAttribute('true-value');
    const fv = el.getAttribute('false-value');
    if (tv !== null || fv !== null) el.value = checked ? (tv ?? el.value) : (fv ?? el.value);
    fire(el, 'change');
    return true;
  }

  // ===== 5) シート → 画面フィールドのマッピング =====

  // フォーム側：排気区分マップ（テキスト→値）
  const HAIKI_MAP = {
    "〜50cc": "60", "~50cc":"60", "～50cc":"60",
    "51〜125":"50","51-125":"50","５１ｃｃ〜１２５ｃｃ":"50",
    "126〜250":"40","１２６ｃｃ〜２５０ｃｃ":"40",
    "251〜400":"30","２５１ｃｃ〜４００ｃｃ":"30",
    "401〜750":"20","４０１ｃｃ〜７５０ｃｃ":"20",
    "751〜":"10","７５１ｃｃ〜":"10",
    "排気他":"70"
  };
  const resolveHaikiValue = (s) => {
    const n = norm(s);
    if (/^(10|20|30|40|50|60|70)$/.test(n)) return n;
    return HAIKI_MAP[n] || n;
  };

  // ===== A) メーカー / 排気区分 / 車種 / 排気量 / 型式 / 色系 =====
  const makerName      = get("メーカー", 2) || get("メーカー"); // スプレッドシート右側のメーカー列を優先
  const haikiKubun     = get("排気区分");
  const bikeName       = get("車種", 2) || get("車種");
  const haikiFree      = get("排気量", 2) || get("排気量"); // 4桁自由入力を想定（右側列優先）
  const modelFree      = get("型式");   // 15文字
  const colorTypeName  = get("系統色");
  const colorName      = get("色", 2) || get("色");

  // メーカー
  (function(){
    const makerSel = document.querySelector('select[name="maker_id"]');
    if (makerSel && makerName) {
      const ok = selectByTextOrValue(makerSel, makerName);
      if (!ok) console.warn('メーカー未検出:', makerName);
    }
  })();

  // 排気区分
  (function(){
    const haikiSel = document.querySelector('select[name="haiki_kubun_id"]');
    if (haikiSel && haikiKubun) {
      const ok = selectByTextOrValue(haikiSel, resolveHaikiValue(haikiKubun));
      if (!ok) console.warn('排気区分未検出:', haikiKubun);
    }
  })();

  // 車種
  (function(){
    const bikeSel = document.querySelector('#bike_types select');
    if (bikeSel && bikeName) setTimeout(() => {
      const ok = selectByTextOrValue(bikeSel, bikeName);
      if (!ok) console.warn('車種未検出:', bikeName);
    }, 50);
  })();

  // 排気量（#haiki_levels の自由入力4桁／セレクト「1300/不明」はシートに無ければ触らない）
  (function(){
    const root = document.getElementById('haiki_levels'); if (!root) return;
    const inp = $('input.c-input', root);
    if (inp && haikiFree) {
      const v = numDigits(haikiFree, 4);
      inp.value = v; fire(inp,'input');
      const live = root.querySelector('.c-counter span'); if (live) live.textContent = String(v.length || 0);
    }
  })();

  // 型式（自由入力）
  (function(){
    const root = document.getElementById('bike_models'); if (!root) return;
    const inp = $('#input_bike_modal_name', root);
    if (inp && modelFree) {
      let v = modelFree.slice(0, 15);
      inp.value = v; fire(inp, 'input');
      const live = root.querySelector('.c-counter span'); if (live) live.textContent = String(v.length);
    }
  })();

  // 系統色 + 色
  (function(){
    const typeRoot = document.getElementById('color_types');
    const colorRoot= document.getElementById('colors');
    if (!typeRoot || !colorRoot) return;

    const typeSel = $('select', typeRoot);
    const colorSel= $('select', colorRoot);
    const colorInp= $('input.c-input', colorRoot);

    if (typeSel && colorTypeName) selectByTextOrValue(typeSel, colorTypeName);

    setTimeout(() => {
      if (colorName) {
        if (colorSel && !colorSel.disabled) {
          const ok = selectByTextOrValue(colorSel, colorName);
          if (!ok && colorInp) { colorInp.disabled = false; colorInp.value = colorName.slice(0,30); fire(colorInp,'input'); }
        } else if (colorInp) {
          colorInp.disabled = false; colorInp.value = colorName.slice(0,30); fire(colorInp,'input');
        }
        const live = colorRoot.querySelector('.c-counter span'); if (live) live.textContent = String((colorInp?.value || "").length);
      }
    }, 120);
  })();

  // ===== B) 区分 / 初度登録年 / モデル年式 =====
  const kubunName  = get("区分");          // 中古車 / 新車（…）
  const shodoRaw   = get("初度登録年");    // 不明 / 未記入 / 国内未登録（中古） / 西暦 or 和暦
  const modelRaw   = get("モデル年式");    // 不明 / 未記入 / 西暦 or 和暦

  // 区分（#nenshiki）
  (function(){
    const root = document.getElementById('nenshiki'); if (!root || !kubunName) return;
    clickRadioByNameValueOrLabel('radio_chk1', kubunName, root);
  })();

  // 初度登録年（#nenshiki_selected_used）
  (function(){
    const root = document.getElementById('nenshiki_selected_used'); if (!root) return;
    const v = shodoRaw;
    if (!v || /選択してください/.test(v)) return;
    const ok = clickRadioByNameValueOrLabel('radio_chk1-1', v, root);
    if (!ok) {
      const yearSel = $('select[name="model_year"]', root);
      const yr = toYear(v);
      if (yr && yearSel) {
        const yearRadio = $('#model_year_select_radio input[type="radio"]', root);
        if (yearRadio) yearRadio.click();
        selectByTextOrValue(yearSel, yr);
      }
    }
  })();

  // モデル年式（#model_year）
  (function(){
    const root = document.getElementById('model_year'); if (!root) return;
    const v = modelRaw;
    if (!v || /選択してください/.test(v)) return;
    const ok = clickRadioByNameValueOrLabel('radio_chk1-2', v, root);
    if (!ok) {
      const yearSel = $('select[name="model_year"]', root);
      const yr = toYear(v);
      if (yr && yearSel) {
        const radioSelect = $('#model_year_select_radio input[type="radio"]', root);
        if (radioSelect) radioSelect.click();
        selectByTextOrValue(yearSel, yr);
      }
    }
  })();

  // ===== C) 走行距離 / ステータス =====
  const soukouStatus = get("走行距離"); // 「実走行」など（ヘッダー「走行距離」＝ステータス欄）
  const soukouValue  = get("実数値");   // 数値

  (function(){
    const root = document.getElementById('soukou'); if (!root) return;
    if (soukouStatus) clickRadioByNameValueOrLabel('radio_chk2', soukouStatus, root);
    // 実数値は「交換後」に入れる（交換前は空）
    const afterInput  = $$('input.c-input.u-txt-right', root)[1];
    const afterUnit   = $$('select[name="range_unit"]', root)[1];
    if (afterInput && soukouValue) { afterInput.value = numDigits(soukouValue, 6); fire(afterInput, 'input'); }
    if (afterUnit) selectByTextOrValue(afterUnit, "K"); // kmに寄せる（必要ならシート拡張可）
  })();

  // ===== D) 車検（今回はシート「車検」「年」などが空例だったので、入力があれば反映）=====
  const syakenStatus = get("車検"); // あり / なし / 未記入 など
  const syakenYear   = get("年");   // 西暦数値を想定（あれば）
  (function(){
    const root = document.getElementById('syaken'); if (!root) return;
    if (syakenStatus) clickRadioByNameValueOrLabel('radio_chk3', syakenStatus, root);
    const yearSel  = $('select', root);
    if (syakenYear && yearSel) { yearSel.disabled = false; selectByTextOrValue(yearSel, syakenYear); }
    // 月は今回のシートに独立列がないのでスキップ（必要なら列追加で対応）
  })();

  // ===== E) 価格 / 車台番号 / 修復歴 =====
  const priceBody = get("本体価格\n小売価格"); // 数値
  const priceTotal= get("支払総額");
  const frameNo   = get("車体番号") || get("車台番号");
  const shufuku   = get("修復歴"); // あり / なし / 1 / 0 など

  (function(){
    const root = document.getElementById('kakaku'); if (root) {
      // 金額があれば「金額入力」を選ぶ。なければ何もしない（ASKはGoobikeのみ等の制約あるため）
      if (priceBody) clickRadioByNameValueOrLabel('radio_chk4', '金額入力', root);
      const inputs = $$('input.c-input.u-txt-right', root);
      const bodyInput  = inputs[0], totalInput = inputs[1];
      if (bodyInput && priceBody)  { bodyInput.value  = numDigits(priceBody, 7);  fire(bodyInput, 'input'); }
      if (totalInput && priceTotal){ totalInput.value = numDigits(priceTotal, 7); fire(totalInput, 'input'); }
    }
    const frameRoot = document.getElementById('frame_no');
    if (frameRoot) {
      const inp = $('input.c-input', frameRoot);
      const v = norm(frameNo || "");
      if (inp && v) {
        let out = v.replace(/\s+/g, "").toUpperCase().slice(0, 18);
        inp.value = out; fire(inp, 'input');
      }
    }
    const shufukuRoot = document.getElementById('syuuhuku');
    if (shufukuRoot && shufuku) {
      clickRadioByNameValueOrLabel('radio_chk5', shufuku, shufukuRoot);
    }
  })();

  // ===== F) オプション群（TRUE/FALSE 等）=====
  // シートの見出し → 画面のcheckbox ID 対応表
  const OPT_MAP = {
    "メーカー保証" : "opt_maker_guarantee",
    "メーカー認定" : "opt_maker_official",
    "セキュリティ" : "opt_security",
    "キャブ車"     : "opt_kyabu",
    "FI"           : "opt_fi",              // true-value="2"
    "ETC車載器付"  : "opt_etc",
    "逆輸入"       : "opt_reimport",
    "フルカスタム" : "opt_fullcustom",
    "ノーマル車"   : "opt_normal",
    "セル付き"     : "opt_cell_add",
    "ボアアップ車" : "opt_boaup",
    "社外マフラー" : "opt_muffler_outside",
    "ナビ"         : "opt_navi",
    "2スト"        : "opt_2suto",
    "4スト"        : "opt_4suto",          // true-value="2"
    "MT"           : "opt_mt",
    "AT"           : "opt_at",              // true-value="2"
    "ワンオーナー" : "opt_one_owner",
    "社外メーター" : "opt_meter_outside",
    "ABS"          : "opt_abs",
    "通販可"       : "opt_tsuuhan",
    "オーディオ"   : "opt_audio",
    "品質評価書"   : "opt_quality",
    "LED／HID付"   : "opt_hid",
    "構造変更済"   : "opt_remodel_official",
  };
  (function(){
    const root = document.getElementById('options'); if (!root) return;
    Object.entries(OPT_MAP).forEach(([h, id]) => {
      const val = get(h);
      if (val === "") return; // 省略可
      setCheckById(id, yn(val));
    });
  })();

  console.log("✅ スプレッドシート→フォーム 自動入力：完了");
})();

// この版でのポイント
// ヘッダー名で参照するので、列の位置がズレてもOK（同名が複数ある場合は左側優先）。
//
// 和暦→西暦に自動変換（令和/平成/昭和、R7.10.13 のような略号も対応）。
//
// **未入力／「選択してください」**は賢くスキップ。
//
// 走行距離はシートの「走行距離（＝ステータス）」と「実数値」に分かれている前提で、実数値は「交換後」欄に投入（km）。
//
// 価格は値があれば「金額入力」を選択し、本体価格／支払総額（いずれも“万円”の整数想定）を入力。
//
// オプションは、TRUE/FALSE／1/0／はい/いいえ などを自動判定してチェック。FI/4スト/AT など true-value="2" な項目も対応済み。
//
// 車台番号は空白除去・大文字化・18文字上限。
