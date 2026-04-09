/**
 * ネポン 御見積書ウィジェット
 * メインアプリケーションロジック
 *
 * Zoho CRM Quotes モジュール用ウィジェット
 * 自動採番: CQR[連番]-[改訂番号5桁]
 */

const app = (() => {

  // ── 状態 ──────────────────────────────────────────────────────
  // ── 見積外工事マスタリスト ──────────────────────────────────────
  const EXCLUSION_MASTER = [
    'ポイラ室建屋工事',
    'ポイラ室の給排水・給排気設備工事',
    'ポイラ付近までの給排水工事',
    '機械室までの給排水工事',
    '膨張タンクへの給水配管工事',
    '燃料設備全ての工事',
    '機器設備全ての工事',
    'LPG燃料設備全ての工事',
    '油タンク設備工事',
    '油配管設備工事',
    '油配管埋設工事',
    '防油堤及びタンク基礎工事',
    '機器関係の基礎工事',
    '放熱配管材料及び取付工事',
    '配管用ピット工事',
    '埋設配管用掘削・埋め戻し工事',
    '道路横断掘小補強工事',
    '掘削時の残土処理',
    '各機器の電源制御配線工事',
    '各機器の電源制御・センサ配線工事',
    '各機器の電源制御線・制御盤取付工事',
    '各機器の制御配線工事',
    '各機器の電源制御線工事',
    '各機器の電源・制御・センサ配線工事',
    '各機器の電源・制御線・制御盤取付工事',
    '各機器の制御線工事',
    '温度制設備工事',
    '温度制御設備工事',
    '電気設備機器及び配線工事',
    'クラウド通信・遠隔制御の月額利用料',
    'クラウド通信の月額利用料',
    '遠隔制御(MC)の月額利用料',
    '遠隔制御(CGC)の月額利用料',
    '溫風暖房設備',
    '温風暖房・炭酸ガス・複合環境設備',
    'ネポンファンの取付金具及び取付工事',
    '機械周り配管工事',
    '温室内配管工事',
    '栽培ベッドへのポリ管敷設工事',
    '温風ポリダクト敷設工事',
    '放熱管敷設工事',
    '放熱管材料及び敷設工事',
    '不整地・沈下によるレール架台調整費用',
    '梱包用残材処理',
    '廃棄物処理費用',
    '作業用電気使用料',
    '作業用仮設工事',
    '躯体の孔明け補修工事',
    '暖房用熱源設備工事',
    '試運転用燃料・電気使用料',
    '試運転用燃料・水・電気使用料',
    '図面作成に関わる一切',
    '消防申請関連の一切',
    '見積記載以外の機器・設備工事',
    '消費税及び地方税',
  ];

  let state = {
    quoteId:      null,   // ZohoCRM の Quote レコードID
    quoteNumber:  null,   // Zoho自動採番 Quote_Number (整数)
    seqNo:        '',     // CQR の連番部分（ユーザー入力 or 自動採番）
    revision:     1,      // 改訂番号
    customerName: '',
    projectName:  '',
    ownerName:    '',
    date:         new Date(),
    deliveryTerm:   'お打ち合わせ願います',
    deliveryMethod: 'お打ち合わせ願います',
    paymentTerm:    'お打ち合わせ願います',
    validDays:      '見積期限は60日限りです。期限後のご用命の節は一応ご照会願います。',
    deliveryPrice:  0,
    laborCost:      null,   // null = 自動計算
    legalWelfareRate: 14.6,
    branchKey:    'honbu',
    exclusions:   [],       // 見積外工事（選択・編集済み）
    sections:     [],       // { id, no, name, items[] }
    nextSectionId: 1,
    nextItemId:    1,
    products:       [],       // 商品マスタ（Zoho Products から取得）
    koujihi:        [],       // 工事費マスタ（CustomModule1 から取得）
    searchResults:  [],       // 直近の検索結果（クリック時に参照）
    savedJson:      null,     // CRM に保存済みの見積JSON
  };

  let zohoReady = false;

  // ── 初期化 ────────────────────────────────────────────────────

  function init() {
    // 今日の日付をセット
    const today = new Date();
    document.getElementById('quoteDate').value = formatDateInput(today);

    // イベント: 数量・単価 → 金額 自動計算
    document.getElementById('sectionsContainer').addEventListener('input', onItemInput);

    document.addEventListener('click', e => {
      if (!e.target.closest('.product-search-group')) {
        document.getElementById('productDropdown').style.display = 'none';
      }
    });

    // 見積外工事チェックリスト構築
    buildExclusionUI();

    // フォント初期化
    initFont();

    // Zoho SDK 初期化
    ZOHO.embeddedApp.on('PageLoad', async (data) => {
      zohoReady = true;
      // ウィジェットポップアップサイズを 1200×800 に設定
      try { ZOHO.CRM.UI.Resize({ height: '800', width: '1200' }); } catch (_) {}
      try {
        await loadFromCRM(data);
      } catch (e) {
        console.error('CRM読み込みエラー:', e);
        showToast('CRMデータの読み込みに失敗しました', 'warn');
      } finally {
        hideLoading();
      }
    });
    ZOHO.embeddedApp.init();

    // Zoho SDK が5秒以内に応答しない場合はローディングを解除
    setTimeout(() => {
      if (!zohoReady) {
        hideLoading();
        showToast('Zoho SDK タイムアウト - デモモードで起動します', 'warn');
        loadDemoData();
      }
    }, 5000);
  }

  // ── フォント初期化 ────────────────────────────────────────────

  async function initFont() {
    const el = document.getElementById('fontStatus');
    el.textContent = '⏳ 日本語フォント読み込み中... (初回は数秒かかります)';
    el.className = 'font-status';

    const ok = await QuotationPDF.init();
    if (ok) {
      el.textContent = '✅ フォント読み込み完了 - PDF生成可能です';
      el.className = 'font-status ok';
      document.getElementById('btnGeneratePDF').disabled = false;
    } else {
      el.innerHTML = '⚠️ フォント読み込み失敗 - <a href="https://fonts.google.com/noto/specimen/Noto+Sans+JP" target="_blank">NotoSansJP-Regular.ttf</a> を widget/fonts/ に配置してください';
      el.className = 'font-status err';
      // フォントなしでも生成を許可（英数字は表示される）
      document.getElementById('btnGeneratePDF').disabled = false;
    }
  }

  // ── Zoho CRM データ読み込み ───────────────────────────────────

  async function loadFromCRM(pageData) {
    const entityId = pageData?.EntityId;
    if (!entityId) { loadDemoData(); return; }

    // EntityId は配列で返ることがある → 最初の要素を文字列として取得
    state.quoteId = Array.isArray(entityId) ? String(entityId[0]) : String(entityId);

    // Quote レコード取得
    const res = await ZOHO.CRM.API.getRecord({ Entity: 'Quotes', RecordID: entityId });
    const quote = res?.data?.[0];
    if (!quote) { loadDemoData(); return; }

    // 基本情報
    state.quoteNumber  = quote.Quote_Number || null;
    state.customerName = quote.Account_Name?.name || quote.Account_Name || '';
    state.projectName  = quote.Subject || '';
    state.ownerName    = quote.Owner?.name || '';
    state.deliveryPrice = Number(quote.Grand_Total) || 0;

    // カスタムフィールドから読み込み
    state.seqNo          = quote.field55 || '';
    state.revision       = Number(quote.field56) || 1;
    state.deliveryTerm   = quote.field6  || state.deliveryTerm;
    state.deliveryMethod = quote.field51 || state.deliveryMethod;
    state.paymentTerm    = quote.field57 || state.paymentTerm;
    state.validDays      = quote.field58 || state.validDays;

    // JSON カスタムフィールドから復元（セクション構造・全明細）
    const savedJson = quote.JSON || '';
    if (savedJson) {
      try {
        const parsed = JSON.parse(savedJson);
        state.sections      = parsed.sections     || [];
        state.deliveryPrice = parsed.deliveryPrice || state.deliveryPrice;
        state.laborCost     = parsed.laborCost     || null;
        state.exclusions    = parsed.exclusions    || [];
        renumberSections();
        // ID重複を防ぐため nextId をロード済み最大値+1 に更新
        const maxSecId  = Math.max(0, ...state.sections.map(s => s.id || 0));
        const maxItemId = Math.max(0, ...state.sections.flatMap(s => (s.items || []).map(i => i.id || 0)));
        state.nextSectionId = maxSecId  + 1;
        state.nextItemId    = maxItemId + 1;
      } catch (e) { console.warn('見積JSON解析失敗:', e); }
    }

    // フォームに反映
    applyStateToForm();

    // 商品マスタ・工事費マスタを取得
    loadProducts();
    loadKoujihi();

    showToast('CRMデータを読み込みました');
  }

  /** デモデータ（SDK未接続時） */
  function loadDemoData() {
    state.customerName  = '株式会社大仙';
    state.projectName   = '某300坪温室暖房設備工事';
    state.ownerName     = '池田';
    state.seqNo         = '166';
    state.revision      = 1;
    state.deliveryPrice = 2480000;
    state.laborCost     = 280366;

    applyStateToForm();
    showToast('デモデータで起動しました（Zoho未接続）', 'warn');
  }

  /** 状態をフォームへ反映 */
  function applyStateToForm() {
    setValue('customerName',    state.customerName);
    setValue('projectName',     state.projectName);
    setValue('ownerName',       state.ownerName);
    setValue('quoteSeqNo',      state.seqNo);
    setValue('quoteRevision',   state.revision);
    setValue('deliveryPrice',   state.deliveryPrice || '');
    setValue('laborCost',       state.laborCost || '');
    setValue('legalWelfareRate',state.legalWelfareRate);
    setValue('deliveryTerm',    state.deliveryTerm);
    setValue('deliveryMethod',  state.deliveryMethod);
    setValue('paymentTerm',     state.paymentTerm);
    setValue('validDays',       state.validDays);

    applyExclusionsToForm();
    updateQuoteNoBadge();
    renderSections();
    updateOutput();
  }

  // ── 見積外工事 ────────────────────────────────────────────────

  /** チェックリストUIを構築（init時に1回呼ぶ） */
  function buildExclusionUI() {
    const container = document.getElementById('exclusionList');
    if (!container) return;
    container.innerHTML = EXCLUSION_MASTER.map((item, i) => `
      <div class="excl-row" id="excl-row-${i}">
        <label class="excl-label">
          <input type="checkbox" class="excl-check" data-index="${i}"
                 onchange="app.onExclusionChange(this)">
          <span class="excl-text">${item}</span>
        </label>
        <input type="text" class="excl-edit" data-index="${i}"
               value="${item}" style="display:none"
               oninput="app.updateExclusionCount()">
      </div>
    `).join('');
    updateExclusionCount();
  }

  /** チェック状態変更 */
  function onExclusionChange(cb) {
    const i = cb.dataset.index;
    const editEl = document.querySelector(`.excl-edit[data-index="${i}"]`);
    if (editEl) editEl.style.display = cb.checked ? 'block' : 'none';
    updateExclusionCount();
  }

  /** カウンター更新・15件超でuncheckedを無効化 */
  function updateExclusionCount() {
    const checks = document.querySelectorAll('.excl-check');
    const checked = [...checks].filter(c => c.checked).length;
    const countEl = document.getElementById('exclusionCount');
    if (countEl) {
      countEl.textContent = `${checked}/15`;
      countEl.className = 'exclusion-count' + (checked >= 15 ? ' over' : '');
    }
    const limit = checked >= 15;
    checks.forEach(c => { if (!c.checked) c.disabled = limit; });
  }

  /** DOM から state.exclusions を収集 */
  function collectExclusions() {
    const result = [];
    document.querySelectorAll('.excl-check:checked').forEach(cb => {
      const i = cb.dataset.index;
      const editEl = document.querySelector(`.excl-edit[data-index="${i}"]`);
      result.push(editEl ? (editEl.value.trim() || EXCLUSION_MASTER[i]) : EXCLUSION_MASTER[i]);
    });
    state.exclusions = result;
  }

  /** state.exclusions をフォームのチェック状態に反映 */
  function applyExclusionsToForm() {
    if (!state.exclusions || state.exclusions.length === 0) return;
    // まず全チェック解除
    document.querySelectorAll('.excl-check').forEach(cb => {
      cb.checked = false;
      const editEl = document.querySelector(`.excl-edit[data-index="${cb.dataset.index}"]`);
      if (editEl) editEl.style.display = 'none';
    });
    state.exclusions.forEach(text => {
      // マスタから最も近いindexを探す（完全一致 → 前方一致 → 0番目）
      let idx = EXCLUSION_MASTER.indexOf(text);
      if (idx === -1) idx = EXCLUSION_MASTER.findIndex(m => text.startsWith(m.slice(0, 5)));
      if (idx === -1) {
        // マスタにない文字列 → 最初の未使用行に追加
        const freeCheck = document.querySelector('.excl-check:not(:checked)');
        if (freeCheck) {
          idx = Number(freeCheck.dataset.index);
        }
      }
      if (idx >= 0) {
        const cb = document.querySelector(`.excl-check[data-index="${idx}"]`);
        const editEl = document.querySelector(`.excl-edit[data-index="${idx}"]`);
        if (cb) { cb.checked = true; }
        if (editEl) { editEl.value = text; editEl.style.display = 'block'; }
      }
    });
    updateExclusionCount();
  }

  // ── 全件ページネーション取得 ──────────────────────────────────

  async function fetchAllRecords(entity, sortBy) {
    const all = [];
    let page = 1;
    while (true) {
      const res = await ZOHO.CRM.API.getAllRecords({
        Entity:   entity,
        sort_by:  sortBy,
        per_page: 200,
        page:     page,
      });
      const data = res?.data || [];
      all.push(...data);
      // more_records が false または取得数が 200 未満なら終了
      if (!res?.info?.more_records || data.length < 200) break;
      page++;
      if (page > 50) break;   // 最大 10,000 件で安全打ち切り
    }
    return all;
  }

  // ── 商品マスタ（Zoho Products）────────────────────────────────

  async function loadProducts() {
    if (!zohoReady) return;
    try {
      const data = await fetchAllRecords('Products', 'Product_Name');
      state.products = data.map(p => ({
        id:       p.id,
        name:     p.Product_Name  || '',
        code:     p.Product_Code  || '',
        price:    Number(p.Unit_Price) || 0,
        cost:     Number(p.Cost)       || 0,
        category: p.Product_Category  || '',
        source:   'product',
      }));
      console.log(`商品マスタ ${state.products.length} 件読み込み`);
    } catch (e) {
      console.warn('商品マスタ取得失敗:', e);
    }
  }

  // ── 工事費マスタ（CustomModule1）────────────────────────────────

  async function loadKoujihi() {
    if (!zohoReady) return;
    try {
      const data = await fetchAllRecords('CustomModule1', 'Name');
      state.koujihi = data.map(k => ({
        id:           k.id,
        name:         k.Name    || '',
        code:         k.field11 || '',   // 商品コード
        model:        k.field5  || '',   // 型式
        price:        Number(k.field6)  || 0,     // 工単価
        cost:         Number(k.field7)  || 0,     // 工原価
        category:     k.field8  || '',
        unit:         k.field9  || '式',
        order:        Number(k.field12) || 9999,  // 順番（型式内表示順）
        calcCategory: k.field13 || '',            // 算出カテゴリ
        houdan:       parseFloat(k.field14) || 0,   // 歩単（小数型）
        houkouKubun:  k.field15 || '',            // 歩工区分
        source:       'koujihi',
      }));
      console.log(`工事費マスタ ${state.koujihi.length} 件読み込み`);
      buildStandardModelSelect();
    } catch (e) {
      console.warn('工事費マスタ取得失敗:', e);
    }
  }

  /** 全角→半角正規化（英数字・記号・スペース除去） */
  function normalize(s) {
    return String(s || '')
      .replace(/[Ａ-Ｚａ-ｚ０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
      .replace(/[－ー‐−]/g, '-')
      .replace(/　/g, ' ')
      .toLowerCase();
  }

  /** 工事費マスタ検索 */
  function searchProducts(query) {
    const dd = document.getElementById('productDropdown');
    if (!query || query.length < 1) { dd.style.display = 'none'; return; }

    const q = normalize(query);

    // 商品マスタ＋工事費マスタを統合検索
    const fromProducts = state.products.filter(p =>
      normalize(p.name).includes(q) ||
      normalize(p.code).includes(q)
    ).slice(0, 10).map(p => ({ ...p, source: 'product' }));

    const fromKoujihi = state.koujihi.filter(k =>
      normalize(k.name).includes(q) ||
      normalize(k.code).includes(q) ||
      normalize(k.model).includes(q) ||
      normalize(k.category).includes(q)
    ).slice(0, 15);

    const filtered = [...fromProducts, ...fromKoujihi].slice(0, 25);
    if (filtered.length === 0) { dd.style.display = 'none'; return; }

    state.searchResults = filtered;

    dd.innerHTML = filtered.map((item, idx) => {
      const isKoujihi = item.source === 'koujihi';
      const badge = isKoujihi
        ? '<span class="p-badge koujihi">工事費</span>'
        : '<span class="p-badge product">商品</span>';
      const sub = isKoujihi
        ? escHtml([item.model, item.category].filter(Boolean).join(' / '))
        : escHtml(item.code);
      return `
        <div class="product-item" data-idx="${idx}">
          <div style="flex:1;min-width:0">
            <div class="p-name">${badge}${escHtml(item.name)}</div>
            <div class="p-code">${sub}</div>
          </div>
          <div class="p-price">¥${item.price.toLocaleString('ja-JP')}</div>
        </div>`;
    }).join('');

    // innerHTML 設定後に各行へ直接リスナーを付ける
    dd.querySelectorAll('.product-item').forEach(el => {
      el.addEventListener('click', e => {
        e.stopPropagation();
        const product = state.searchResults[Number(el.dataset.idx)];
        console.log('[dropdown click] idx:', el.dataset.idx, 'product:', product?.name);
        if (product) selectProduct(product);
      });
    });

    dd.style.display = 'block';
  }

  /** 工事費1件をセクションに追加する共通処理 */
  function addKoujihiItem(section, k) {
    const item = createItem();
    item.productId    = null;
    item.name         = k.name;
    item.spec         = k.model || '';
    item.unit         = k.unit  || '式';
    item.unitPrice    = k.price;
    item.amount       = k.price;
    item.calcCategory = k.calcCategory || '';
    item.houdan       = k.houdan      || 0;
    item.houkouKubun  = k.houkouKubun || '';
    section.items.push(item);
  }

  /** 標準項: 型式の一覧をセレクトボックスに構築 */
  function buildStandardModelSelect() {
    const sel  = document.getElementById('standardModelSelect');
    const hint = document.getElementById('standardHint');
    if (!sel) return;

    // 型式が存在するレコードから重複なしリストを作成
    // （型式ごとにorderが最小のものを代表として並び替え）
    const modelMap = new Map(); // model -> min order
    state.koujihi.forEach(k => {
      if (!k.model || !k.model.trim()) return;
      const m = k.model.trim();
      const cur = modelMap.get(m);
      if (cur === undefined || k.order < cur) modelMap.set(m, k.order);
    });

    // 型式名をアルファベット順に並べる
    const models = [...modelMap.keys()].sort((a, b) => a.localeCompare(b, 'ja'));

    // セレクトを再構築（最初のオプションは残す）
    while (sel.options.length > 1) sel.remove(1);
    models.forEach(m => {
      const count = state.koujihi.filter(k => k.model && k.model.trim() === m).length;
      const opt = document.createElement('option');
      opt.value = m;
      opt.textContent = `${m}（${count}件）`;
      sel.appendChild(opt);
    });

    if (hint) hint.textContent = `${models.length} 型式`;
    console.log(`標準項: ${models.length} 型式を構築`);
  }

  /** ③追加ボタン: 選択中の型式を指定セクションへ追加 */
  function execStandardAdd() {
    const model = document.getElementById('standardModelSelect')?.value || '';
    if (!model) {
      showToast('① 型式を選択してください', 'warn');
      return;
    }

    // セクションが無ければ追加
    if (state.sections.length === 0) addSection();

    // 追加先セクションを取得
    const targetVal = document.getElementById('standardTargetSection')?.value || 'last';
    let targetSection;
    if (targetVal === 'last') {
      targetSection = state.sections[state.sections.length - 1];
    } else {
      const targetId = Number(targetVal);
      targetSection = state.sections.find(s => s.id === targetId) || state.sections[state.sections.length - 1];
    }

    // 末尾の空行があれば削除
    const lastItem = targetSection.items[targetSection.items.length - 1];
    if (lastItem && !lastItem.name && !lastItem.spec && !lastItem.unitPrice && !lastItem.amount) {
      targetSection.items.pop();
    }

    // 型式に一致するものを field12（順番）でソートして一括追加
    const matched = state.koujihi
      .filter(k => k.model && k.model.trim() === model.trim())
      .sort((a, b) => a.order - b.order);

    if (matched.length === 0) {
      showToast(`型式「${model}」の工事費が見つかりません`, 'warn');
      return;
    }

    matched.forEach(k => addKoujihiItem(targetSection, k));
    renderSections();
    updateOutput();
    showToast(`No.${targetSection.no}「${targetSection.name || '無題'}」に ${matched.length} 件追加しました`);
  }

  /** 追加先セクションセレクトを更新（セクション追加・削除時に呼ぶ） */
  function updateTargetSectionSelect() {
    const sel = document.getElementById('standardTargetSection');
    if (!sel) return;
    const cur = sel.value;
    while (sel.options.length > 0) sel.remove(0);
    state.sections.forEach(s => {
      const opt = document.createElement('option');
      opt.value = s.id;
      opt.textContent = `No.${s.no}`;
      sel.appendChild(opt);
    });
    // 末尾オプションを追加
    const lastOpt = document.createElement('option');
    lastOpt.value = 'last';
    lastOpt.textContent = '最後';
    sel.appendChild(lastOpt);
    // 以前の選択を維持（初期値 'last' の場合は No.1 に切り替え）
    const isInitial = cur === 'last' && state.sections.length > 0;
    if (!isInitial && [...sel.options].some(o => o.value === cur)) {
      sel.value = cur;
    } else {
      sel.value = state.sections.length > 0 ? String(state.sections[0].id) : 'last';
    }
  }

  /** 検索結果から選択 */
  function selectProduct(product) {
    document.getElementById('productDropdown').style.display = 'none';
    document.getElementById('productSearch').value = '';

    if (!product) return;

    // セクションが無ければ追加
    if (state.sections.length === 0) addSection();
    const lastSection = state.sections[state.sections.length - 1];

    // 末尾の空行があれば削除
    const lastItem = lastSection.items[lastSection.items.length - 1];
    if (lastItem && !lastItem.name && !lastItem.spec && !lastItem.unitPrice && !lastItem.amount) {
      lastSection.items.pop();
    }

    if (product.source === 'product') {
      // ── 商品マスタ: 商品1行 + 同型式の工事費を追加 ──────────────
      const item = createItem();
      item.productId = product.id;
      item.name      = product.name;
      item.spec      = product.code;
      item.unit      = '式';
      item.unitPrice = product.price;
      item.amount    = product.price;
      lastSection.items.push(item);

      if (product.code) {
        const matched = state.koujihi
          .filter(k => k.model && normalize(k.model) === normalize(product.code))
          .sort((a, b) => a.order - b.order);
        matched.forEach(k => addKoujihiItem(lastSection, k));
        if (matched.length > 0) showToast(`${product.name} と工事費 ${matched.length} 件を追加しました`);
        else showToast(`${product.name} を追加しました`);
      }
    } else {
      // ── 工事費マスタ: 同型式を field12順で全件追加 ────────────────
      const model   = product.model ? normalize(product.model) : '';
      const matched = model
        ? state.koujihi
            .filter(k => k.model && normalize(k.model) === model)
            .sort((a, b) => a.order - b.order)
        : [product];

      matched.forEach(k => addKoujihiItem(lastSection, k));
      showToast(`型式 ${product.model || product.name} の工事費 ${matched.length} 件を追加しました`);
    }

    renderSections();
    updateOutput();
  }

  // ── 自動採番 ─────────────────────────────────────────────────

  async function autoNumber() {
    const btn = document.getElementById('btnAutoNumber');
    btn.disabled = true;
    btn.textContent = '採番中...';

    try {
      let maxSeq = 0;

      if (zohoReady) {
        // 最近200件の見積を取得し field_seq_no の最大値を探す
        // ※ Quote_Number は Zoho 内部 ID（18桁）のため使用しない
        const res = await ZOHO.CRM.API.getAllRecords({
          Entity:     'Quotes',
          sort_by:    'Created_Time',
          sort_order: 'desc',
          per_page:   200,
        });
        const list = res?.data || [];

        list.forEach(q => {
          const seq = Number(q.field55) || 0;
          if (seq > maxSeq) maxSeq = seq;
        });

        // field55 が全レコード未設定なら件数ベースで採番
        if (maxSeq === 0) maxSeq = list.length;
      }

      state.seqNo    = maxSeq + 1;
      state.revision = 1;
      setValue('quoteSeqNo',    state.seqNo);
      setValue('quoteRevision', state.revision);
      updateQuoteNoBadge();
      showToast(`採番完了: CQR${state.seqNo}-00001`);
    } catch (e) {
      const msg = e?.message || JSON.stringify(e);
      showToast('採番に失敗しました: ' + msg, 'err');
    } finally {
      btn.disabled = false;
      btn.textContent = '🔢 自動採番';
    }
  }

  // ── セクション操作 ────────────────────────────────────────────

  function addSection() {
    const section = {
      id:    state.nextSectionId++,
      no:    state.sections.length + 1,
      name:  '',
      items: [],
    };
    // 最初の行を1つ追加
    section.items.push(createItem());
    state.sections.push(section);
    renderSections();
    updateOutput();
    document.getElementById('noSectionsMsg').style.display = 'none';
  }

  function removeSection(btn) {
    const block = btn.closest('.section-block');
    const id    = Number(block.dataset.sectionId);
    if (!confirm('このセクションを削除しますか？')) return;
    state.sections = state.sections.filter(s => s.id !== id);
    renumberSections();
    renderSections();
    updateOutput();
  }

  function moveSectionUp(btn) {
    const block = btn.closest('.section-block');
    const id    = Number(block.dataset.sectionId);
    const idx   = state.sections.findIndex(s => s.id === id);
    if (idx <= 0) return;
    [state.sections[idx - 1], state.sections[idx]] = [state.sections[idx], state.sections[idx - 1]];
    renumberSections();
    renderSections();
  }

  function moveSectionDown(btn) {
    const block = btn.closest('.section-block');
    const id    = Number(block.dataset.sectionId);
    const idx   = state.sections.findIndex(s => s.id === id);
    if (idx < 0 || idx >= state.sections.length - 1) return;
    [state.sections[idx], state.sections[idx + 1]] = [state.sections[idx + 1], state.sections[idx]];
    renumberSections();
    renderSections();
  }

  function renumberSections() {
    state.sections.forEach((s, i) => { s.no = i + 1; });
  }

  // ── 明細行操作 ────────────────────────────────────────────────

  function createItem() {
    return {
      id: state.nextItemId++, productId: null,
      name: '', spec: '', qty: 1, unit: '式', unitPrice: null, amount: 0,
      calcCategory: '',
      houdan: 0,        // 歩単（マスタから）
      houkouKubun: '',  // 歩工区分（マスタから）
      houkouGoukei: 0,  // 歩工合計（手動上書き、0=自動計算）
    };
  }

  function addItem(btn) {
    const block = btn.closest('.section-block');
    const id    = Number(block.dataset.sectionId);
    const sec   = state.sections.find(s => s.id === id);
    if (sec) {
      sec.items.push(createItem());
      renderSection(sec, block);
      updateSectionSubtotal(block);
    }
  }

  function removeItem(btn) {
    const row   = btn.closest('.item-row');
    const itemId = Number(row.dataset.itemId);
    const block = btn.closest('.section-block');
    const secId = Number(block.dataset.sectionId);
    const sec   = state.sections.find(s => s.id === secId);
    if (sec) {
      sec.items = sec.items.filter(i => i.id !== itemId);
      row.remove();
      updateSectionSubtotal(block);
      updateOutput();
    }
  }

  /** 数量・単価変更 → 金額自動計算 */
  function onItemInput(e) {
    const row = e.target.closest('.item-row');
    if (!row) return;

    const block = e.target.closest('.section-block');
    const secId = Number(block.dataset.sectionId);
    const itemId = Number(row.dataset.itemId);
    const sec   = state.sections.find(s => s.id === secId);
    if (!sec) return;
    const item = sec.items.find(i => i.id === itemId);
    if (!item) return;

    // フォームから状態を更新
    const nameEl   = row.querySelector('.item-name');
    const specEl   = row.querySelector('.item-spec');
    const qtyEl    = row.querySelector('.item-qty');
    const unitEl   = row.querySelector('.item-unit');
    const priceEl  = row.querySelector('.item-price');
    const amountEl = row.querySelector('.item-amount');

    item.name      = nameEl?.value   || '';
    item.spec      = specEl?.value   || '';
    item.qty       = Number(qtyEl?.value)   || 0;
    item.unit      = unitEl?.value   || '式';
    item.unitPrice = priceEl?.value  ? Number(priceEl.value) : null;

    // 金額自動計算（数量×単価 が入力されている場合）
    if (item.unitPrice !== null && item.qty) {
      item.amount = item.unitPrice * item.qty;
      if (amountEl) amountEl.value = item.amount;
    } else if (e.target === amountEl) {
      item.amount = Number(amountEl?.value) || 0;
    }

    // 歩工合計（手動上書き）を読み取り、歩工を再計算して表示
    const goukeiEl = row.querySelector('.item-houkou-goukei');
    const houkouEl = row.querySelector('.item-houkou');
    if (goukeiEl) item.houkouGoukei = parseFloat(goukeiEl.value) || 0;
    if (houkouEl) {
      const raw = item.houkouGoukei > 0
        ? item.houkouGoukei
        : (item.qty * (item.houdan || 0));
      houkouEl.textContent = raw ? parseFloat(raw.toFixed(4)) : '';
    }

    updateSectionSubtotal(block);
    updateOutput();
  }

  // ── レンダリング ──────────────────────────────────────────────

  function renderSections() {
    const container = document.getElementById('sectionsContainer');
    const noMsg     = document.getElementById('noSectionsMsg');

    updateTargetSectionSelect();

    if (state.sections.length === 0) {
      container.innerHTML = '';
      noMsg.style.display = 'block';
      return;
    }
    noMsg.style.display = 'none';

    // 既存DOMを再利用（IDで照合）
    const existingIds = new Set([...container.querySelectorAll('.section-block')].map(b => Number(b.dataset.sectionId)));
    const newIds      = new Set(state.sections.map(s => s.id));

    // 削除
    existingIds.forEach(id => {
      if (!newIds.has(id)) container.querySelector(`[data-section-id="${id}"]`)?.remove();
    });

    // 追加・更新（順序を保つ）
    state.sections.forEach(sec => {
      let block = container.querySelector(`[data-section-id="${sec.id}"]`);
      if (!block) {
        block = createSectionDOM(sec);
        container.appendChild(block);
      }
      // 番号とタイトル更新
      block.querySelector('.section-no').textContent = sec.no;
      const nameInput = block.querySelector('.section-name-input');
      if (nameInput && nameInput !== document.activeElement) nameInput.value = sec.name;

      renderSection(sec, block);
    });

    // 順序修正
    state.sections.forEach((sec, idx) => {
      const block = container.querySelector(`[data-section-id="${sec.id}"]`);
      if (block && container.children[idx] !== block) {
        container.insertBefore(block, container.children[idx]);
      }
    });
  }

  function createSectionDOM(sec) {
    const tmpl  = document.getElementById('sectionTemplate');
    const clone = tmpl.content.cloneNode(true);
    const block = clone.querySelector('.section-block');
    block.dataset.sectionId = sec.id;

    // セクション名変更 → 状態更新
    block.querySelector('.section-name-input').addEventListener('change', e => {
      const s = state.sections.find(s => s.id === sec.id);
      if (s) s.name = e.target.value;
    });

    return block;
  }

  function renderSection(sec, block) {
    const tbody = block.querySelector('.items-tbody');
    const existingIds = new Set([...tbody.querySelectorAll('.item-row')].map(r => Number(r.dataset.itemId)));
    const newIds      = new Set(sec.items.map(i => i.id));

    // 削除
    existingIds.forEach(id => {
      if (!newIds.has(id)) tbody.querySelector(`[data-item-id="${id}"]`)?.remove();
    });

    // 追加・更新
    sec.items.forEach(item => {
      let row = tbody.querySelector(`[data-item-id="${item.id}"]`);
      if (!row) {
        row = createItemRowDOM(item);
        tbody.appendChild(row);
      }
      // フォーカス中の行は上書きしない
      if (![...row.querySelectorAll('input,select')].some(el => el === document.activeElement)) {
        row.querySelector('.item-name').value   = item.name;
        row.querySelector('.item-spec').value   = item.spec;
        row.querySelector('.item-qty').value    = item.qty;
        row.querySelector('.item-unit').value   = item.unit;
        row.querySelector('.item-price').value  = item.unitPrice ?? '';
        row.querySelector('.item-amount').value = item.amount || '';
      }
      // 算出カテゴリバッジを常に更新
      const badge = row.querySelector('.calc-cat-badge');
      if (badge) {
        if (item.calcCategory) {
          badge.textContent = item.calcCategory;
          badge.style.display = '';
        } else {
          badge.style.display = 'none';
        }
      }
      // data属性にも保持（保存・読み込み時に利用）
      row.dataset.calcCategory = item.calcCategory || '';

      // 歩単・歩工区分・歩工合計・歩工（計算）の反映
      const houdanEl      = row.querySelector('.item-houdan');
      const kubunEl       = row.querySelector('.item-houkou-kubun');
      const goukeiEl      = row.querySelector('.item-houkou-goukei');
      const houkouDispEl  = row.querySelector('.item-houkou');
      // 歩単：小数型のため parseFloat で保持し最大4桁で表示
      if (houdanEl) houdanEl.textContent = item.houdan != null && item.houdan !== 0
        ? parseFloat(item.houdan.toFixed(4))
        : '';
      if (kubunEl)  kubunEl.textContent  = item.houkouKubun || '';
      if (goukeiEl && goukeiEl !== document.activeElement) {
        goukeiEl.value = item.houkouGoukei || '';
      }
      if (houkouDispEl) {
        const raw = (item.houkouGoukei > 0)
          ? item.houkouGoukei
          : (item.qty * (item.houdan || 0));
        // 小数演算の誤差を丸めて表示
        const houkou = raw ? parseFloat(raw.toFixed(4)) : '';
        houkouDispEl.textContent = houkou;
      }
    });

    // 順序修正
    sec.items.forEach((item, idx) => {
      const row = tbody.querySelector(`[data-item-id="${item.id}"]`);
      if (row && tbody.children[idx] !== row) tbody.insertBefore(row, tbody.children[idx]);
    });

    updateSectionSubtotal(block);
  }

  function createItemRowDOM(item) {
    const tmpl  = document.getElementById('itemRowTemplate');
    const clone = tmpl.content.cloneNode(true);
    const row   = clone.querySelector('.item-row');
    row.dataset.itemId = item.id;
    return row;
  }

  function updateSectionSubtotal(block) {
    const secId  = Number(block.dataset.sectionId);
    const sec    = state.sections.find(s => s.id === secId);
    if (!sec) return;
    const subtotal = sec.items.reduce((sum, i) => sum + (Number(i.amount) || 0), 0);
    const el = block.querySelector('.subtotal-val');
    if (el) el.textContent = '¥' + subtotal.toLocaleString('ja-JP');
    updateOutput();
  }

  // ── PDF出力タブ 更新 ──────────────────────────────────────────

  function updateOutput() {
    readFormToState();

    const sections     = state.sections;
    const grandTotal   = sections.reduce((sum, s) =>
      sum + s.items.reduce((ss, i) => ss + (Number(i.amount) || 0), 0), 0);
    const deliveryPrice = Number(getValue('deliveryPrice')) || grandTotal;
    const legalRate    = (Number(getValue('legalWelfareRate')) || 14.6) / 100;
    const laborCost    = Number(getValue('laborCost')) || Math.round(deliveryPrice * 0.115);
    const legalWelfare = Math.round(laborCost * legalRate);
    const materialCost = deliveryPrice - laborCost - legalWelfare;

    const seqNo        = getValue('quoteSeqNo');
    const revision     = getValue('quoteRevision') || 1;
    const quoteNoStr   = seqNo ? `CQR${seqNo}-${String(revision).padStart(5, '0')}` : '（未採番）';

    setText('sum-quoteNo',    quoteNoStr);
    setText('sum-date',       formatDisplayDate(new Date(getValue('quoteDate') || Date.now())));
    setText('sum-customer',   getValue('customerName') || '-');
    setText('sum-project',    getValue('projectName')  || '-');
    setText('sum-sections',   sections.length);
    setText('sum-total',      '¥' + grandTotal.toLocaleString('ja-JP'));
    setText('sum-delivery',   '¥' + deliveryPrice.toLocaleString('ja-JP'));
    setText('sum-material',   '¥' + Math.max(0, materialCost).toLocaleString('ja-JP'));
    setText('sum-labor',      '¥' + laborCost.toLocaleString('ja-JP'));
    setText('sum-welfare',    '¥' + legalWelfare.toLocaleString('ja-JP'));
    setText('sum-welfare-label', `　3) 法定福利費(${(legalRate * 100).toFixed(1)}%)`);
    setText('grandTotalDisplay', '¥' + grandTotal.toLocaleString('ja-JP'));

    updateQuoteNoBadge();
  }

  function updateQuoteNoBadge() {
    const seqNo   = getValue('quoteSeqNo');
    const revision = getValue('quoteRevision') || 1;
    const badge   = document.getElementById('quoteNoBadge');
    badge.textContent = seqNo
      ? `CQR${seqNo}-${String(revision).padStart(5, '0')}`
      : '採番待ち';
  }

  /** フォーム値を state へ読み込む */
  function readFormToState() {
    state.seqNo          = getValue('quoteSeqNo');
    state.revision       = Number(getValue('quoteRevision'))   || 1;
    state.customerName   = getValue('customerName');
    state.projectName    = getValue('projectName');
    state.ownerName      = getValue('ownerName');
    state.deliveryTerm   = getValue('deliveryTerm');
    state.deliveryMethod = getValue('deliveryMethod');
    state.paymentTerm    = getValue('paymentTerm');
    state.validDays      = getValue('validDays') || '';
    state.deliveryPrice  = Number(getValue('deliveryPrice')) || 0;
    state.laborCost      = getValue('laborCost') ? Number(getValue('laborCost')) : null;
    state.legalWelfareRate = Number(getValue('legalWelfareRate')) || 14.6;
    state.branchKey      = getValue('branchSelect');
    const dateVal = getValue('quoteDate');
    state.date = dateVal ? new Date(dateVal) : new Date();
  }

  // ── PDF 生成 ─────────────────────────────────────────────────

  async function generatePDF() {
    readFormToState();

    if (!state.seqNo) {
      if (!confirm('見積番号が未設定です。このまま生成しますか？')) return;
    }

    const btn = document.getElementById('btnGeneratePDF');
    btn.disabled = true;
    btn.textContent = '⏳ 生成中...';
    const statusEl = document.getElementById('pdfStatus');
    statusEl.textContent = '';

    try {
      const data = buildPdfData();
      QuotationPDF.download(data);
      statusEl.textContent = '✅ PDFをダウンロードしました';
      showToast('PDF生成完了');
    } catch (e) {
      console.error('PDF生成エラー:', e);
      statusEl.textContent = '❌ PDF生成失敗: ' + e.message;
      showToast('PDF生成に失敗しました', 'err');
    } finally {
      btn.disabled = false;
      btn.textContent = '📄 PDF生成・ダウンロード';
    }
  }

  /** PDF 生成用データオブジェクトを組み立てる */
  function buildPdfData() {
    collectExclusions();
    return {
      seqNo:           state.seqNo,
      revision:        state.revision,
      date:            state.date,
      customerName:    state.customerName,
      projectName:     state.projectName,
      ownerName:       state.ownerName,
      deliveryTerm:    state.deliveryTerm,
      deliveryMethod:  state.deliveryMethod,
      paymentTerm:     state.paymentTerm,
      validDays:       state.validDays,
      deliveryPrice:   state.deliveryPrice,
      laborCost:       state.laborCost,
      legalWelfareRate: state.legalWelfareRate,
      branchKey:       state.branchKey,
      sections:        state.sections,
      exclusions:      state.exclusions.length > 0 ? state.exclusions : undefined,
    };
  }

  // ── CRM 保存 ─────────────────────────────────────────────────

  async function saveToCRM() {
    const statusEl = document.getElementById('saveStatus');
    if (!zohoReady || !state.quoteId) {
      statusEl.textContent = '⚠️ ZohoCRM に接続されていません（zohoReady=' + zohoReady + ', quoteId=' + state.quoteId + '）';
      showToast('ZohoCRM に接続されていません', 'warn');
      return;
    }
    const btn = document.getElementById('btnSaveToCRM');
    btn.disabled = true;
    statusEl.textContent = '⏳ 保存中...';

    try {
      readFormToState();
      console.log('保存開始 quoteId:', state.quoteId, 'sections:', state.sections.length);

      collectExclusions();
      const jsonStr = JSON.stringify({
        sections:      state.sections,
        deliveryPrice: state.deliveryPrice,
        laborCost:     state.laborCost,
        seqNo:         state.seqNo,
        revision:      state.revision,
        exclusions:    state.exclusions,
      });

      const apiData = {
        id:      state.quoteId,
        JSON:    jsonStr,
        field55: state.seqNo      ? String(state.seqNo)      : undefined,
        field56: state.revision   ? Number(state.revision)   : undefined,
        field6:  state.deliveryTerm,
        field51: state.deliveryMethod,
        field57: state.paymentTerm,
        field58: state.validDays,
      };
      // undefined のキーを除去
      Object.keys(apiData).forEach(k => { if (apiData[k] === undefined) delete apiData[k]; });

      console.log('【保存内容】', {
        quoteId:        state.quoteId,
        seqNo:          apiData.field55,
        revision:       apiData.field56,
        deliveryTerm:   apiData.field6,
        deliveryMethod: apiData.field51,
        paymentTerm:    apiData.field57,
        validDays:      apiData.field58,
        jsonLen:        jsonStr.length,
      });

      const updateRes = await ZOHO.CRM.API.updateRecord({
        Entity:  'Quotes',
        APIData: apiData,
        Trigger: [],
      });

      console.log('updateRecord response:', JSON.stringify(updateRes));

      // Zoho API のレスポンス確認
      const resData = updateRes?.data?.[0];
      if (resData?.code === 'SUCCESS') {
        statusEl.textContent = '✅ 保存しました（' + new Date().toLocaleTimeString('ja-JP') + '）';
        showToast('CRMに保存しました');
      } else {
        const apiName = resData?.details?.api_name || '';
        const msg = apiName
          ? `フィールド "${apiName}" が見つかりません。CRMにカスタムフィールドを作成してください。`
          : (resData?.message || JSON.stringify(resData));
        statusEl.textContent = '⚠️ ' + msg;
        showToast('保存に問題があります', 'warn');
        console.warn('CRM保存レスポンス:', JSON.stringify(updateRes));
      }
    } catch (e) {
      const msg = e?.message || String(e) || 'Unknown error';
      statusEl.textContent = '❌ 保存失敗: ' + msg;
      showToast('保存失敗: ' + msg, 'err');
      console.error('saveToCRM error:', e);
    } finally {
      btn.disabled = false;
    }
  }

  // ── タブ切り替え ──────────────────────────────────────────────

  function goToTab(tabName) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
    document.getElementById(`tab-${tabName}`).classList.add('active');
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

    if (tabName === 'output') updateOutput();
  }

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => goToTab(btn.dataset.tab));
  });

  // ── 営業所切り替え ────────────────────────────────────────────

  function onBranchChange() {
    const val = getValue('branchSelect');
    document.getElementById('branchCustomInput').style.display = val === 'other' ? 'block' : 'none';
  }

  // ── ユーティリティ ────────────────────────────────────────────

  function getValue(id) {
    const el = document.getElementById(id);
    return el ? el.value : '';
  }

  function setValue(id, val) {
    const el = document.getElementById(id);
    if (el && val !== undefined && val !== null) el.value = val;
  }

  function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
  }

  function escHtml(str) {
    return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  function formatDateInput(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  function formatDisplayDate(date) {
    if (isNaN(date)) return '-';
    return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;
  }

  function hideLoading() {
    const el = document.getElementById('loadingOverlay');
    if (el) el.style.display = 'none';
  }

  function showToast(msg, type = 'info') {
    const colors = { info: '#1a4d8f', warn: '#856404', err: '#c0392b' };
    const div = document.createElement('div');
    div.textContent = msg;
    Object.assign(div.style, {
      position: 'fixed', bottom: '70px', right: '16px',
      background: colors[type] || '#333', color: '#fff',
      padding: '8px 14px', borderRadius: '6px',
      fontSize: '12px', zIndex: '200',
      boxShadow: '0 2px 8px rgba(0,0,0,0.3)',
      opacity: '1', transition: 'opacity 0.5s',
    });
    document.body.appendChild(div);
    setTimeout(() => { div.style.opacity = '0'; }, 2500);
    setTimeout(() => div.remove(), 3100);
  }

  // ── 公開API ───────────────────────────────────────────────────

  return {
    init,
    // タブ
    goToTab,
    // セクション
    addSection,
    removeSection,
    moveSectionUp,
    moveSectionDown,
    // 明細行
    addItem,
    removeItem,
    // 商品検索
    searchProducts,
    selectProduct,
    // 採番
    autoNumber,
    // 営業所
    onBranchChange,
    // 標準項
    execStandardAdd,
    // 見積外工事
    onExclusionChange,
    updateExclusionCount,
    // PDF
    generatePDF,
    // CRM保存
    saveToCRM,
  };

})();

// ── 起動 ─────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  window.app = app;
  document.getElementById('btnAutoNumber').addEventListener('click', () => app.autoNumber());
  app.init();
});
