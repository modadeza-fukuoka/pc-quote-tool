// ========================================
// PC Quote Tool
// Live spreadsheet + compatibility filter + AI suggest + multi extras
// ========================================

document.addEventListener('DOMContentLoaded', async () => {
  let pcType = 'desktop';   // 'desktop' or 'notebook'
  let desktopSub = 'gaming'; // 'gaming' or 'slim'
  let extraIdCounter = 0;
  // (slim model is auto-detected from part selections)

  const btnDesktop    = document.getElementById('btn-desktop');
  const btnNotebook   = document.getElementById('btn-notebook');
  const btnGaming     = document.getElementById('btn-gaming');
  const btnSlim       = document.getElementById('btn-slim');
  const step3Desktop  = document.getElementById('step3-desktop');
  const step3Slim     = document.getElementById('step3-slim');
  const step3Notebook = document.getElementById('step3-notebook');
  const desktopSubEl  = document.getElementById('desktop-subtype');
  const quotePaper    = document.getElementById('quote-paper');
  const partsList     = document.getElementById('parts-list');
  const loadingEl     = document.getElementById('loading-indicator');

  const osOptions = [
    { name: "Windows11 Home", price: 15000 },
    { name: "Windows11 Pro", price: 22000 },
    { name: "不要", price: 0 },
  ];

  // Main parts (extra is handled separately now)
  const partDefs = [
    { key: 'cpu',         badge: 'CPU',    quoteLabel: 'CPU' },
    { key: 'cooler',      badge: 'クーラー', quoteLabel: 'CPUクーラー' },
    { key: 'motherboard', badge: 'M/B',    quoteLabel: 'マザーボード' },
    { key: 'memory',      badge: 'メモリ',  quoteLabel: 'メモリ' },
    { key: 'gpu',         badge: 'GPU',    quoteLabel: 'GPU' },
    { key: 'ssd',         badge: 'SSD',    quoteLabel: 'SSD' },
    { key: 'case',        badge: 'ケース',  quoteLabel: 'ケース' },
    { key: 'caseFan',     badge: 'ファン',  quoteLabel: 'ケースファン' },
    { key: 'psu',         badge: '電源',   quoteLabel: '電源' },
    { key: 'os',          badge: 'OS',     quoteLabel: 'OS' },
  ];

  // --- Fetch live data ---
  if (loadingEl) loadingEl.classList.remove('hidden');
  const fetched = await fetchMasterData();
  await fetchSlimData();
  await fetchRecommendations();
  if (loadingEl) loadingEl.classList.add('hidden');
  if (fetched) {
    document.getElementById('data-status').textContent = 'スプレッドシートから取得済み';
    document.getElementById('data-status').classList.add('live');
  }

  // --- Compatibility filtering ---
  function getItemsForKey(key) {
    if (key === 'os') return osOptions;
    if (key === 'motherboard') {
      const cpuInfo = getPartInfo(partDefs.find(d => d.key === 'cpu'));
      if (cpuInfo.name) return filterMotherboards(cpuInfo.name);
    }
    if (key === 'memory') {
      const mbInfo = getPartInfo(partDefs.find(d => d.key === 'motherboard'));
      if (mbInfo.name) return filterMemory(mbInfo.name);
    }
    return MASTER[key] || [];
  }

  function buildOptions(key) {
    const items = getItemsForKey(key);
    let html = '<option value="">-- 選択 --</option>';
    items.forEach(item => {
      html += `<option value="${item.name}" data-price="${item.price}">${item.name}</option>`;
    });
    html += '<option value="__custom__">手動入力</option>';
    return html;
  }

  function buildExtraOptions() {
    const items = MASTER.extra || [];
    let html = '<option value="">-- 選択 --</option>';
    items.forEach(item => {
      html += `<option value="${item.name}" data-price="${item.price}">${item.name}</option>`;
    });
    html += '<option value="__custom__">手動入力</option>';
    return html;
  }

  // --- Build main part rows ---
  function buildPartRows() {
    partsList.innerHTML = '';
    partDefs.forEach(def => {
      const row = document.createElement('div');
      row.className = 'part-row';
      row.innerHTML = `
        <div class="part-badge">${def.badge}</div>
        <div class="part-select-wrap">
          <select id="${def.key}-select" class="input select part-select">${buildOptions(def.key)}</select>
          <div class="part-custom hidden" id="${def.key}-custom-wrap">
            <input type="text" id="${def.key}-name" class="input" placeholder="パーツ名">
            <input type="number" id="${def.key}-price" class="input input-price" placeholder="¥ 金額">
          </div>
        </div>
        <div class="part-price-display" id="${def.key}-price-display"></div>
      `;
      partsList.appendChild(row);

      const select = row.querySelector('select');
      select.addEventListener('change', () => {
        const val = select.value;
        const cw = row.querySelector('.part-custom');
        const pd = row.querySelector('.part-price-display');
        if (val === '__custom__') { cw.classList.remove('hidden'); pd.textContent = ''; }
        else if (val) { cw.classList.add('hidden'); pd.textContent = '¥' + (parseInt(select.selectedOptions[0]?.dataset.price) || 0).toLocaleString(); }
        else { cw.classList.add('hidden'); pd.textContent = ''; }
        if (def.key === 'cpu') { refreshSelect('motherboard'); refreshSelect('memory'); }
        if (def.key === 'motherboard') { refreshSelect('memory'); updateAutoOtherInfo(); }
        if (def.key === 'case') { updateAutoOtherInfo(); }
        render();
      });
      row.querySelectorAll('.part-custom input').forEach(el => el.addEventListener('input', render));
    });
  }

  // --- Extra parts (multiple) ---
  function addExtraPart() {
    const id = 'extra-' + (extraIdCounter++);
    const container = document.getElementById('extra-parts-list');
    const row = document.createElement('div');
    row.className = 'extra-row';
    row.dataset.extraId = id;
    row.innerHTML = `
      <select class="input select extra-select" data-id="${id}">${buildExtraOptions()}</select>
      <input type="number" class="input input-price extra-price" data-id="${id}" placeholder="¥">
      <button class="btn-remove-extra" data-id="${id}" title="削除">×</button>
    `;
    container.appendChild(row);

    const sel = row.querySelector('select');
    const priceInput = row.querySelector('.extra-price');
    const removeBtn = row.querySelector('.btn-remove-extra');

    sel.addEventListener('change', () => {
      if (sel.value && sel.value !== '__custom__') {
        const price = parseInt(sel.selectedOptions[0]?.dataset.price) || 0;
        priceInput.value = price;
      } else if (sel.value === '__custom__') {
        priceInput.value = '';
      }
      render();
    });
    priceInput.addEventListener('input', render);
    removeBtn.addEventListener('click', () => {
      row.remove();
      render();
    });
  }

  function getExtraParts() {
    const rows = document.querySelectorAll('.extra-row');
    const result = [];
    rows.forEach(row => {
      const sel = row.querySelector('select');
      const priceInput = row.querySelector('.extra-price');
      const val = sel.value;
      if (val && val !== '__custom__') {
        result.push({ name: val, price: parseInt(sel.selectedOptions[0]?.dataset.price) || 0 });
      } else if (val === '__custom__' || priceInput.value) {
        result.push({ name: val === '__custom__' ? '' : val, price: parseFloat(priceInput.value) || 0 });
      }
    });
    return result;
  }

  function getExtraTotal() {
    return getExtraParts().reduce((s, p) => s + p.price, 0);
  }

  document.getElementById('btn-add-extra')?.addEventListener('click', () => {
    addExtraPart();
  });

  // --- Refresh select ---
  function refreshSelect(key) {
    const sel = document.getElementById(key + '-select');
    if (!sel) return;
    const prev = sel.value;
    sel.innerHTML = buildOptions(key);
    if (Array.from(sel.options).find(o => o.value === prev)) sel.value = prev;
    else { sel.value = ''; document.getElementById(key + '-custom-wrap')?.classList.add('hidden'); const pd = document.getElementById(key + '-price-display'); if (pd) pd.textContent = ''; }
  }

  function getPartInfo(def) {
    const sel = document.getElementById(def.key + '-select');
    if (!sel) return { name: '', price: 0 };
    if (sel.value === '__custom__') return { name: document.getElementById(def.key + '-name')?.value || '', price: parseFloat(document.getElementById(def.key + '-price')?.value) || 0 };
    if (sel.value) return { name: sel.value, price: parseInt(sel.selectedOptions[0]?.dataset.price) || 0 };
    return { name: '', price: 0 };
  }

  // --- Auto-fill "その他" from MB (AD) + Case (AE) ---
  function updateAutoOtherInfo() {
    const mbInfo = getPartInfo(partDefs.find(d => d.key === 'motherboard'));
    const caseInfo = getPartInfo(partDefs.find(d => d.key === 'case'));
    const autoText = getAutoOtherInfo(mbInfo.name, caseInfo.name);
    const textarea = document.getElementById('other-info');
    if (textarea && autoText) {
      textarea.value = autoText;
    }
  }

  // --- PC Type + Subtype ---
  function updateVisibility() {
    btnDesktop.classList.toggle('active', pcType === 'desktop');
    btnNotebook.classList.toggle('active', pcType === 'notebook');
    desktopSubEl.classList.toggle('hidden', pcType !== 'desktop');
    step3Desktop.classList.toggle('hidden', !(pcType === 'desktop' && desktopSub === 'gaming'));
    step3Slim.classList.toggle('hidden', !(pcType === 'desktop' && desktopSub === 'slim'));
    step3Notebook.classList.toggle('hidden', pcType !== 'notebook');
    btnGaming.classList.toggle('active', desktopSub === 'gaming');
    btnSlim.classList.toggle('active', desktopSub === 'slim');
  }

  function setPCType(type) {
    pcType = type;
    updateVisibility();
    render();
  }

  function setDesktopSub(sub) {
    desktopSub = sub;
    updateVisibility();
    render();
  }

  btnDesktop.addEventListener('click', () => setPCType('desktop'));
  btnNotebook.addEventListener('click', () => setPCType('notebook'));
  btnGaming.addEventListener('click', () => setDesktopSub('gaming'));
  btnSlim.addEventListener('click', () => setDesktopSub('slim'));

  // --- Slim parts (same approach as gaming but using MASTER_SLIM data) ---
  const slimPartDefs = [
    { key: 'cpu',         badge: 'CPU',    quoteLabel: 'CPU' },
    { key: 'cooler',      badge: 'クーラー', quoteLabel: 'CPUクーラー' },
    { key: 'motherboard', badge: 'M/B',    quoteLabel: 'マザーボード' },
    { key: 'memory',      badge: 'メモリ',  quoteLabel: 'メモリ' },
    { key: 'gpu',         badge: 'GPU',    quoteLabel: 'GPU' },
    { key: 'ssd',         badge: 'SSD',    quoteLabel: 'SSD' },
    { key: 'case',        badge: 'ケース',  quoteLabel: 'ケース' },
    { key: 'caseFan',     badge: 'ファン',  quoteLabel: 'ケースファン' },
    { key: 'psu',         badge: '電源',   quoteLabel: '電源' },
    { key: 'os',          badge: 'OS',     quoteLabel: 'OS' },
  ];

  function buildSlimSelectOptions(key) {
    const items = MASTER_SLIM[key] || [];
    let html = '<option value="">-- 選択 --</option>';
    items.forEach(item => {
      html += `<option value="${item.name}" data-price="${item.price}">${item.name}</option>`;
    });
    html += '<option value="__custom__">手動入力</option>';
    return html;
  }

  function buildSlimPartRows() {
    const list = document.getElementById('slim-parts-list');
    list.innerHTML = '';
    slimPartDefs.forEach(def => {
      const row = document.createElement('div');
      row.className = 'part-row';
      row.innerHTML = `
        <div class="part-badge">${def.badge}</div>
        <div class="part-select-wrap">
          <select id="slim-${def.key}-select" class="input select part-select">${buildSlimSelectOptions(def.key)}</select>
          <div class="part-custom hidden" id="slim-${def.key}-custom-wrap">
            <input type="text" id="slim-${def.key}-name" class="input" placeholder="パーツ名">
            <input type="number" id="slim-${def.key}-price" class="input input-price" placeholder="¥ 金額">
          </div>
        </div>
        <div class="part-price-display" id="slim-${def.key}-price-display"></div>
      `;
      list.appendChild(row);

      const select = row.querySelector('select');
      select.addEventListener('change', () => {
        const val = select.value;
        const cw = row.querySelector('.part-custom');
        const pd = row.querySelector('.part-price-display');
        if (val === '__custom__') { cw.classList.remove('hidden'); pd.textContent = ''; }
        else if (val) { cw.classList.add('hidden'); pd.textContent = '¥' + (parseInt(select.selectedOptions[0]?.dataset.price) || 0).toLocaleString(); }
        else { cw.classList.add('hidden'); pd.textContent = ''; }
        updateSlimModelMatch();
        updateSlimAutoOther();
        render();
      });
      row.querySelectorAll('.part-custom input').forEach(el => el.addEventListener('input', render));
    });
  }

  function getSlimPartInfo(def) {
    const sel = document.getElementById('slim-' + def.key + '-select');
    if (!sel) return { name: '', price: 0 };
    if (sel.value === '__custom__') return { name: document.getElementById('slim-' + def.key + '-name')?.value || '', price: parseFloat(document.getElementById('slim-' + def.key + '-price')?.value) || 0 };
    if (sel.value) return { name: sel.value, price: parseInt(sel.selectedOptions[0]?.dataset.price) || 0 };
    return { name: '', price: 0 };
  }

  function calcSlimPartsTotal() {
    return slimPartDefs.reduce((s, d) => s + getSlimPartInfo(d).price, 0) + getSlimExtraTotal();
  }

  // Slim extra parts
  function addSlimExtraPart() {
    const id = 'slim-extra-' + (extraIdCounter++);
    const container = document.getElementById('slim-extra-parts-list');
    const row = document.createElement('div');
    row.className = 'extra-row';
    row.innerHTML = `
      <select class="input select extra-select" data-id="${id}"><option value="">-- 選択 --</option><option value="__custom__">手動入力</option></select>
      <input type="number" class="input input-price extra-price" data-id="${id}" placeholder="¥">
      <button class="btn-remove-extra" data-id="${id}" title="削除">×</button>
    `;
    container.appendChild(row);
    const sel = row.querySelector('select');
    const priceInput = row.querySelector('.extra-price');
    sel.addEventListener('change', () => { render(); });
    priceInput.addEventListener('input', render);
    row.querySelector('.btn-remove-extra').addEventListener('click', () => { row.remove(); render(); });
  }

  function getSlimExtraParts() {
    const rows = document.querySelectorAll('#slim-extra-parts-list .extra-row');
    const result = [];
    rows.forEach(row => {
      const sel = row.querySelector('select');
      const pi = row.querySelector('.extra-price');
      if (sel.value === '__custom__' || pi.value) result.push({ name: '', price: parseFloat(pi.value) || 0 });
      else if (sel.value) result.push({ name: sel.value, price: parseFloat(pi.value) || 0 });
    });
    return result;
  }

  function getSlimExtraTotal() {
    return getSlimExtraParts().reduce((s, p) => s + p.price, 0);
  }

  document.getElementById('btn-add-slim-extra')?.addEventListener('click', () => addSlimExtraPart());

  // Auto-detect model name from current slim part selections
  function updateSlimModelMatch() {
    const selections = {};
    slimPartDefs.forEach(def => { selections[def.key] = getSlimPartInfo(def).name; });
    const modelName = matchSlimModel(selections);
    const badge = document.getElementById('slim-model-badge');
    const nameEl = document.getElementById('slim-model-name');
    if (modelName) {
      badge.classList.remove('hidden');
      nameEl.textContent = modelName;
    } else {
      badge.classList.add('hidden');
      nameEl.textContent = '';
    }
  }

  // Auto-fill "その他" for slim
  function updateSlimAutoOther() {
    const mbInfo = getSlimPartInfo(slimPartDefs.find(d => d.key === 'motherboard'));
    const caseInfo = getSlimPartInfo(slimPartDefs.find(d => d.key === 'case'));
    const autoText = getSlimAutoOtherInfo(mbInfo.name, caseInfo.name);
    const textarea = document.getElementById('other-info');
    if (textarea && autoText) textarea.value = autoText;
  }

  // --- Helpers ---
  function v(id) { return document.getElementById(id)?.value || ''; }
  function n(id) { return parseFloat(document.getElementById(id)?.value) || 0; }
  function yen(x) { return '¥' + Math.round(x).toLocaleString('ja-JP'); }
  function today() { const d = new Date(); return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`; }

  function calcPartsTotal() {
    if (pcType === 'notebook') return n('nb-cost');
    if (pcType === 'desktop' && desktopSub === 'slim') {
      return calcSlimPartsTotal();
    }
    return partDefs.reduce((s, d) => s + getPartInfo(d).price, 0) + getExtraTotal();
  }

  // --- Memo Templates ---
  const MEMO_TEMPLATES = {
    default: `見積書の有効期限は見積日より7日間とさせていただきます。
ご購入時期により、販売価格が変更になる場合がございます。
※指定パーツの価格は変動する可能性がございます。
※マザーボードの構成内容は実際に搭載される内容と異なる場合がございます。
ご指定がある場合は事前にお申し付けください。
有効期間内であっても市場価格の変動により、最終的なお支払金額が変更となる場合がございます。`,
    short: `見積書の有効期限は見積日より7日間です。
パーツ価格は市場変動により変更となる場合がございます。`,
    warranty: `見積書の有効期限は見積日より7日間とさせていただきます。
ご購入時期により、販売価格が変更になる場合がございます。
※指定パーツの価格は変動する可能性がございます。
※マザーボードの構成内容は実際に搭載される内容と異なる場合がございます。
ご指定がある場合は事前にお申し付けください。
有効期間内であっても市場価格の変動により、最終的なお支払金額が変更となる場合がございます。
【保証】本製品には1年間のパーツ保証が付属します。初期不良は納品後14日以内にご連絡ください。`,
  };

  document.getElementById('btn-apply-template')?.addEventListener('click', () => {
    const sel = document.getElementById('memo-template-select');
    const key = sel?.value;
    if (key && MEMO_TEMPLATES[key]) {
      document.getElementById('memo').value = MEMO_TEMPLATES[key];
      render();
    }
  });

  // --- Profit Calculation ---
  function updateProfit() {
    const costUnitRaw = calcPartsTotal();
    const costUnitExTax = Math.round(costUnitRaw / 1.1);
    const profitInput = n('profit-input');                        // 利益額（手入力）
    const qty = Math.max(1, Math.round(n('pc-qty'))) || 1;
    const shippingUnit = n('shipping-fee');
    const shippingQty = parseInt(document.getElementById('shipping-qty')?.value) || 1;
    const shipping = shippingUnit * shippingQty;
    const discountEx = n('discount-value');

    // 原価合計（税抜）= パーツ原価（税抜）× 数量 + 送料
    const costTotal = (costUnitExTax * qty) + shipping;
    // 販売金額（税抜）= 原価合計 + 利益額
    const sellingBefore = costTotal + profitInput;
    // 利益率
    const rateBefore = sellingBefore > 0 ? (profitInput / sellingBefore * 100) : 0;

    document.getElementById('cost-total').textContent = yen(costTotal);
    document.getElementById('selling-display').textContent = yen(sellingBefore);

    const profitEl = document.getElementById('profit-amount');
    profitEl.textContent = yen(profitInput);
    profitEl.className = 'profit-value ' + (profitInput >= 0 ? 'positive' : 'negative');

    const rateEl = document.getElementById('profit-rate');
    rateEl.textContent = rateBefore.toFixed(1) + '%';
    rateEl.className = 'profit-value ' + (profitInput >= 0 ? 'positive' : 'negative');

    // 特別値引きボックス（値引き後）
    const sellingAfter = sellingBefore - discountEx;
    const profitAfter = profitInput - discountEx;
    const rateAfter = sellingAfter > 0 ? (profitAfter / sellingAfter * 100) : 0;

    document.getElementById('discount-selling-display').textContent = yen(sellingAfter);

    const dpEl = document.getElementById('discount-profit-display');
    dpEl.textContent = yen(profitAfter);
    dpEl.className = 'profit-value ' + (profitAfter >= 0 ? 'positive' : 'negative');

    const drEl = document.getElementById('discount-rate-display');
    drEl.textContent = rateAfter.toFixed(1) + '%';
    drEl.className = 'profit-value ' + (profitAfter >= 0 ? 'positive' : 'negative');
  }

  // --- AI Suggest: search recommendations + apply ---
  document.getElementById('btn-ai-suggest')?.addEventListener('click', () => {
    const query = v('ai-prompt');
    const budgetMin = n('ai-budget-min');
    const budgetMax = n('ai-budget-max');

    const results = searchRecommendations(query, budgetMin, budgetMax);
    const resultsEl = document.getElementById('ai-results');
    resultsEl.classList.remove('hidden');

    if (results.length === 0) {
      resultsEl.innerHTML = `<div class="ai-card-placeholder">条件に一致する構成が見つかりませんでした。<br>キーワードや予算範囲を変更してお試しください。</div>`;
      return;
    }

    resultsEl.innerHTML = results.map((r, i) => {
      const price = calcRecommendationPrice(r);
      return `
        <div class="ai-card" data-rec-index="${i}">
          <div class="ai-card-title">${r.name}</div>
          <div class="ai-card-desc">${r.description}</div>
          <div class="ai-card-price">概算 ¥${price.toLocaleString()}（税抜）</div>
        </div>
      `;
    }).join('');

    // Store results for click handler
    resultsEl._results = results;

    // Card click → apply parts
    resultsEl.querySelectorAll('.ai-card').forEach(card => {
      card.addEventListener('click', () => {
        const idx = parseInt(card.dataset.recIndex);
        const rec = resultsEl._results[idx];
        if (!rec) return;

        // Highlight selected card
        resultsEl.querySelectorAll('.ai-card').forEach(c => c.classList.remove('selected'));
        card.classList.add('selected');

        // Apply each part to selects
        applyRecommendation(rec);
      });
    });
  });

  function applyRecommendation(rec) {
    const keyMap = [
      'cpu', 'cooler', 'motherboard', 'memory', 'gpu',
      'ssd', 'case', 'caseFan', 'psu', 'os'
    ];

    // Apply in order (CPU first, then MB, then memory — for compatibility filters)
    keyMap.forEach(key => {
      const partName = rec.parts[key];
      if (!partName) return;

      const sel = document.getElementById(key + '-select');
      if (!sel) return;

      // If it's motherboard or memory, refresh options first (compatibility)
      if (key === 'motherboard') refreshSelect('motherboard');
      if (key === 'memory') refreshSelect('memory');

      // Try to find matching option
      const option = Array.from(sel.options).find(o => o.value === partName);
      if (option) {
        sel.value = partName;
        sel.dispatchEvent(new Event('change'));
      }
    });

    render();
  }

  // --- Render ---
  function render() {
    const customer = v('customer-name');
    const costUnitRaw = calcPartsTotal();                        // 1台あたりパーツ原価（税込）
    const costUnitExTax = Math.round(costUnitRaw / 1.1);         // 1台あたりパーツ原価（税抜）
    const profitInput = n('profit-input');                        // 利益額（手入力）
    const pcUnit = costUnitExTax + profitInput;                   // 1台あたり販売単価（税抜）
    const qty = Math.max(1, Math.round(n('pc-qty'))) || 1;
    const pcBody = pcUnit * qty;                                  // PC本体合計（税抜）
    const discountEx = n('discount-value');
    const shippingUnit = n('shipping-fee');
    const shippingQty = parseInt(document.getElementById('shipping-qty')?.value) || 1;
    const shipping = shippingUnit * shippingQty;
    const otherInfo = v('other-info');
    const memo = v('memo');

    // すべて税抜きベース
    const breakdownTotal = pcBody - discountEx + shipping;        // 合計金額（税抜）
    const subtaxEx = pcBody + shipping;                           // 小計（税抜）
    const tax = Math.round((subtaxEx - discountEx) * 0.1);       // 消費税
    const grandTotal = (subtaxEx - discountEx) + tax;             // 合計金額（税込）

    // Spec rows
    let specRows = '';
    if (pcType === 'desktop' && desktopSub === 'slim') {
      // Slim: パーツ個別選択 with display names from X列以降
      slimPartDefs.forEach(def => {
        const info = getSlimPartInfo(def);
        const displayName = getSlimDisplayName(def.key, info.name);
        specRows += `<tr><td class="sl">${def.quoteLabel}</td><td class="sv">${displayName}</td></tr>`;
      });
      // Extra parts
      const slimExtras = getSlimExtraParts();
      if (slimExtras.length > 0) {
        slimExtras.forEach((ex, i) => {
          specRows += `<tr><td class="sl">${i === 0 ? '追加パーツ' : ''}</td><td class="sv">${ex.name}</td></tr>`;
        });
      } else {
        specRows += `<tr><td class="sl">追加パーツ</td><td class="sv"></td></tr>`;
      }
      specRows += `<tr><td class="sl">その他</td><td class="sv other-cell">${otherInfo.replace(/\n/g, '<br>')}</td></tr>`;
    } else if (pcType === 'desktop') {
      // Gaming: individual parts
      partDefs.forEach(def => {
        const info = getPartInfo(def);
        const displayName = getDisplayName(def.key, info.name);
        specRows += `<tr><td class="sl">${def.quoteLabel}</td><td class="sv">${displayName}</td></tr>`;
      });
      const extras = getExtraParts();
      if (extras.length > 0) {
        extras.forEach((ex, i) => {
          specRows += `<tr><td class="sl">${i === 0 ? '追加パーツ' : ''}</td><td class="sv">${ex.name}</td></tr>`;
        });
      } else {
        specRows += `<tr><td class="sl">追加パーツ</td><td class="sv"></td></tr>`;
      }
      specRows += `<tr><td class="sl">その他</td><td class="sv other-cell">${otherInfo.replace(/\n/g, '<br>')}</td></tr>`;
    } else {
      [['メーカー/型番', v('nb-model')], ['CPU', v('nb-cpu')], ['メモリ', v('nb-memory')], ['ストレージ', v('nb-storage')], ['画面サイズ', v('nb-display')],
       ['状態', document.getElementById('nb-condition')?.options[document.getElementById('nb-condition').selectedIndex]?.text || '']
      ].forEach(([l, val]) => { specRows += `<tr><td class="sl">${l}</td><td class="sv">${val}</td></tr>`; });
      specRows += `<tr><td class="sl">その他</td><td class="sv other-cell">${otherInfo.replace(/\n/g, '<br>')}</td></tr>`;
    }

    const pcLabel = pcType === 'notebook' ? 'ノートPC本体'
      : desktopSub === 'slim' ? 'スリムPC本体'
      : 'デスクトップPC本体';

    quotePaper.innerHTML = `<div class="q">
      <div class="q-logo"><img src="logo.png" alt="MDL.make" class="q-logo-img" onerror="this.parentElement.innerHTML='<span style=&quot;font-size:16px;font-weight:800;&quot;>MDL.make</span>'"></div>
      <h1 class="q-title">見 積 書</h1>
      <div class="q-customer-row">
        <div class="q-customer-line"><span class="q-customer-text">${customer}</span></div>
        <span class="q-sama">様</span>
        <span class="q-date">見積日：　${today()}</span>
      </div>
      <div class="q-company-block">
        <div class="q-company-info"><div class="q-company-name">株式会社モダンデザイン</div><div>〒893-0022</div><div>鹿児島県鹿屋市旭原町2615-5</div></div>
        <div class="q-stamp-area"><img src="stamp.png" alt="会社印" class="q-stamp-img" onerror="this.style.display='none'"></div>
      </div>
      <div class="q-total-bar">
        <span class="q-total-lbl">合計金額</span>
        <span class="q-total-amt">${Math.round(breakdownTotal).toLocaleString('ja-JP')}</span>
        <span class="q-total-sfx">円 (税抜)</span>
      </div>
      <table class="q-bd">
        <thead><tr><th class="c-item">内訳</th><th class="c-qty">数量</th><th class="c-unit">単価</th><th class="c-amt">金額</th></tr></thead>
        <tbody>
          <tr><td class="c-item">${pcLabel}</td><td class="c-qty">${qty}　台</td><td class="c-unit">${yen(pcUnit)}</td><td class="c-amt">${yen(pcBody)}</td></tr>
          <tr class="discount-row"><td class="c-item">特別値引き</td><td class="c-qty">${discountEx > 0 ? qty + '　台' : ''}</td><td class="c-unit">${discountEx > 0 ? yen(discountEx) : ''}</td><td class="c-amt">${discountEx > 0 ? yen(discountEx) : ''}</td></tr>
          <tr class="sep-row"><td></td><td></td><td></td><td></td></tr>
          <tr><td class="c-item">送料</td><td class="c-qty">${shippingQty}　式</td><td class="c-unit">${yen(shippingUnit)}</td><td class="c-amt">${yen(shipping)}</td></tr>
          <tr class="sum-row"><td></td><td></td><td class="c-unit" style="font-weight:700;">合計</td><td class="c-amt" style="font-weight:700;">${yen(breakdownTotal)}</td></tr>
        </tbody>
      </table>
      <table class="q-sp">${specRows}</table>
      <div class="q-price-wrap"><table class="q-ps">
        <tr><td class="pl">小計(税抜)</td><td class="pv">${yen(subtaxEx)}</td></tr>
        <tr class="discount-row"><td class="pl">特別値引き</td><td class="pv">${discountEx > 0 ? yen(discountEx) : ''}</td></tr>
        <tr><td class="pl">消費税</td><td class="pv">${yen(tax)}</td></tr>
        <tr class="grand-row"><td class="pl">合計金額(税込)</td><td class="pv">${yen(grandTotal)}</td></tr>
      </table></div>
      <div class="q-notes"><b>【備考】</b><br>
        ${memo.replace(/\n/g, '<br>')}
      </div>
    </div>`;

    updateProfit();
  }

  // --- Global listeners ---
  document.querySelectorAll('aside input, aside select, aside textarea').forEach(el => {
    el.addEventListener('input', render);
    el.addEventListener('change', render);
  });

  // --- Reload ---
  document.getElementById('btn-reload')?.addEventListener('click', async () => {
    if (loadingEl) loadingEl.classList.remove('hidden');
    const ok = await fetchMasterData();
    if (loadingEl) loadingEl.classList.add('hidden');
    await fetchSlimData();
    if (ok) { buildPartRows(); buildSlimPartRows(); render(); document.getElementById('data-status').textContent = '再取得済み (' + new Date().toLocaleTimeString() + ')'; }
  });

  document.getElementById('btn-print').addEventListener('click', () => {
    // iframe印刷でブラウザのヘッダー/フッターを回避
    const quoteHTML = document.getElementById('quote-paper').innerHTML;
    const printCSS = Array.from(document.styleSheets)
      .map(s => { try { return Array.from(s.cssRules).map(r => r.cssText).join('\n'); } catch(e) { return ''; } })
      .join('\n');

    const iframe = document.createElement('iframe');
    iframe.style.cssText = 'position:fixed;top:-9999px;left:-9999px;width:0;height:0;border:none;';
    document.body.appendChild(iframe);

    const doc = iframe.contentDocument;
    doc.open();
    doc.write(`<!DOCTYPE html>
<html><head>
<meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
@page { size: A4 portrait; margin: 10mm; }
html, body { margin: 0; padding: 0; background: #fff; }
${printCSS}
.quote-paper { box-shadow: none; width: 100%; min-height: auto; }
.q { padding: 16px 24px; font-size: 10px; }
.q-title { font-size: 22px; margin: 2px 0 12px; }
.q-logo-img { height: 32px; }
.q-customer-row { margin-top: 12px; }
.q-customer-text, .q-sama { font-size: 13px; }
.q-date { font-size: 11px; }
.q-stamp-img { width: 64px; height: 64px; }
.q-total-bar { margin: 8px 0 12px; padding: 6px 4px; }
.q-total-amt { font-size: 22px; }
.q-bd { margin-bottom: 8px; }
.q-bd th { padding: 3px 6px; font-size: 9px; }
.q-bd td { padding: 3px 6px; font-size: 10px; }
.q-sp { margin-bottom: 8px; }
.q-sp td { padding: 2px 6px; font-size: 9.5px; }
.q-sp .sl { width: 80px; font-size: 9px; }
.q-sp .other-cell { font-size: 8.5px; line-height: 1.6; min-height: auto; }
.q-price-wrap { margin-bottom: 10px; }
.q-ps { width: 240px; font-size: 10px; }
.q-ps td { padding: 2px 8px; }
.q-ps .grand-row td { font-size: 11px; padding: 4px 8px; }
.q-notes { padding: 8px 10px; font-size: 8px; line-height: 1.7; }
.q-notes b { font-size: 8.5px; }
</style>
</head><body>
<div class="quote-paper">${quoteHTML}</div>
</body></html>`);
    doc.close();

    iframe.contentWindow.onafterprint = () => {
      document.body.removeChild(iframe);
    };

    // フォント読み込み待ちしてから印刷
    setTimeout(() => {
      iframe.contentWindow.print();
      // フォールバック: onafterprintが効かないブラウザ用
      setTimeout(() => {
        if (iframe.parentNode) document.body.removeChild(iframe);
      }, 5000);
    }, 500);
  });

  document.getElementById('btn-reset').addEventListener('click', () => {
    if (!confirm('入力内容をすべてリセットしますか？')) return;
    document.querySelectorAll('input[type="text"], input[type="number"]').forEach(el => el.value = '');
    document.querySelectorAll('.part-custom').forEach(el => el.classList.add('hidden'));
    document.querySelectorAll('.part-price-display').forEach(el => el.textContent = '');
    document.getElementById('other-info').value = '';
    document.getElementById('memo').value = '';
    document.getElementById('pc-qty').value = '1';
    document.getElementById('discount-value').value = '0';
    document.getElementById('shipping-fee').value = '2000';
    document.getElementById('shipping-qty').value = '1';
    document.getElementById('extra-parts-list').innerHTML = '';
    document.getElementById('ai-results').classList.add('hidden');
    document.getElementById('ai-prompt').value = '';
    document.getElementById('slim-extra-parts-list').innerHTML = '';
    document.getElementById('slim-model-badge')?.classList.add('hidden');
    desktopSub = 'gaming';
    buildPartRows(); setPCType('desktop');
  });

  // --- Init ---
  buildPartRows();
  buildSlimPartRows();
  updateVisibility();
  render();
});
