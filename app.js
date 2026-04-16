// ========================================
// PC Quote Tool
// Live spreadsheet + compatibility filter + AI suggest + multi extras
// ========================================

document.addEventListener('DOMContentLoaded', async () => {
  let pcType = 'desktop';
  let extraParts = []; // Array of { id, selectEl, priceEl }
  let extraIdCounter = 0;

  const btnDesktop    = document.getElementById('btn-desktop');
  const btnNotebook   = document.getElementById('btn-notebook');
  const step3Desktop  = document.getElementById('step3-desktop');
  const step3Notebook = document.getElementById('step3-notebook');
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
      html += `<option value="${item.name}" data-price="${item.price}">${item.name}　(¥${item.price.toLocaleString()})</option>`;
    });
    html += '<option value="__custom__">手動入力</option>';
    return html;
  }

  function buildExtraOptions() {
    const items = MASTER.extra || [];
    let html = '<option value="">-- 選択 --</option>';
    items.forEach(item => {
      html += `<option value="${item.name}" data-price="${item.price}">${item.name}　(¥${item.price.toLocaleString()})</option>`;
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

  // --- PC Type ---
  function setPCType(type) {
    pcType = type;
    btnDesktop.classList.toggle('active', type === 'desktop');
    btnNotebook.classList.toggle('active', type === 'notebook');
    step3Desktop.classList.toggle('hidden', type !== 'desktop');
    step3Notebook.classList.toggle('hidden', type !== 'notebook');
    render();
  }
  btnDesktop.addEventListener('click', () => setPCType('desktop'));
  btnNotebook.addEventListener('click', () => setPCType('notebook'));

  // --- Helpers ---
  function v(id) { return document.getElementById(id)?.value || ''; }
  function n(id) { return parseFloat(document.getElementById(id)?.value) || 0; }
  function yen(x) { return '¥' + Math.round(x).toLocaleString('ja-JP'); }
  function today() { const d = new Date(); return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`; }

  function calcPartsTotal() {
    if (pcType === 'notebook') return n('nb-cost');
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
    const costUnit = calcPartsTotal();
    const profitInput = n('profit-input');
    const qty = Math.max(1, Math.round(n('pc-qty'))) || 1;
    const cost = costUnit * qty;                        // 原価合計
    const pcBody = (costUnit + profitInput) * qty;      // PC本体価格（原価+利益）
    const shipping = n('shipping-fee');
    const discountEx = n('discount-value');

    // 利益管理（値引き前）
    const sellingBefore = pcBody + shipping;
    const profitBefore = sellingBefore - cost;
    const rateBefore = sellingBefore > 0 ? (profitBefore / sellingBefore * 100) : 0;

    document.getElementById('cost-total').textContent = yen(cost);
    document.getElementById('selling-display').textContent = yen(sellingBefore);

    const profitEl = document.getElementById('profit-amount');
    profitEl.textContent = yen(profitBefore);
    profitEl.className = 'profit-value ' + (profitBefore >= 0 ? 'positive' : 'negative');

    const rateEl = document.getElementById('profit-rate');
    rateEl.textContent = rateBefore.toFixed(1) + '%';
    rateEl.className = 'profit-value ' + (profitBefore >= 0 ? 'positive' : 'negative');

    // 特別値引きボックス（値引き後）
    const sellingAfter = pcBody - discountEx + shipping;
    const profitAfter = sellingAfter - cost;
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
    const costUnit = calcPartsTotal();             // 1台あたり原価
    const profitInput = n('profit-input');          // 利益額
    const pcUnit = costUnit + profitInput;          // 1台あたり販売単価（原価+利益）
    const qty = Math.max(1, Math.round(n('pc-qty'))) || 1;
    const pcBody = pcUnit * qty;
    const discountEx = n('discount-value');
    const shipping = n('shipping-fee');
    const otherInfo = v('other-info');
    const memo = v('memo');

    const breakdownTotal = pcBody - discountEx + shipping;
    const subtaxEx = pcBody + shipping;
    const tax = Math.round((subtaxEx - discountEx) * 0.1);
    const grandTotal = (subtaxEx - discountEx) + tax;

    // Spec rows
    let specRows = '';
    if (pcType === 'desktop') {
      partDefs.forEach(def => {
        const info = getPartInfo(def);
        const displayName = getDisplayName(def.key, info.name);
        specRows += `<tr><td class="sl">${def.quoteLabel}</td><td class="sv">${displayName}</td></tr>`;
      });
      // Extra parts rows
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

    const pcLabel = pcType === 'desktop' ? 'デスクトップPC本体' : 'ノートPC本体';

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
          <tr><td class="c-item">送料</td><td class="c-qty">1　式</td><td class="c-unit">${yen(shipping)}</td><td class="c-amt">${yen(shipping)}</td></tr>
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
    if (ok) { buildPartRows(); render(); document.getElementById('data-status').textContent = '再取得済み (' + new Date().toLocaleTimeString() + ')'; }
  });

  document.getElementById('btn-print').addEventListener('click', () => window.print());

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
    document.getElementById('extra-parts-list').innerHTML = '';
    document.getElementById('ai-results').classList.add('hidden');
    document.getElementById('ai-prompt').value = '';
    buildPartRows(); setPCType('desktop');
  });

  // --- Init ---
  buildPartRows();
  render();
});
