// ========================================
// Master Data - Live fetch from Google Sheets
// + CPU/Motherboard/Memory compatibility rules
// ========================================

const SHEET_ID = '1NDLxP6WXsxVVcZvWdoGM-o7TV_PexJHJHNu_peja0bI';
const GAMING_GID = '1024358522';
const SLIM_GID   = '1943617209';

// Use gviz endpoint (more reliable — returns calculated values)
function sheetURL(gid) {
  return `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&gid=${gid}`;
}

// Global master data (populated on load)
let MASTER = {
  cpu: [], cooler: [], motherboard: [], memory: [], gpu: [],
  ssd: [], case: [], caseFan: [], psu: [], extra: [],
};

// Slim master data — array of preset models
let MASTER_SLIM = [];  // [{ model, cpu, cooler, mb, memory, gpu, ssd, case, fan, psu, extra, os, prices:{}, display:{} }]

// ========================================
// Compatibility Rules
// ========================================

// CPU → Socket mapping
const CPU_SOCKET = {
  // AM4
  'Ryzen 5 5500':    'AM4', 'Ryzen 5 5500GT':  'AM4',
  'Ryzen 7 5700X':   'AM4', 'Ryzen 7 5700X3D': 'AM4',
  // AM5 (7000 series)
  'Ryzen 5 7500F':   'AM5', 'Ryzen 7 7700':    'AM5',
  'Ryzen 7 7800X3D': 'AM5', 'Ryzen 9 7950X':   'AM5',
  'Ryzen 9 7950X3D': 'AM5',
  // AM5 (9000 series)
  'Ryzen 5 9600X':   'AM5', 'Ryzen 7 9700X':   'AM5',
  'Ryzen 7 9800X3D': 'AM5', 'Ryzen 7 9850X3D': 'AM5',
  'Ryzen 9 9900X':   'AM5', 'Ryzen 9 9950X':   'AM5',
  'Ryzen 9 9950X3D': 'AM5',
  // LGA1700 (12th gen)
  'Core i5 12400':   'LGA1700', 'Core i5 12400F':  'LGA1700',
  // LGA1700 (13th gen)
  'Core i5 13400F':  'LGA1700', 'Core i7 13700F':  'LGA1700',
  'Core i7 13700KF': 'LGA1700',
  // LGA1700 (14th gen)
  'Core i5 14400F':  'LGA1700', 'Core i5 14600KF': 'LGA1700',
  'Core i7 14700F':  'LGA1700', 'Core i7 14700KF': 'LGA1700',
  'Core i7 14700K':  'LGA1700', 'Core i9 14900KF': 'LGA1700',
  'Core i9 14900K':  'LGA1700',
  // LGA1851 (Core Ultra)
  'Core Ultra 5 225':   'LGA1851', 'Core Ultra 5 245KF': 'LGA1851',
  'Core Ultra 7 265KF': 'LGA1851', 'Core Ultra 9 285K':  'LGA1851',
};

// Motherboard → Socket mapping (by chipset prefix)
function getMBSocket(name) {
  const n = name.toUpperCase();
  if (n.startsWith('A520') || n.startsWith('B550'))  return 'AM4';
  if (n.startsWith('B650') || n.startsWith('X670') || n.startsWith('X870') || n.startsWith('B660')) return 'AM5';
  if (n.startsWith('B760') || n.startsWith('Z790'))  return 'LGA1700';
  if (n.startsWith('B860') || n.startsWith('Z890'))  return 'LGA1851';
  return null; // unknown → show always
}

// Motherboard → DDR type
function getMBDDR(name) {
  if (name.includes('D4') || name.includes('_D4')) return 'DDR4';
  if (name.includes('D5') || name.includes('_D5')) return 'DDR5';
  // Fallback based on chipset
  const n = name.toUpperCase();
  if (n.startsWith('A520') || n.startsWith('B550')) return 'DDR4';
  return 'DDR5'; // Modern chipsets default to DDR5
}

// Memory → DDR type
function getMemDDR(name) {
  if (name.includes('D4') || name.includes('_D4')) return 'DDR4';
  if (name.includes('D5') || name.includes('_D5')) return 'DDR5';
  return null;
}

// Filter motherboards by CPU socket
function filterMotherboards(cpuName) {
  const socket = CPU_SOCKET[cpuName];
  if (!socket) return MASTER.motherboard; // no filter if unknown
  return MASTER.motherboard.filter(mb => {
    const mbSocket = getMBSocket(mb.name);
    return !mbSocket || mbSocket === socket;
  });
}

// Filter memory by motherboard DDR type
function filterMemory(mbName) {
  if (!mbName) return MASTER.memory;
  const ddr = getMBDDR(mbName);
  if (!ddr) return MASTER.memory;
  return MASTER.memory.filter(m => {
    const memDDR = getMemDDR(m.name);
    return !memDDR || memDDR === ddr;
  });
}

// ========================================
// CSV Fetch & Parse
// ========================================

function parsePrice(str) {
  if (!str) return 0;
  return parseInt(str.replace(/[¥,\s]/g, ''), 10) || 0;
}

function parseCSV(text) {
  // Full CSV parser that handles quoted fields with newlines, commas, and double-quotes
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  let i = 0;

  while (i < text.length) {
    const ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < text.length && text[i + 1] === '"') {
          field += '"';  // escaped quote
          i += 2;
        } else {
          inQuotes = false;  // end of quoted field
          i++;
        }
      } else {
        field += ch;
        i++;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
        i++;
      } else if (ch === ',') {
        row.push(field.trim());
        field = '';
        i++;
      } else if (ch === '\n' || (ch === '\r' && i + 1 < text.length && text[i + 1] === '\n')) {
        row.push(field.trim());
        rows.push(row);
        row = [];
        field = '';
        i += (ch === '\r') ? 2 : 1;
      } else if (ch === '\r') {
        row.push(field.trim());
        rows.push(row);
        row = [];
        field = '';
        i++;
      } else {
        field += ch;
        i++;
      }
    }
  }
  // Last field/row
  if (field || row.length > 0) {
    row.push(field.trim());
    rows.push(row);
  }

  return rows;
}

function extractColumn(rows, nameCol, priceCol, displayCol, otherCol) {
  const items = [];
  const seen = new Set();
  for (let i = 1; i < rows.length; i++) {
    const name = (rows[i][nameCol] || '').replace(/_$/, '').trim();
    const price = parsePrice(rows[i][priceCol]);
    let display = '';
    if (displayCol !== undefined) {
      const raw = (rows[i][displayCol] || '').trim();
      if (raw && !raw.startsWith('見積書')) {
        display = raw;
      }
    }
    // otherCol: AD/AE column for "その他" auto-fill
    let otherInfo = '';
    if (otherCol !== undefined) {
      const raw = (rows[i][otherCol] || '').trim();
      if (raw && !raw.startsWith('見積書')) {
        otherInfo = raw;
      }
    }
    if (name && !seen.has(name)) {
      seen.add(name);
      items.push({ name, price, display, otherInfo });
    }
  }
  return items;
}

// Build display name lookup: partKey → { internalName → displayName }
let DISPLAY_MAP = {};

function buildDisplayMap() {
  DISPLAY_MAP = {};
  const categories = ['cpu','cooler','motherboard','memory','gpu','ssd','case','psu'];
  categories.forEach(key => {
    DISPLAY_MAP[key] = {};
    (MASTER[key] || []).forEach(item => {
      if (item.display) {
        DISPLAY_MAP[key][item.name] = item.display;
      }
    });
  });
}

// Get display name for quote preview (fallback to internal name)
function getDisplayName(partKey, internalName) {
  if (!internalName) return '';
  return (DISPLAY_MAP[partKey] && DISPLAY_MAP[partKey][internalName]) || internalName;
}

// Get "その他" info from motherboard (AD) + case (AE)
function getAutoOtherInfo(mbName, caseName) {
  let parts = [];
  if (mbName) {
    const mb = (MASTER.motherboard || []).find(m => m.name === mbName);
    if (mb?.otherInfo) parts.push(mb.otherInfo);
  }
  if (caseName) {
    const c = (MASTER.case || []).find(m => m.name === caseName);
    if (c?.otherInfo) parts.push(c.otherInfo);
  }
  return parts.join('\n');
}

// Fetch with retry (Google Sheets sometimes returns "読み込んでいます..." on first try)
async function fetchCSVWithRetry(url, maxRetries = 3) {
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    const res = await fetch(url);
    if (!res.ok) throw new Error('Fetch failed: ' + res.status);
    const text = await res.text();
    // Check if it's a loading placeholder or error
    if (text.includes('読み込んでいます') || text.trim() === '#REF!' || text.trim().length < 20) {
      console.warn(`Attempt ${attempt + 1}: Sheet still loading, retrying...`);
      await new Promise(r => setTimeout(r, 2000 * (attempt + 1))); // Wait 2s, 4s, 6s
      continue;
    }
    return text;
  }
  throw new Error('Sheet data not ready after retries');
}

async function fetchMasterData() {
  try {
    const text = await fetchCSVWithRetry(sheetURL(GAMING_GID));
    const rows = parseCSV(text);

    if (rows.length < 2) throw new Error('No data');

    // Column mapping:
    // A=0:CPU名, B=1:仕入価格, C=2:AIO名, D=3:仕入価格,
    // E=4:MB名, F=5:仕入価格, G=6:メモリ名, H=7:仕入価格,
    // I=8:GPU名, J=9:仕入価格, K=10:SSD名, L=11:仕入価格,
    // M=12:ケース名, N=13:仕入価格, O=14:ケースファン名, P=15:仕入価格,
    // Q=16:電源名, R=17:仕入価格, S=18:追加パーツ, T=19:仕入価格
    // 見積書の表記 columns (actual indices after proper CSV parsing):
    // [21]=CPU表記, [22]=クーラー表記, [23]=MB表記, [24]=メモリ表記,
    // [25]=GPU表記, [26]=SSD表記, [27]=ケース表記(サイズ),
    // [28]=電源表記, [29]=その他(拡張), [30]=その他(USB)
    MASTER.cpu         = extractColumn(rows, 0, 1, 21);
    MASTER.cooler      = extractColumn(rows, 2, 3, 22);
    MASTER.motherboard = extractColumn(rows, 4, 5, 23, 29);  // col29 = AD(拡張スロット等)
    MASTER.memory      = extractColumn(rows, 6, 7, 24);
    MASTER.gpu         = extractColumn(rows, 8, 9, 25);
    MASTER.ssd         = extractColumn(rows, 10, 11, 26);
    MASTER.case        = extractColumn(rows, 12, 13, 27, 30); // col30 = AE(USBポート等)
    MASTER.caseFan     = extractColumn(rows, 14, 15);
    MASTER.psu         = extractColumn(rows, 16, 17, 28);
    MASTER.extra       = extractColumn(rows, 18, 19);

    // Remove items with price 0 that look like duplicates (e.g. 鹿児島 entries)
    Object.keys(MASTER).forEach(key => {
      MASTER[key] = MASTER[key].filter(item =>
        item.price > 0 || item.name === '不要'
      );
    });

    // Also extract "その他" display data from cols 29, 30
    // These are row-aligned with the data (not per-part), stored separately
    MASTER._otherInfo = [];
    for (let i = 1; i < rows.length; i++) {
      const col29 = (rows[i][29] || rows[i][28] || '').trim(); // expansion slots
      const col30 = (rows[i][30] || rows[i][29] || '').trim(); // USB ports
      if (col29 || col30) {
        MASTER._otherInfo.push({ row: i, expansion: col29, usb: col30 });
      }
    }

    buildDisplayMap();
    console.log('Master data loaded from Google Sheets');
    return true;
  } catch (err) {
    console.warn('Failed to fetch live data, using fallback:', err);
    return false;
  }
}

// ========================================
// 【スリム】マスタ管理 — Fetch preset models
// ========================================

async function fetchSlimData() {
  MASTER_SLIM = [];
  try {
    const text = await fetchCSVWithRetry(sheetURL(SLIM_GID));
    const rows = parseCSV(text);
    if (rows.length < 2) throw new Error('No slim data');

    // Slim columns (offset by 1 from gaming due to model name in col 0):
    // [0]=モデル名, [1]=CPU名, [2]=仕入価格, [3]=AIO名, [4]=仕入価格,
    // [5]=MB名, [6]=仕入価格, [7]=メモリ名, [8]=仕入価格,
    // [9]=GPU名, [10]=仕入価格, [11]=SSD名, [12]=仕入価格,
    // [13]=ケース名, [14]=仕入価格, [15]=ファン名, [16]=仕入価格,
    // [17]=電源名, [18]=仕入価格, [19]=追加パーツ, [20]=仕入価格,
    // [21]=OS, [22]=仕入価格(OS)
    // Display: [23]=CPU, [24]=クーラー, [25]=MB, [26]=メモリ,
    // [27]=GPU, [28]=SSD, [29]=ケース(サイズ), [30]=ファン,
    // [31]=電源, [32]=その他, [33]=USB

    for (let i = 1; i < rows.length; i++) {
      const model = (rows[i][0] || '').trim();
      if (!model) continue;

      const p = (col) => parsePrice(rows[i][col]);
      const s = (col) => (rows[i][col] || '').trim();

      MASTER_SLIM.push({
        model,
        parts: {
          cpu:    s(1), cooler: s(3), motherboard: s(5), memory: s(7),
          gpu:    s(9), ssd:    s(11), case: s(13), caseFan: s(15),
          psu:    s(17), extra:  s(19), os: s(21),
        },
        prices: {
          cpu: p(2), cooler: p(4), motherboard: p(6), memory: p(8),
          gpu: p(10), ssd: p(12), case: p(14), caseFan: p(16),
          psu: p(18), extra: p(20), os: p(22),
        },
        display: {
          cpu: s(23), cooler: s(24), motherboard: s(25), memory: s(26),
          gpu: s(27), ssd: s(28), case: s(29), caseFan: s(30),
          psu: s(31),
        },
        otherInfo: s(32),
        usbInfo: rows[i][33] ? (rows[i][33] || '').trim() : '',
        totalCost: p(2) + p(4) + p(6) + p(8) + p(10) + p(12) + p(14) + p(16) + p(18) + p(20) + p(22),
      });
    }

    console.log(`Loaded ${MASTER_SLIM.length} slim models`);
    return true;
  } catch (err) {
    console.warn('Failed to fetch slim data:', err);
    return false;
  }
}

// ========================================
// おすすめ構成 — Fetch from separate sheet
// ========================================

// Set this GID after creating the "おすすめ構成" sheet
const RECOMMEND_GID = '0'; // TODO: replace with actual GID

const RECOMMEND_CSV_URL = sheetURL(RECOMMEND_GID);

// Fallback test data (used when sheet doesn't exist yet)
const FALLBACK_RECOMMENDATIONS = [
  {
    name: 'ゲーミング入門（Apex/Valorant）',
    tags: 'ゲーミング,Apex,Valorant,FPS,入門,初心者',
    budgetMin: 100000, budgetMax: 150000,
    description: 'Apex LegendsやValorantが快適に動作するエントリーモデル。フルHD 144fpsを目指す構成。',
    parts: { cpu:'Ryzen 5 5500', cooler:'空冷ファンクーラー', motherboard:'B550M_D4_ARGB', memory:'16GB_D4_BYL', gpu:'5060_Black', ssd:'1TB_SSD', case:'CS032_M_Black', caseFan:'3個_Black', psu:'650W_BRONZE', os:'Windows11 Home' }
  },
  {
    name: 'ゲーミングミドル（高画質144fps）',
    tags: 'ゲーミング,Apex,フォートナイト,配信,ミドル',
    budgetMin: 150000, budgetMax: 220000,
    description: '高画質設定で144fps安定。ゲーム配信も可能なミドルレンジ構成。',
    parts: { cpu:'Ryzen 7 5700X', cooler:'強力空冷クーラー', motherboard:'B550M_D4_ARGB', memory:'16GB_D4_BYL', gpu:'5060 Ti_Black', ssd:'1TB_SSD', case:'CS032_M_Black', caseFan:'3個_Black', psu:'650W_GOLD', os:'Windows11 Home' }
  },
  {
    name: '動画編集・クリエイター向け',
    tags: '動画編集,クリエイター,Premiere,DaVinci,デザイン,イラスト',
    budgetMin: 180000, budgetMax: 300000,
    description: '4K動画編集やAdobe系ソフトが快適に動作。大容量メモリとGPU搭載。',
    parts: { cpu:'Ryzen 7 9700X', cooler:'240mm_Black', motherboard:'B650M_D5_ARGB', memory:'32GB_D5_BYL', gpu:'5070_Black', ssd:'2TB_SSD', case:'H5 Flow_ATX_Black', caseFan:'LED_3個_Black', psu:'750W_GOLD', os:'Windows11 Pro' }
  },
  {
    name: '事務・一般業務PC',
    tags: '事務,オフィス,業務,Excel,Word,一般,法人,テレワーク',
    budgetMin: 60000, budgetMax: 120000,
    description: 'Office作業、Web会議、メール等の一般業務に最適。コスパ重視。',
    parts: { cpu:'Ryzen 5 5500', cooler:'空冷ファンクーラー', motherboard:'A520M_D4', memory:'16GB_D4_BYL', gpu:'', ssd:'500GB_SSD', case:'CS032_M_Black', caseFan:'2個_Black', psu:'650W_BRONZE', os:'Windows11 Pro' }
  },
  {
    name: 'ハイエンドゲーミング（4K対応）',
    tags: 'ゲーミング,ハイエンド,4K,最高設定,配信,VR',
    budgetMin: 300000, budgetMax: 500000,
    description: '4K最高設定でプレイ可能。VRや高負荷ゲームも余裕のハイエンド。',
    parts: { cpu:'Ryzen 7 9800X3D', cooler:'360mm_液晶表示_Black', motherboard:'X870_D5_ARGB', memory:'32GB_D5_BYL', gpu:'5080_Black', ssd:'2TB_SSD', case:'H9 Flow_ATX_Black', caseFan:'LED_10個_Black', psu:'1000W_GOLD', os:'Windows11 Pro' }
  },
  {
    name: 'CAD・3D設計用',
    tags: 'CAD,3D,設計,建築,AutoCAD,SolidWorks,法人',
    budgetMin: 200000, budgetMax: 350000,
    description: 'AutoCADやSolidWorksが快適に動作。ISV認証GPUは要相談。',
    parts: { cpu:'Core i7 14700F', cooler:'240mm_Black', motherboard:'B760M_D4_ARGB', memory:'32GB_D4_BYL', gpu:'5070_Black', ssd:'2TB_SSD', case:'P30_M_Black', caseFan:'3個_Black', psu:'750W_GOLD', os:'Windows11 Pro' }
  },
];

async function fetchRecommendations() {
  MASTER.recommendations = [];

  // Try fetching from spreadsheet
  try {
    const res = await fetch(RECOMMEND_CSV_URL);
    if (!res.ok) throw new Error('Fetch failed');
    const text = await res.text();
    const rows = parseCSV(text);

    if (rows.length >= 2) {
      for (let i = 1; i < rows.length; i++) {
        const name = (rows[i][0] || '').trim();
        if (!name) continue;
        MASTER.recommendations.push({
          name,
          tags: (rows[i][1] || '').trim(),
          budgetMin: parseInt((rows[i][2] || '0').replace(/[,\s]/g, '')) || 0,
          budgetMax: parseInt((rows[i][3] || '0').replace(/[,\s]/g, '')) || 0,
          description: (rows[i][4] || '').trim(),
          parts: {
            cpu: (rows[i][5] || '').replace(/_$/, '').trim(),
            cooler: (rows[i][6] || '').replace(/_$/, '').trim(),
            motherboard: (rows[i][7] || '').replace(/_$/, '').trim(),
            memory: (rows[i][8] || '').replace(/_$/, '').trim(),
            gpu: (rows[i][9] || '').replace(/_$/, '').trim(),
            ssd: (rows[i][10] || '').replace(/_$/, '').trim(),
            case: (rows[i][11] || '').replace(/_$/, '').trim(),
            caseFan: (rows[i][12] || '').replace(/_$/, '').trim(),
            psu: (rows[i][13] || '').replace(/_$/, '').trim(),
            os: (rows[i][14] || '').replace(/_$/, '').trim(),
          }
        });
      }
      console.log(`Loaded ${MASTER.recommendations.length} recommendations from sheet`);
      return;
    }
  } catch (e) {
    console.warn('Recommendations sheet not found, using fallback data');
  }

  // Fallback
  MASTER.recommendations = FALLBACK_RECOMMENDATIONS;
  console.log('Using fallback recommendations');
}

// Search recommendations by keywords + budget
function searchRecommendations(query, budgetMin, budgetMax) {
  const recs = MASTER.recommendations || [];
  if (!query && !budgetMin && !budgetMax) return recs.slice(0, 3);

  // Split query into keywords
  const keywords = query.toLowerCase().split(/[\s,、　]+/).filter(Boolean);

  const scored = recs.map(r => {
    let score = 0;
    const tags = r.tags.toLowerCase();
    const desc = (r.description || '').toLowerCase();
    const name = r.name.toLowerCase();

    // Keyword match scoring
    keywords.forEach(kw => {
      if (tags.includes(kw)) score += 3;       // Tag match = strong
      if (name.includes(kw)) score += 2;       // Name match
      if (desc.includes(kw)) score += 1;       // Description match
    });

    // Budget match
    if (budgetMin || budgetMax) {
      const min = budgetMin || 0;
      const max = budgetMax || Infinity;
      if (r.budgetMax >= min && r.budgetMin <= max) {
        score += 2; // Budget overlap
      } else {
        score -= 5; // Out of budget penalty
      }
    }

    return { ...r, score };
  });

  return scored
    .filter(r => r.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 3);
}

// Calculate estimated price for a recommendation
function calcRecommendationPrice(rec) {
  let total = 0;
  const partKeys = ['cpu','cooler','motherboard','memory','gpu','ssd','case','caseFan','psu','os'];
  partKeys.forEach(key => {
    const partName = rec.parts[key];
    if (!partName) return;
    const source = key === 'os'
      ? [{ name:'Windows11 Home', price:15000 }, { name:'Windows11 Pro', price:22000 }]
      : (MASTER[key] || []);
    const found = source.find(item => item.name === partName);
    if (found) total += found.price;
  });
  return total;
}
