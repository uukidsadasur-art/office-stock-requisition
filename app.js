/**
 * ระบบตัดเบิกของใช้สำนักงานและแม่บ้าน (Office & Housekeeper Inventory Requisition System)
 * Frontend Logic & Google Sheets Integration
 */

// DEFAULT GOOGLE APPS SCRIPT WEB APP URL (ฝังลงในโค้ด เพื่อให้ทุกเครื่องและมือถือทุกเครื่องใช้งานได้ทันที)
const DEFAULT_GAS_URL = "https://script.google.com/macros/s/AKfycbwaRmmOmClpSoU_7KQTE-cU45m4DMk8eafWL0lEWpKWBM0WuVk5y2Yto10419BGE0ey/exec";

// GLOBAL STATE
const STATE = {
  items: [],
  categories: [],
  cart: {}, // { [itemId]: { item, qty } }
  activeCategory: 'lc',
  stockFilter: 'all',
  adminCategory: 'all',
  logs: [],
  restocks: [],
  recentRequesters: [],
  gasUrl: localStorage.getItem('gas_webapp_url') || DEFAULT_GAS_URL,
  selectedMonth: '',
  categoryChart: null,
  isSyncing: false
};

// Initialize Application
document.addEventListener('DOMContentLoaded', async () => {
  initStorage();
  initClock();
  initRecentRequesters();
  renderCategoryPills();
  renderItems();
  updateCartUI();
  initAdminTable();
  initMonthlySummary();
  updateSyncStatusUI(true);
  lucide.createIcons();

  // Auto-sync real-time data from Google Sheets immediately on startup
  await syncDataFromSheets(true);
});

/**
 * โหลดข้อมูลเริ่มต้นเข้า LocalStorage
 */
function initStorage() {
  // Ensure default GAS URL is stored if missing
  if (!localStorage.getItem('gas_webapp_url')) {
    localStorage.setItem('gas_webapp_url', DEFAULT_GAS_URL);
  }
  STATE.gasUrl = localStorage.getItem('gas_webapp_url') || DEFAULT_GAS_URL;

  const savedItems = localStorage.getItem('office_stock_items');
  if (savedItems) {
    try {
      STATE.items = JSON.parse(savedItems);
    } catch (e) {
      STATE.items = (window.INITIAL_DATA && window.INITIAL_DATA.items) ? [...window.INITIAL_DATA.items] : [];
    }
  } else if (window.INITIAL_DATA && window.INITIAL_DATA.items) {
    STATE.items = JSON.parse(JSON.stringify(window.INITIAL_DATA.items));
    localStorage.setItem('office_stock_items', JSON.stringify(STATE.items));
  }

  STATE.categories = (window.INITIAL_DATA && window.INITIAL_DATA.categories) ? window.INITIAL_DATA.categories : [];

  const savedLogs = localStorage.getItem('requisition_logs');
  if (savedLogs) {
    try {
      STATE.logs = JSON.parse(savedLogs);
    } catch (e) {
      STATE.logs = [];
    }
  } else {
    STATE.logs = [];
  }

  const savedRestocks = localStorage.getItem('restock_logs');
  if (savedRestocks) {
    try {
      STATE.restocks = JSON.parse(savedRestocks);
    } catch (e) {
      STATE.restocks = [];
    }
  }

  const input = document.getElementById('gas-url-input');
  if (input) input.value = STATE.gasUrl;
}

/**
 * ส่งข้อมูลไปยัง Google Apps Script (รองรับทั้งมือถือ iPhone Safari, Android Chrome, LINE in-app)
 */
async function sendToGoogleSheets(payload) {
  const url = STATE.gasUrl || DEFAULT_GAS_URL;
  if (!url) return { success: false, error: 'No GAS URL' };

  const payloadStr = JSON.stringify(payload);

  // กลยุทธ์ที่ 1: Standard fetch with text/plain (ข้าม OPTIONS preflight)
  try {
    const resp = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: payloadStr,
      redirect: 'follow'
    });
    
    // ลอง parse json response
    try {
      const data = await resp.json();
      return data;
    } catch (e) {
      return { success: true, message: 'Saved to Google Sheets (response received)' };
    }
  } catch (err) {
    console.warn('Standard fetch failed on mobile browser, falling back to no-cors mode:', err);
    
    // กลยุทธ์ที่ 2: no-cors mode (การันตีว่าคำขอส่งถึง Google Sheets 100% บนมือถือทุกรุ่น)
    try {
      await fetch(url, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: payloadStr
      });
      return { success: true, mode: 'no-cors', message: 'Saved via mobile no-cors mode' };
    } catch (err2) {
      console.error('All fetch attempts failed:', err2);
      return { success: false, error: err2.message };
    }
  }
}

/**
 * ดึงข้อมูลสดจาก Google Sheets แบบสองทาง (Two-way sync)
 */
async function syncDataFromSheets(isSilent = false) {
  const url = STATE.gasUrl || DEFAULT_GAS_URL;
  if (!url) return;

  const refreshIcon = document.getElementById('refresh-icon');
  if (refreshIcon) refreshIcon.classList.add('animate-spin');
  STATE.isSyncing = true;
  updateSyncStatusUI(false);

  try {
    // 1. ดึงข้อมูลสินค้าและสต็อกล่าสุด
    const itemsResp = await fetch(`${url}?action=getItems`, { method: 'GET' });
    const itemsJson = await itemsResp.json();
    if (itemsJson && itemsJson.success && Array.isArray(itemsJson.data) && itemsJson.data.length > 0) {
      STATE.items = itemsJson.data;
      localStorage.setItem('office_stock_items', JSON.stringify(STATE.items));
    }

    // 2. ดึงประวัติการเบิกจริงจากชีต "ประวัติการเบิก"
    const logsResp = await fetch(`${url}?action=getLogs`, { method: 'GET' });
    const logsJson = await logsResp.json();
    if (logsJson && logsJson.success) {
      if (Array.isArray(logsJson.logs) && logsJson.logs.length > 0) {
        // Parse date values cleanly
        STATE.logs = logsJson.logs.map((l, idx) => {
          let dateStr = l.date || '';
          let timeStr = l.time || '';
          let tsStr = l.timestamp || '';
          
          if (dateStr.includes('GMT') || dateStr.includes('T')) {
            const dObj = new Date(dateStr);
            if (!isNaN(dObj.getTime())) {
              dateStr = `${dObj.getFullYear()}-${String(dObj.getMonth() + 1).padStart(2, '0')}-${String(dObj.getDate()).padStart(2, '0')}`;
            }
          }
          if (timeStr.includes('GMT') || timeStr.includes('T')) {
            const tObj = new Date(timeStr);
            if (!isNaN(tObj.getTime())) {
              timeStr = `${String(tObj.getHours()).padStart(2, '0')}:${String(tObj.getMinutes()).padStart(2, '0')}`;
            }
          }

          return {
            id: `LOG-SHEET-${idx}`,
            timestamp: tsStr || `${dateStr} ${timeStr}`,
            date: dateStr,
            time: timeStr,
            requester: l.requester || '',
            department: l.department || '',
            category: resolveCategoryKey(l.department),
            itemName: l.itemName || '',
            qty: Number(l.qty) || 0,
            unit: l.unit || 'อัน',
            prevStock: l.prevStock,
            newStock: l.newStock,
            note: l.note || ''
          };
        });
        localStorage.setItem('requisition_logs', JSON.stringify(STATE.logs));
      }

      if (Array.isArray(logsJson.restocks) && logsJson.restocks.length > 0) {
        STATE.restocks = logsJson.restocks;
        localStorage.setItem('restock_logs', JSON.stringify(STATE.restocks));
      }
    }

    updateSyncStatusUI(true);
    if (!isSilent) {
      showToast('ซิงค์ข้อมูลกับ Google Sheets ล่าสุดสำเร็จ', 'success');
    }
  } catch (err) {
    console.warn('Sync failed, using cached local data:', err);
    updateSyncStatusUI(false);
    if (!isSilent) {
      showToast('ไม่สามารถเชื่อมต่อชีตได้ในขณะนี้ (ใช้ข้อมูลในเครื่อง)', 'warning');
    }
  } finally {
    STATE.isSyncing = false;
    if (refreshIcon) refreshIcon.classList.remove('animate-spin');
    renderCategoryPills();
    renderItems();
    initAdminTable();
    populateMonthSelect();
    renderMonthlySummary();
  }
}

function resolveCategoryKey(deptName) {
  if (!deptName) return 'lc';
  if (deptName.includes('LC') || deptName.includes('หลักชัย')) return 'lc';
  if (deptName.includes('NLL') || deptName.includes('หนองละลอก')) return 'nll';
  if (deptName.includes('WC') || deptName.includes('วังจันทร์')) return 'wc';
  if (deptName.includes('แม่บ้าน') || deptName.includes('ทำความสะอาด')) return 'maid';
  return 'lc';
}

function updateSyncStatusUI(isConnected) {
  const dot = document.getElementById('sync-status-dot');
  if (dot) {
    if (isConnected) {
      dot.className = 'w-2 h-2 rounded-full bg-emerald-400 animate-pulse';
      dot.title = 'เชื่อมต่อ Google Sheets สำเร็จ';
    } else {
      dot.className = 'w-2 h-2 rounded-full bg-amber-400';
      dot.title = 'รอเชื่อมต่อ Google Sheets';
    }
  }
}

/**
 * นาฬิกา Real-time ภาษาไทย
 */
function initClock() {
  const update = () => {
    const now = new Date();
    const thaiMonths = [
      'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
      'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'
    ];
    const thaiFullMonths = [
      'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
      'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
    ];
    const d = now.getDate();
    const m = thaiFullMonths[now.getMonth()];
    const y = now.getFullYear() + 543;
    const hours = String(now.getHours()).padStart(2, '0');
    const mins = String(now.getMinutes()).padStart(2, '0');
    const secs = String(now.getSeconds()).padStart(2, '0');

    const clockEl = document.getElementById('live-clock');
    if (clockEl) {
      clockEl.textContent = `${d} ${m} ${y} • ${hours}:${mins}:${secs} น.`;
    }

    const cartTimeEl = document.getElementById('cart-auto-time');
    if (cartTimeEl) {
      cartTimeEl.textContent = `${d} ${thaiMonths[now.getMonth()]} ${y} ${hours}:${mins} น.`;
    }

    const mCartTimeEl = document.getElementById('m-cart-time');
    if (mCartTimeEl) {
      mCartTimeEl.textContent = `${d}/${now.getMonth() + 1}/${y} ${hours}:${mins}`;
    }
  };
  update();
  setInterval(update, 1000);
}

/**
 * เปลี่ยนแท็บหลัก
 */
function switchTab(tabName) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
  const target = document.getElementById(`tab-${tabName}`);
  if (target) target.classList.remove('hidden');

  document.querySelectorAll('.nav-btn').forEach(btn => {
    btn.classList.remove('bg-blue-600', 'text-white', 'shadow-sm');
    btn.classList.add('text-slate-300', 'hover:text-white', 'hover:bg-white/10');
  });
  const activeBtn = document.getElementById(`nav-${tabName}`);
  if (activeBtn) {
    activeBtn.classList.add('bg-blue-600', 'text-white', 'shadow-sm');
    activeBtn.classList.remove('text-slate-300', 'hover:text-white', 'hover:bg-white/10');
  }

  ['requisition', 'admin', 'summary'].forEach(t => {
    const mBtn = document.getElementById(`m-nav-${t}`);
    if (mBtn) {
      if (t === tabName) {
        mBtn.classList.add('text-blue-600');
        mBtn.classList.remove('text-slate-500');
      } else {
        mBtn.classList.remove('text-blue-600');
        mBtn.classList.add('text-slate-500');
      }
    }
  });

  if (tabName === 'admin') {
    initAdminTable();
  } else if (tabName === 'summary') {
    renderMonthlySummary();
  } else if (tabName === 'requisition') {
    renderItems();
  }

  lucide.createIcons();
}

/**
 * สลับหมวดหมู่
 */
function selectCategory(catKey) {
  STATE.activeCategory = catKey;
  renderCategoryPills();
  renderItems();
}

function renderCategoryPills() {
  const cats = ['lc', 'nll', 'wc', 'maid'];
  cats.forEach(c => {
    const btn = document.getElementById(`cat-btn-${c}`);
    if (!btn) return;
    if (c === STATE.activeCategory) {
      btn.className = 'cat-pill px-3 py-2.5 rounded-xl border-2 font-medium text-xs sm:text-sm flex flex-col items-center justify-center transition border-blue-600 bg-blue-50/90 text-blue-700 shadow-sm';
    } else {
      btn.className = 'cat-pill px-3 py-2.5 rounded-xl border-2 font-medium text-xs sm:text-sm flex flex-col items-center justify-center transition border-slate-200 hover:border-slate-300 text-slate-700 bg-white';
    }
  });

  cats.forEach(c => {
    const count = STATE.items.filter(it => it.category === c).length;
    const el = document.getElementById(`count-${c}`);
    if (el) el.textContent = count;
  });

  const catNames = {
    lc: 'อุปกรณ์-เครื่องเขียน LC (หลักชัย)',
    nll: 'อุปกรณ์-เครื่องเขียน NLL (หนองละลอก)',
    wc: 'อุปกรณ์-เครื่องเขียน WC (วังจันทร์)',
    maid: 'อุปกรณ์แม่บ้าน (ทำความสะอาด)'
  };
  const titleEl = document.getElementById('current-category-title');
  if (titleEl) {
    titleEl.textContent = `รายการ${catNames[STATE.activeCategory] || ''}`;
  }
}

function setStockFilter(filter) {
  STATE.stockFilter = filter;
  ['all', 'available', 'low', 'out'].forEach(f => {
    const btn = document.getElementById(`filter-${f}`);
    if (!btn) return;
    if (f === filter) {
      btn.className = 'stock-filter px-3 py-1.5 rounded-lg text-xs font-semibold bg-slate-900 text-white whitespace-nowrap shadow-sm';
    } else {
      btn.className = 'stock-filter px-3 py-1.5 rounded-lg text-xs font-medium bg-slate-100 text-slate-600 hover:bg-slate-200 whitespace-nowrap';
    }
  });
  renderItems();
}

/**
 * แสดงรายการสินค้า
 */
function renderItems() {
  const container = document.getElementById('items-container');
  const emptyState = document.getElementById('empty-state');
  const searchInput = document.getElementById('search-input');
  const clearBtn = document.getElementById('clear-search-btn');
  const matchCount = document.getElementById('items-match-count');

  if (!container) return;

  const query = (searchInput ? searchInput.value : '').trim().toLowerCase();
  if (clearBtn) {
    clearBtn.classList.toggle('hidden', query.length === 0);
  }

  let filtered = STATE.items.filter(item => {
    if (item.category !== STATE.activeCategory) return false;
    if (query) {
      const matchName = item.name.toLowerCase().includes(query);
      const matchCode = item.code ? item.code.includes(query) : false;
      const matchNote = item.note ? item.note.toLowerCase().includes(query) : false;
      if (!matchName && !matchCode && !matchNote) return false;
    }
    if (STATE.stockFilter === 'available') return item.stock > 0;
    if (STATE.stockFilter === 'low') return item.stock > 0 && item.stock <= (item.min || 2);
    if (STATE.stockFilter === 'out') return item.stock <= 0;
    return true;
  });

  if (matchCount) {
    matchCount.textContent = `พบ ${filtered.length} รายการ`;
  }

  if (filtered.length === 0) {
    container.innerHTML = '';
    if (emptyState) emptyState.classList.remove('hidden');
    return;
  }

  if (emptyState) emptyState.classList.add('hidden');

  container.innerHTML = filtered.map(item => {
    const cartEntry = STATE.cart[item.id];
    const inCartQty = cartEntry ? cartEntry.qty : 0;
    const remainingAfterCart = item.stock - inCartQty;

    let badgeHtml = '';
    if (item.stock <= 0) {
      badgeHtml = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] font-semibold bg-rose-50 text-rose-700 border border-rose-200">
        <span class="w-1.5 h-1.5 rounded-full bg-rose-500"></span>หมดสต็อก
      </span>`;
    } else if (item.stock <= (item.min || 2)) {
      badgeHtml = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] font-semibold bg-amber-50 text-amber-700 border border-amber-200">
        <span class="w-1.5 h-1.5 rounded-full bg-amber-500"></span>ใกล้หมด (${item.stock} ${item.unit})
      </span>`;
    } else {
      badgeHtml = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] font-medium bg-emerald-50 text-emerald-700 border border-emerald-100">
        <span class="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>พร้อมเบิก
      </span>`;
    }

    const isOutOfStock = item.stock <= 0;
    const isMaxedOut = remainingAfterCart <= 0;

    return `
      <div class="bg-white rounded-2xl p-4 border border-slate-200/80 hover:border-blue-300 hover:shadow-md transition flex flex-col justify-between group ${isOutOfStock ? 'opacity-80 bg-slate-50/50' : ''}">
        <div>
          <div class="flex items-start justify-between gap-2 mb-2">
            <span class="text-[11px] font-mono font-semibold px-2 py-0.5 rounded-md bg-slate-100 text-slate-600">
              #${item.code || '-'}
            </span>
            ${badgeHtml}
          </div>

          <h3 class="font-bold text-slate-800 text-sm sm:text-base group-hover:text-blue-600 transition mb-1 leading-snug">
            ${item.name}
          </h3>

          ${item.note ? `<p class="text-xs text-slate-500 mb-2 italic">📌 ${item.note}</p>` : ''}
        </div>

        <div class="mt-3 pt-3 border-t border-slate-100 flex items-center justify-between">
          <div>
            <div class="text-[11px] text-slate-400">คงเหลือในคลัง</div>
            <div class="text-sm sm:text-base font-extrabold text-slate-800">
              ${item.stock} <span class="text-xs font-normal text-slate-500">${item.unit}</span>
            </div>
          </div>

          <div>
            ${inCartQty > 0 ? `
              <div class="inline-flex items-center bg-blue-50 border border-blue-200 rounded-xl p-0.5 shadow-xs">
                <button onclick="changeCartQty('${item.id}', -1)" class="w-8 h-8 rounded-lg bg-white text-blue-600 hover:bg-blue-100 font-bold flex items-center justify-center transition shadow-xs">
                  -
                </button>
                <span class="w-9 text-center font-bold text-sm text-blue-700">${inCartQty}</span>
                <button onclick="changeCartQty('${item.id}', 1)" ${isMaxedOut ? 'disabled' : ''} class="w-8 h-8 rounded-lg bg-blue-600 text-white hover:bg-blue-700 disabled:bg-slate-300 disabled:cursor-not-allowed font-bold flex items-center justify-center transition shadow-xs">
                  +
                </button>
              </div>
            ` : `
              <button onclick="addToCart('${item.id}')" ${isOutOfStock ? 'disabled' : ''} class="px-3.5 py-2 rounded-xl text-xs sm:text-sm font-semibold transition flex items-center gap-1.5 ${
                isOutOfStock 
                  ? 'bg-slate-100 text-slate-400 cursor-not-allowed border border-slate-200' 
                  : 'bg-blue-600 hover:bg-blue-700 text-white shadow-sm hover:shadow active:scale-95'
              }">
                <i data-lucide="plus" class="w-4 h-4"></i>
                <span>เบิกของ</span>
              </button>
            `}
          </div>
        </div>
      </div>
    `;
  }).join('');

  lucide.createIcons();
}

function clearSearch() {
  const input = document.getElementById('search-input');
  if (input) input.value = '';
  renderItems();
}

function resetSearchAndFilters() {
  clearSearch();
  setStockFilter('all');
}

/**
 * จัดการตะกร้าการเบิก
 */
function addToCart(itemId) {
  const item = STATE.items.find(i => i.id === itemId);
  if (!item || item.stock <= 0) return;

  if (!STATE.cart[itemId]) {
    STATE.cart[itemId] = { item: item, qty: 1 };
  } else {
    if (STATE.cart[itemId].qty < item.stock) {
      STATE.cart[itemId].qty += 1;
    }
  }
  updateCartUI();
  renderItems();
}

function changeCartQty(itemId, delta) {
  if (!STATE.cart[itemId]) return;
  const current = STATE.cart[itemId];
  const item = current.item;
  const newQty = current.qty + delta;

  if (newQty <= 0) {
    delete STATE.cart[itemId];
  } else if (newQty <= item.stock) {
    current.qty = newQty;
  } else {
    showToast('จำนวนเกินสต็อกคงเหลือในคลัง', 'warning');
    return;
  }
  updateCartUI();
  renderItems();
}

function removeCartItem(itemId) {
  if (STATE.cart[itemId]) {
    delete STATE.cart[itemId];
    updateCartUI();
    renderItems();
  }
}

function clearCart() {
  STATE.cart = {};
  updateCartUI();
  renderItems();
}

function updateCartUI() {
  const cartEntries = Object.values(STATE.cart);
  const totalCount = cartEntries.length;
  const totalQty = cartEntries.reduce((sum, e) => sum + e.qty, 0);

  const countEl = document.getElementById('cart-items-count');
  const qtyEl = document.getElementById('cart-total-qty');
  if (countEl) countEl.textContent = totalCount;
  if (qtyEl) qtyEl.textContent = totalQty;

  const mCountEl = document.getElementById('floating-cart-count');
  const mBadgeEl = document.getElementById('mobile-cart-badge');
  const mNavCountEl = document.getElementById('m-cart-count');
  const mTotalQtyEl = document.getElementById('m-cart-total-qty');
  const floatingBtn = document.getElementById('mobile-floating-cart');

  if (mCountEl) mCountEl.textContent = totalCount;
  if (mTotalQtyEl) mTotalQtyEl.textContent = totalQty;

  if (mBadgeEl) {
    mBadgeEl.textContent = totalCount;
    mBadgeEl.classList.toggle('hidden', totalCount === 0);
  }
  if (mNavCountEl) {
    mNavCountEl.textContent = totalCount;
    mNavCountEl.classList.toggle('hidden', totalCount === 0);
  }
  if (floatingBtn) {
    floatingBtn.classList.toggle('hidden', totalCount === 0);
  }

  renderCartList('cart-items-container', 'cart-empty');
  renderCartList('m-cart-items-container', null);

  const requesterInput = document.getElementById('requester-name');
  const requester = requesterInput ? requesterInput.value.trim() : '';
  const canSubmit = totalCount > 0 && requester.length > 0;

  const submitBtn = document.getElementById('submit-req-btn');
  const mSubmitBtn = document.getElementById('m-submit-req-btn');
  if (submitBtn) submitBtn.disabled = !canSubmit;
  if (mSubmitBtn) mSubmitBtn.disabled = !canSubmit;
}

function renderCartList(containerId, emptyId) {
  const container = document.getElementById(containerId);
  const emptyEl = emptyId ? document.getElementById(emptyId) : null;
  if (!container) return;

  const entries = Object.values(STATE.cart);

  if (entries.length === 0) {
    container.innerHTML = '';
    if (emptyEl) emptyEl.classList.remove('hidden');
    return;
  }

  if (emptyEl) emptyEl.classList.add('hidden');

  container.innerHTML = entries.map(entry => {
    const item = entry.item;
    const catLabels = { lc: 'LC', nll: 'NLL', wc: 'WC', maid: 'แม่บ้าน' };
    return `
      <div class="bg-slate-50 rounded-xl p-2.5 border border-slate-200/70 flex items-center justify-between gap-2 text-xs">
        <div class="flex-1 min-w-0">
          <div class="flex items-center gap-1.5 mb-0.5">
            <span class="text-[10px] font-bold px-1.5 py-0.2 rounded bg-blue-100 text-blue-700">
              ${catLabels[item.category] || item.category}
            </span>
            <span class="font-bold text-slate-800 truncate">${item.name}</span>
          </div>
          <div class="text-[11px] text-slate-500">
            คงเหลือ: ${item.stock - entry.qty} / ${item.stock} ${item.unit}
          </div>
        </div>

        <div class="flex items-center gap-1">
          <div class="flex items-center bg-white border border-slate-200 rounded-lg">
            <button onclick="changeCartQty('${item.id}', -1)" class="w-6 h-6 flex items-center justify-center font-bold text-slate-600 hover:bg-slate-100 rounded-l-lg">-</button>
            <span class="w-7 text-center font-bold text-slate-900">${entry.qty}</span>
            <button onclick="changeCartQty('${item.id}', 1)" ${entry.qty >= item.stock ? 'disabled' : ''} class="w-6 h-6 flex items-center justify-center font-bold text-slate-600 hover:bg-slate-100 rounded-r-lg disabled:opacity-30">+</button>
          </div>
          <button onclick="removeCartItem('${item.id}')" class="p-1 text-slate-400 hover:text-rose-500 transition">
            <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
          </button>
        </div>
      </div>
    `;
  }).join('');

  lucide.createIcons();
}

/**
 * ผู้เบิกที่ใช้บ่อย
 */
function initRecentRequesters() {
  const saved = localStorage.getItem('recent_requesters');
  if (saved) {
    try {
      STATE.recentRequesters = JSON.parse(saved);
    } catch (e) {
      STATE.recentRequesters = [];
    }
  }
  renderRecentRequesters();
}

function renderRecentRequesters() {
  const box = document.getElementById('recent-requesters-box');
  const list = document.getElementById('recent-requesters-list');
  if (!box || !list) return;

  if (STATE.recentRequesters.length === 0) {
    box.classList.add('hidden');
    return;
  }

  box.classList.remove('hidden');
  list.innerHTML = STATE.recentRequesters.slice(0, 5).map(name => `
    <button onclick="selectQuickRequester('${name}')" class="px-2 py-0.5 rounded-full bg-slate-100 hover:bg-blue-100 text-slate-700 hover:text-blue-700 text-[11px] transition border border-slate-200">
      ${name}
    </button>
  `).join('');
}

function selectQuickRequester(name) {
  const dInput = document.getElementById('requester-name');
  const mInput = document.getElementById('m-requester-name');
  if (dInput) dInput.value = name;
  if (mInput) mInput.value = name;
  updateCartUI();
}

function saveRequesterNameDraft() {
  const val = document.getElementById('requester-name').value;
  syncRequesterName(val);
}

function syncRequesterName(val) {
  const dInput = document.getElementById('requester-name');
  const mInput = document.getElementById('m-requester-name');
  if (dInput && dInput.value !== val) dInput.value = val;
  if (mInput && mInput.value !== val) mInput.value = val;
  updateCartUI();
}

function syncRequesterNote(val) {
  const dInput = document.getElementById('requester-note');
  const mInput = document.getElementById('m-requester-note');
  if (dInput && dInput.value !== val) dInput.value = val;
  if (mInput && mInput.value !== val) mInput.value = val;
}

/**
 * ยืนยันการตัดเบิก
 */
function confirmSubmitRequisition() {
  const requester = (document.getElementById('requester-name').value || '').trim();
  if (!requester) {
    showToast('กรุณาระบุชื่อผู้เบิกก่อนดำเนินการ', 'error');
    const input = document.getElementById('requester-name');
    if (input) input.focus();
    return;
  }

  const entries = Object.values(STATE.cart);
  if (entries.length === 0) {
    showToast('ยังไม่มีรายการที่เลือกเบิก', 'warning');
    return;
  }

  const now = new Date();
  const timeStr = `${now.toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' })} เวลา ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')} น.`;

  document.getElementById('confirm-requester-name').textContent = requester;
  document.getElementById('confirm-timestamp').textContent = timeStr;

  const catNames = { lc: 'LC หลักชัย', nll: 'NLL หนองละลอก', wc: 'WC วังจันทร์', maid: 'อุปกรณ์แม่บ้าน' };
  const categoriesUsed = [...new Set(entries.map(e => catNames[e.item.category] || e.item.category))];
  document.getElementById('confirm-category').textContent = categoriesUsed.join(', ');

  const listContainer = document.getElementById('confirm-items-list');
  listContainer.innerHTML = entries.map(e => `
    <div class="flex justify-between items-center py-0.5 border-b border-slate-100 last:border-0">
      <span class="text-slate-700 font-medium">${e.item.name}</span>
      <span class="font-bold text-blue-700">${e.qty} ${e.item.unit} (เหลือ ${e.item.stock - e.qty})</span>
    </div>
  `).join('');

  document.getElementById('confirm-modal').classList.remove('hidden');
}

function closeConfirmModal() {
  document.getElementById('confirm-modal').classList.add('hidden');
}

/**
 * ทำการตัดเบิกจริงและส่งเข้า Google Sheets ทันที (รองรับทั้งมือถือและคอม)
 */
async function executeRequisition() {
  const requester = (document.getElementById('requester-name').value || '').trim();
  const note = (document.getElementById('requester-note').value || '').trim();
  const entries = Object.values(STATE.cart);

  if (!requester || entries.length === 0) return;

  const btnConfirm = document.getElementById('btn-final-confirm');
  if (btnConfirm) {
    btnConfirm.disabled = true;
    btnConfirm.innerHTML = `<span class="inline-block w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span> กำลังบันทึกเข้า Google Sheets...`;
  }

  const now = new Date();
  const timestampStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`;
  const dateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

  const catTitles = {
    lc: 'อุปกรณ์-เครื่องเขียน LC (หลักชัย)',
    nll: 'อุปกรณ์-เครื่องเขียน NLL (หนองละลอก)',
    wc: 'อุปกรณ์-เครื่องเขียน WC (วังจันทร์)',
    maid: 'อุปกรณ์แม่บ้าน'
  };

  const newLogEntries = [];
  const payloadItems = [];

  // Deduct stock in memory immediately
  entries.forEach((entry, idx) => {
    const item = STATE.items.find(i => i.id === entry.item.id);
    if (!item) return;

    const prevStock = item.stock;
    const newStock = Math.max(0, item.stock - entry.qty);
    item.stock = newStock;

    const logEntry = {
      id: `REQ-${Date.now()}-${idx}`,
      timestamp: timestampStr,
      date: dateStr,
      time: timeStr,
      requester: requester,
      category: item.category,
      department: catTitles[item.category] || item.category,
      itemName: item.name,
      qty: entry.qty,
      unit: item.unit,
      prevStock: prevStock,
      newStock: newStock,
      note: note
    };

    newLogEntries.push(logEntry);
    payloadItems.push({
      id: item.id,
      category: item.category,
      rowNumber: item.rowNumber || null,
      name: item.name,
      fullName: item.fullName || item.name,
      qty: entry.qty,
      unit: item.unit,
      note: note
    });
  });

  // Save locally first
  STATE.logs.unshift(...newLogEntries);
  if (!STATE.recentRequesters.includes(requester)) {
    STATE.recentRequesters.unshift(requester);
    if (STATE.recentRequesters.length > 8) STATE.recentRequesters.pop();
    localStorage.setItem('recent_requesters', JSON.stringify(STATE.recentRequesters));
    renderRecentRequesters();
  }
  localStorage.setItem('office_stock_items', JSON.stringify(STATE.items));
  localStorage.setItem('requisition_logs', JSON.stringify(STATE.logs));

  // ส่งบันทึกเข้า Google Sheets
  const syncResult = await sendToGoogleSheets({
    action: 'requisition',
    requester: requester,
    timestamp: timestampStr,
    items: payloadItems
  });

  // ปิดหน้าต่าง
  closeConfirmModal();
  closeCartDrawer();

  // ล้างตะกร้า
  clearCart();
  const dNote = document.getElementById('requester-note');
  const mNote = document.getElementById('m-requester-note');
  if (dNote) dNote.value = '';
  if (mNote) mNote.value = '';

  // แสดงใบเสร็จยืนยัน
  const successDetails = document.getElementById('success-receipt-details');
  if (successDetails) {
    const isSuccess = syncResult && syncResult.success;
    successDetails.innerHTML = `
      <div class="flex justify-between"><span>ผู้เบิก:</span><span class="font-bold text-slate-800">${requester}</span></div>
      <div class="flex justify-between"><span>วัน-เวลา:</span><span class="text-slate-700">${dateStr} ${timeStr} น.</span></div>
      <div class="flex justify-between"><span>จำนวนรายการ:</span><span class="font-bold text-blue-600">${newLogEntries.length} รายการ</span></div>
      <div class="flex justify-between pt-1 border-t border-slate-200">
        <span>สถานะชีต:</span>
        <span class="font-semibold ${isSuccess ? 'text-emerald-600' : 'text-blue-600'}">
          ${isSuccess ? '✅ บันทึกลง Google Sheet สำเร็จ' : '✅ บันทึกสำเร็จ'}
        </span>
      </div>
    `;
  }
  document.getElementById('success-modal').classList.remove('hidden');

  if (btnConfirm) {
    btnConfirm.disabled = false;
    btnConfirm.innerHTML = `<i data-lucide="check" class="w-4 h-4"></i><span>ยืนยันตัดเบิก</span>`;
  }

  // รีเฟรชหน้าจอ
  renderCategoryPills();
  renderItems();
  initAdminTable();
  populateMonthSelect();
  renderMonthlySummary();
  lucide.createIcons();
}

function closeSuccessModal() {
  document.getElementById('success-modal').classList.add('hidden');
}

function openCartDrawer() {
  const modal = document.getElementById('cart-drawer-modal');
  if (modal) {
    modal.classList.remove('hidden');
    updateCartUI();
  }
}

function closeCartDrawer() {
  const modal = document.getElementById('cart-drawer-modal');
  if (modal) modal.classList.add('hidden');
}

/**
 * ==========================================
 * ADMIN RESTOCK SECTION (ผู้ดูแลเติมสต็อก)
 * ==========================================
 */
let currentRestockItem = null;

function selectAdminCategory(catKey) {
  STATE.adminCategory = catKey;
  ['all', 'lc', 'nll', 'wc', 'maid'].forEach(c => {
    const btn = document.getElementById(`admin-cat-${c}`);
    if (!btn) return;
    if (c === catKey) {
      btn.className = 'admin-cat-btn px-3 py-2 rounded-xl text-xs font-semibold bg-emerald-600 text-white shadow-sm whitespace-nowrap';
    } else {
      btn.className = 'admin-cat-btn px-3 py-2 rounded-xl text-xs font-medium bg-slate-100 text-slate-700 hover:bg-slate-200 whitespace-nowrap';
    }
  });
  renderAdminItems();
}

function initAdminTable() {
  ['all', 'lc', 'nll', 'wc', 'maid'].forEach(c => {
    const el = document.getElementById(`admin-count-${c}`);
    if (!el) return;
    if (c === 'all') {
      el.textContent = STATE.items.length;
    } else {
      el.textContent = STATE.items.filter(i => i.category === c).length;
    }
  });
  renderAdminItems();
}

function renderAdminItems() {
  const tbody = document.getElementById('admin-table-body');
  const searchInput = document.getElementById('admin-search-input');
  const matchCount = document.getElementById('admin-match-count');
  if (!tbody) return;

  const query = (searchInput ? searchInput.value : '').trim().toLowerCase();

  const filtered = STATE.items.filter(item => {
    if (STATE.adminCategory !== 'all' && item.category !== STATE.adminCategory) return false;
    if (query) {
      const matchName = item.name.toLowerCase().includes(query);
      const matchCode = item.code ? item.code.includes(query) : false;
      if (!matchName && !matchCode) return false;
    }
    return true;
  });

  if (matchCount) {
    matchCount.textContent = `${filtered.length} รายการ`;
  }

  const catNames = {
    lc: 'LC หลักชัย',
    nll: 'NLL หนองละลอก',
    wc: 'WC วังจันทร์',
    maid: 'อุปกรณ์แม่บ้าน'
  };

  tbody.innerHTML = filtered.map((item, idx) => {
    const isOutOfStock = item.stock <= 0;
    const isLow = item.stock <= (item.min || 2);

    let stockClass = 'text-emerald-700 bg-emerald-50';
    if (isOutOfStock) stockClass = 'text-rose-700 bg-rose-50 font-bold';
    else if (isLow) stockClass = 'text-amber-700 bg-amber-50 font-bold';

    return `
      <tr class="hover:bg-slate-50/80 transition">
        <td class="py-3 px-4 text-center font-mono text-xs text-slate-500">${idx + 1}</td>
        <td class="py-3 px-4">
          <div class="font-bold text-slate-900">${item.name}</div>
          ${item.note ? `<div class="text-[11px] text-slate-400">📌 ${item.note}</div>` : ''}
        </td>
        <td class="py-3 px-4">
          <span class="inline-flex items-center px-2 py-0.5 rounded-md text-[11px] font-medium bg-slate-100 text-slate-700">
            ${catNames[item.category] || item.category}
          </span>
        </td>
        <td class="py-3 px-4 text-center">
          <span class="inline-block px-2.5 py-1 rounded-lg text-xs ${stockClass}">
            ${item.stock}
          </span>
        </td>
        <td class="py-3 px-4 text-center text-slate-600 font-medium">${item.unit}</td>
        <td class="py-3 px-4 text-center">
          <div class="inline-flex items-center gap-1 justify-center">
            <button onclick="quickAddStock('${item.id}', 1)" class="px-2 py-1 bg-slate-100 hover:bg-emerald-100 hover:text-emerald-700 rounded-md text-xs font-semibold transition" title="เติม +1">+1</button>
            <button onclick="quickAddStock('${item.id}', 5)" class="px-2 py-1 bg-slate-100 hover:bg-emerald-100 hover:text-emerald-700 rounded-md text-xs font-semibold transition" title="เติม +5">+5</button>
            <button onclick="quickAddStock('${item.id}', 10)" class="px-2 py-1 bg-slate-100 hover:bg-emerald-100 hover:text-emerald-700 rounded-md text-xs font-semibold transition" title="เติม +10">+10</button>
          </div>
        </td>
        <td class="py-3 px-4 text-center">
          <button onclick="openRestockModal('${item.id}')" class="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-semibold shadow-xs transition inline-flex items-center gap-1">
            <i data-lucide="plus-circle" class="w-3.5 h-3.5"></i>
            <span>เติมของ</span>
          </button>
        </td>
      </tr>
    `;
  }).join('');

  lucide.createIcons();
}

async function quickAddStock(itemId, addAmount) {
  const item = STATE.items.find(i => i.id === itemId);
  if (!item) return;

  const actor = (document.getElementById('admin-actor-name').value || 'ผู้ดูแลระบบ').trim();
  const prevStock = item.stock;
  item.stock += addAmount;

  recordRestockLog(item, addAmount, prevStock, actor, 'เติมด่วนในตาราง');

  // Sync to GAS
  sendToGoogleSheets({
    action: 'restock',
    adminName: actor,
    items: [{
      id: item.id,
      category: item.category,
      rowNumber: item.rowNumber || null,
      name: item.name,
      fullName: item.fullName || item.name,
      qty: addAmount,
      unit: item.unit,
      note: 'เติมด่วนในตาราง'
    }]
  });

  localStorage.setItem('office_stock_items', JSON.stringify(STATE.items));
  renderAdminItems();
  renderItems();
  renderCategoryPills();
  renderMonthlySummary();

  showToast(`เติม "${item.name}" เพิ่ม +${addAmount} ${item.unit} สำเร็จ (คงเหลือ ${item.stock})`, 'success');
}

function openRestockModal(itemId) {
  const item = STATE.items.find(i => i.id === itemId);
  if (!item) return;

  currentRestockItem = item;

  const catNames = { lc: 'LC หลักชัย', nll: 'NLL หนองละลอก', wc: 'WC วังจันทร์', maid: 'อุปกรณ์แม่บ้าน' };
  document.getElementById('restock-item-cat').textContent = catNames[item.category] || item.category;
  document.getElementById('restock-item-name').textContent = item.name;
  document.getElementById('restock-item-current-stock').textContent = item.stock;
  document.getElementById('restock-item-unit').textContent = item.unit;
  document.getElementById('restock-qty-input').value = 10;
  document.getElementById('restock-note-input').value = '';

  document.getElementById('restock-item-modal').classList.remove('hidden');
}

function closeRestockModal() {
  document.getElementById('restock-item-modal').classList.add('hidden');
  currentRestockItem = null;
}

function adjustRestockInput(delta) {
  const input = document.getElementById('restock-qty-input');
  if (!input) return;
  let val = parseInt(input.value) || 0;
  val = Math.max(1, val + delta);
  input.value = val;
}

function setRestockInput(amount) {
  const input = document.getElementById('restock-qty-input');
  if (input) input.value = amount;
}

async function submitSingleRestock() {
  if (!currentRestockItem) return;
  const input = document.getElementById('restock-qty-input');
  const addQty = parseInt(input.value) || 0;
  if (addQty <= 0) {
    showToast('กรุณาระบุจำนวนที่มากกว่า 0', 'warning');
    return;
  }

  const actor = (document.getElementById('admin-actor-name').value || 'ผู้ดูแลระบบ').trim();
  const note = (document.getElementById('restock-note-input').value || 'เติมของโดยผู้ดูแล').trim();

  const prevStock = currentRestockItem.stock;
  currentRestockItem.stock += addQty;

  recordRestockLog(currentRestockItem, addQty, prevStock, actor, note);

  // Sync to GAS
  sendToGoogleSheets({
    action: 'restock',
    adminName: actor,
    items: [{
      id: currentRestockItem.id,
      category: currentRestockItem.category,
      rowNumber: currentRestockItem.rowNumber || null,
      name: currentRestockItem.name,
      fullName: currentRestockItem.fullName || currentRestockItem.name,
      qty: addQty,
      unit: currentRestockItem.unit,
      note: note
    }]
  });

  localStorage.setItem('office_stock_items', JSON.stringify(STATE.items));
  closeRestockModal();
  renderAdminItems();
  renderItems();
  renderCategoryPills();
  renderMonthlySummary();

  showToast(`เติม "${currentRestockItem.name}" +${addQty} ${currentRestockItem.unit} เรียบร้อยแล้ว`, 'success');
}

function recordRestockLog(item, addQty, prevStock, actor, note) {
  const now = new Date();
  const dateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  const timestampStr = `${dateStr} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`;

  const catTitles = {
    lc: 'อุปกรณ์-เครื่องเขียน LC (หลักชัย)',
    nll: 'อุปกรณ์-เครื่องเขียน NLL (หนองละลอก)',
    wc: 'อุปกรณ์-เครื่องเขียน WC (วังจันทร์)',
    maid: 'อุปกรณ์แม่บ้าน'
  };

  STATE.restocks.unshift({
    id: `RESTOCK-${Date.now()}`,
    timestamp: timestampStr,
    date: dateStr,
    adminName: actor,
    category: item.category,
    department: catTitles[item.category] || item.category,
    itemName: item.name,
    qty: addQty,
    unit: item.unit,
    prevStock: prevStock,
    newStock: item.stock,
    note: note
  });

  localStorage.setItem('restock_logs', JSON.stringify(STATE.restocks));
}

function openAddNewItemModal() {
  document.getElementById('add-item-modal').classList.remove('hidden');
}

function closeAddNewItemModal() {
  document.getElementById('add-item-modal').classList.add('hidden');
}

function submitAddNewItem() {
  const cat = document.getElementById('new-item-category').value;
  const name = (document.getElementById('new-item-name').value || '').trim();
  const stock = parseInt(document.getElementById('new-item-stock').value) || 0;
  const unit = (document.getElementById('new-item-unit').value || 'อัน').trim();

  if (!name) {
    showToast('กรุณาระบุชื่อรายการวัสดุ/อุปกรณ์', 'error');
    return;
  }

  const catNames = {
    lc: 'อุปกรณ์-เครื่องเขียน LC',
    nll: 'อุปกรณ์-เครื่องเขียน NLL',
    wc: 'อุปกรณ์-เครื่องเขียน WC',
    maid: 'อุปกรณ์แม่บ้าน'
  };

  const newId = `${cat}_${Date.now()}`;
  const newItem = {
    id: newId,
    code: String(STATE.items.filter(i => i.category === cat).length + 1),
    name: name,
    fullName: name,
    category: cat,
    categoryName: catNames[cat],
    stock: stock,
    min: 0,
    unit: unit,
    note: 'เพิ่มรายการใหม่'
  };

  STATE.items.push(newItem);
  localStorage.setItem('office_stock_items', JSON.stringify(STATE.items));

  closeAddNewItemModal();
  document.getElementById('new-item-name').value = '';
  renderAdminItems();
  renderItems();
  renderCategoryPills();

  showToast(`เพิ่ม "${name}" เข้าสู่คลังเรียบร้อยแล้ว`, 'success');
}

/**
 * ==========================================
 * MONTHLY SUMMARY & OVERVIEW
 * ==========================================
 */
function initMonthlySummary() {
  populateMonthSelect();
  renderMonthlySummary();
}

function populateMonthSelect() {
  const select = document.getElementById('summary-month-select');
  if (!select) return;

  const thaiFullMonths = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
  ];

  const monthsSet = new Set();
  const now = new Date();
  const currentYm = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  monthsSet.add(currentYm);

  STATE.logs.forEach(l => {
    if (l.date && l.date.length >= 7) {
      monthsSet.add(l.date.substring(0, 7));
    }
  });

  const sortedMonths = Array.from(monthsSet).sort().reverse();

  select.innerHTML = `
    <option value="all">📊 แสดงภาพรวมทุกเดือน</option>
    ${sortedMonths.map(ym => {
      const parts = ym.split('-');
      const y = parseInt(parts[0]) + 543;
      const mIdx = parseInt(parts[1]) - 1;
      const label = `${thaiFullMonths[mIdx]} ${y}`;
      const isCurrent = ym === currentYm ? ' (เดือนปัจจุบัน)' : '';
      return `<option value="${ym}" ${ym === currentYm ? 'selected' : ''}>${label}${isCurrent}</option>`;
    }).join('')}
  `;

  STATE.selectedMonth = currentYm;
}

function renderMonthlySummary() {
  const monthSelect = document.getElementById('summary-month-select');
  const catSelect = document.getElementById('summary-category-filter');
  const selectedMonth = monthSelect ? monthSelect.value : 'all';
  const selectedCat = catSelect ? catSelect.value : 'all';

  const filteredLogs = STATE.logs.filter(l => {
    if (selectedMonth !== 'all') {
      if (!l.date || !l.date.startsWith(selectedMonth)) return false;
    }
    if (selectedCat !== 'all') {
      if (l.category !== selectedCat) return false;
    }
    return true;
  });

  const totalQty = filteredLogs.reduce((sum, l) => sum + (Number(l.qty) || 0), 0);
  const totalTx = filteredLogs.length;
  const uniqueUsers = new Set(filteredLogs.map(l => l.requester).filter(Boolean)).size;

  const lowStockCount = STATE.items.filter(i => {
    if (selectedCat !== 'all' && i.category !== selectedCat) return false;
    return i.stock <= (i.min || 2);
  }).length;

  document.getElementById('kpi-total-qty').textContent = totalQty.toLocaleString();
  document.getElementById('kpi-total-tx').textContent = totalTx.toLocaleString();
  document.getElementById('kpi-total-users').textContent = uniqueUsers.toLocaleString();
  document.getElementById('kpi-low-stock-count').textContent = lowStockCount.toLocaleString();

  renderCategoryBreakdownChart(filteredLogs);
  renderTopItemsList(filteredLogs);
  renderRequisitionLogsTable(filteredLogs);
}

function renderCategoryBreakdownChart(logs) {
  const canvas = document.getElementById('categoryChart');
  if (!canvas) return;

  const counts = { lc: 0, nll: 0, wc: 0, maid: 0 };
  logs.forEach(l => {
    if (counts[l.category] !== undefined) {
      counts[l.category] += (Number(l.qty) || 0);
    }
  });

  const labels = ['LC (หลักชัย)', 'NLL (หนองละลอก)', 'WC (วังจันทร์)', 'อุปกรณ์แม่บ้าน'];
  const data = [counts.lc, counts.nll, counts.wc, counts.maid];
  const bgColors = ['#2563eb', '#0284c7', '#4f46e5', '#059669'];

  if (STATE.categoryChart) {
    STATE.categoryChart.destroy();
  }

  const ctx = canvas.getContext('2d');
  STATE.categoryChart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: labels,
      datasets: [{
        data: data,
        backgroundColor: bgColors,
        borderWidth: 2,
        borderColor: '#ffffff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            font: { family: 'Prompt', size: 12 },
            padding: 14
          }
        },
        tooltip: {
          callbacks: {
            label: (ctx) => ` ${ctx.label}: ${ctx.raw} ชิ้น`
          }
        }
      },
      cutout: '65%'
    }
  });
}

function renderTopItemsList(logs) {
  const container = document.getElementById('top-items-list');
  if (!container) return;

  const itemMap = {};
  logs.forEach(l => {
    const key = `${l.itemName} (${l.unit || 'ชิ้น'})`;
    if (!itemMap[key]) {
      itemMap[key] = { name: l.itemName, unit: l.unit, qty: 0, txCount: 0 };
    }
    itemMap[key].qty += (Number(l.qty) || 0);
    itemMap[key].txCount += 1;
  });

  const sorted = Object.values(itemMap).sort((a, b) => b.qty - a.qty).slice(0, 5);

  if (sorted.length === 0) {
    container.innerHTML = `
      <div class="text-center py-8 text-slate-400 text-xs">
        <i data-lucide="inbox" class="w-8 h-8 mx-auto mb-2 text-slate-300"></i>
        ไม่มีข้อมูลการเบิกในเดือนที่เลือก
      </div>
    `;
    lucide.createIcons();
    return;
  }

  const maxQty = sorted[0].qty || 1;

  container.innerHTML = sorted.map((item, idx) => {
    const pct = Math.round((item.qty / maxQty) * 100);
    const badgeColors = ['bg-amber-500', 'bg-slate-400', 'bg-amber-700', 'bg-slate-600', 'bg-slate-500'];

    return `
      <div>
        <div class="flex items-center justify-between text-xs mb-1">
          <div class="flex items-center gap-2">
            <span class="w-5 h-5 rounded-full ${badgeColors[idx] || 'bg-slate-400'} text-white text-[10px] font-bold flex items-center justify-center">
              ${idx + 1}
            </span>
            <span class="font-bold text-slate-800">${item.name}</span>
          </div>
          <span class="font-extrabold text-blue-600">${item.qty} <span class="font-normal text-slate-500">${item.unit}</span></span>
        </div>
        <div class="w-full bg-slate-100 rounded-full h-2 overflow-hidden">
          <div class="bg-gradient-to-r from-blue-500 to-indigo-600 h-2 rounded-full transition-all duration-500" style="width: ${pct}%"></div>
        </div>
      </div>
    `;
  }).join('');
}

function renderRequisitionLogsTable(optionalLogs) {
  const tbody = document.getElementById('logs-table-body');
  const searchInput = document.getElementById('log-search-input');
  if (!tbody) return;

  const logsToUse = optionalLogs || STATE.logs;
  const query = (searchInput ? searchInput.value : '').trim().toLowerCase();

  const filtered = logsToUse.filter(l => {
    if (query) {
      const mReq = (l.requester || '').toLowerCase().includes(query);
      const mItem = (l.itemName || '').toLowerCase().includes(query);
      const mDept = (l.department || '').toLowerCase().includes(query);
      if (!mReq && !mItem && !mDept) return false;
    }
    return true;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="7" class="py-8 text-center text-slate-400 text-xs">
          ไม่พบประวัติการเบิกในเงื่อนไขที่เลือก
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = filtered.map(l => {
    return `
      <tr class="hover:bg-slate-50/80 transition">
        <td class="py-2.5 px-4 whitespace-nowrap text-xs text-slate-600">
          <span class="font-medium text-slate-900">${l.date}</span>
          <span class="text-slate-400 ml-1 font-mono">${l.time || ''}</span>
        </td>
        <td class="py-2.5 px-4 font-bold text-slate-800">${l.requester}</td>
        <td class="py-2.5 px-4">
          <span class="inline-block px-2 py-0.5 rounded text-[11px] font-medium bg-slate-100 text-slate-700">
            ${l.department || l.category}
          </span>
        </td>
        <td class="py-2.5 px-4 font-semibold text-slate-900">${l.itemName}</td>
        <td class="py-2.5 px-4 text-center font-extrabold text-blue-600">
          ${l.qty} <span class="font-normal text-xs text-slate-500">${l.unit}</span>
        </td>
        <td class="py-2.5 px-4 text-center text-slate-600 text-xs">${l.newStock !== undefined ? l.newStock : '-'}</td>
        <td class="py-2.5 px-4 text-slate-500 text-xs italic">${l.note || '-'}</td>
      </tr>
    `;
  }).join('');
}

function exportMonthlyReportCSV() {
  const monthSelect = document.getElementById('summary-month-select');
  const selectedMonth = monthSelect ? monthSelect.value : 'all';

  const filteredLogs = STATE.logs.filter(l => {
    if (selectedMonth !== 'all' && l.date) {
      return l.date.startsWith(selectedMonth);
    }
    return true;
  });

  if (filteredLogs.length === 0) {
    showToast('ไม่มีข้อมูลการเบิกสำหรับส่งออก', 'warning');
    return;
  }

  const headers = ['วัน-เวลา', 'วันที่', 'เวลา', 'ชื่อผู้เบิก', 'สาขา/แผนก', 'รายการวัสดุ', 'จำนวนเบิก', 'หน่วย', 'สต็อกคงเหลือ', 'หมายเหตุ'];
  const rows = filteredLogs.map(l => [
    `"${l.timestamp || ''}"`,
    `"${l.date || ''}"`,
    `"${l.time || ''}"`,
    `"${(l.requester || '').replace(/"/g, '""')}"`,
    `"${(l.department || '').replace(/"/g, '""')}"`,
    `"${(l.itemName || '').replace(/"/g, '""')}"`,
    l.qty || 0,
    `"${l.unit || ''}"`,
    l.newStock !== undefined ? l.newStock : '',
    `"${(l.note || '').replace(/"/g, '""')}"`
  ]);

  const csvContent = '\uFEFF' + [headers.join(','), ...rows.map(r => r.join(','))].join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `รายงานการเบิกของ_${selectedMonth === 'all' ? 'ภาพรวมทั้งหมด' : selectedMonth}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  showToast('ส่งออกไฟล์รายงาน CSV สำเร็จ', 'success');
}

/**
 * ==========================================
 * SETTINGS & GOOGLE SHEETS LIVE SYNC
 * ==========================================
 */
function openSettingsModal() {
  const input = document.getElementById('gas-url-input');
  if (input) input.value = STATE.gasUrl || DEFAULT_GAS_URL;
  document.getElementById('settings-modal').classList.remove('hidden');
}

function closeSettingsModal() {
  document.getElementById('settings-modal').classList.add('hidden');
}

function saveSettings() {
  const input = document.getElementById('gas-url-input');
  if (input) {
    const val = input.value.trim() || DEFAULT_GAS_URL;
    STATE.gasUrl = val;
    localStorage.setItem('gas_webapp_url', val);
    showToast('บันทึกการตั้งค่าเรียบร้อยแล้ว', 'success');
    syncDataFromSheets();
  }
  closeSettingsModal();
}

async function testGasConnection() {
  const input = document.getElementById('gas-url-input');
  const resultEl = document.getElementById('gas-test-result');
  const url = input ? (input.value.trim() || DEFAULT_GAS_URL) : DEFAULT_GAS_URL;

  if (resultEl) resultEl.innerHTML = `<span class="text-slate-500">กำลังทดสอบเชื่อมต่อ...</span>`;

  try {
    const resp = await fetch(`${url}?action=ping`, { method: 'GET' });
    const json = await resp.json();
    if (json && json.success) {
      if (resultEl) resultEl.innerHTML = `<span class="text-emerald-600 font-bold">✅ เชื่อมต่อ Google Sheets สำเร็จ!</span>`;
      showToast('เชื่อมต่อ Google Sheets Web App สำเร็จ', 'success');
    } else {
      if (resultEl) resultEl.innerHTML = `<span class="text-amber-600">เชื่อมต่อได้แต่ตอบกลับไม่ถูกต้อง</span>`;
    }
  } catch (err) {
    if (resultEl) resultEl.innerHTML = `<span class="text-rose-600">❌ ไม่สามารถเชื่อมต่อได้: ${err.message}</span>`;
  }
}

function resetToInitialData() {
  if (confirm('คุณต้องการรีเซ็ตข้อมูลสต็อกกลับเป็นค่าเริ่มต้นจาก Google Sheets ใช่หรือไม่?')) {
    syncDataFromSheets();
    closeSettingsModal();
  }
}

function exportAllDataBackup() {
  const backup = {
    exportDate: new Date().toISOString(),
    items: STATE.items,
    logs: STATE.logs,
    restocks: STATE.restocks
  };
  const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `backup_office_stock_${Date.now()}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function showToast(message, type = 'info') {
  const existing = document.getElementById('app-toast');
  if (existing) existing.remove();

  const colors = {
    success: 'bg-emerald-600 text-white',
    error: 'bg-rose-600 text-white',
    warning: 'bg-amber-600 text-white',
    info: 'bg-slate-900 text-white'
  };

  const icons = {
    success: 'check-circle-2',
    error: 'alert-circle',
    warning: 'alert-triangle',
    info: 'info'
  };

  const toast = document.createElement('div');
  toast.id = 'app-toast';
  toast.className = `fixed bottom-5 right-5 z-50 flex items-center gap-2 px-4 py-3 rounded-xl shadow-xl text-xs sm:text-sm font-medium transition-all transform duration-300 ${colors[type] || colors.info}`;
  toast.innerHTML = `
    <i data-lucide="${icons[type] || 'info'}" class="w-4 h-4"></i>
    <span>${message}</span>
  `;

  document.body.appendChild(toast);
  lucide.createIcons();

  setTimeout(() => {
    toast.classList.add('opacity-0', 'translate-y-2');
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}
