// Global Config
const ACCESS_HASH = '42e709fc06944cc7d7c1944275a873bce3035fa3b0f7547736589a6c41b1c6ee';

async function handleSecurityCheck(e) {
  e.preventDefault();
  const input = document.getElementById('security-input');
  const error = document.getElementById('security-error');
  const modal = document.getElementById('security-modal');
  
  let inputVal = input.value.trim();

  // Helper: Extract ID if user pastes full URL
  if (inputVal.includes('/s/') && inputVal.includes('/exec')) {
      const match = inputVal.match(/\/s\/([a-zA-Z0-9_-]+)\/exec/);
      if (match && match[1]) {
          inputVal = match[1];
          // Update input field to show the clean ID
          input.value = inputVal;
      }
  }

  // Allow any valid-looking Script ID (starts with AKfy, reasonable length)
  // This allows you to redeploy and use new IDs without code changes.
  if (inputVal.startsWith('AKfy') && inputVal.length > 20) {
    // Success
    modal.classList.add('hidden');
    // Save session to avoid asking again on reload
    sessionStorage.setItem('auth_token', inputVal);
    
    // Load initial section
    loadSection('dashboard');
    // Initial Sync
    refreshAllData();
  } else {
    // Error
    error.innerHTML = '<i class="fa-solid fa-circle-exclamation mr-1.5"></i> Invalid Script ID format. It should start with "AKfy".';
    error.classList.remove('hidden');
    input.classList.add('border-red-500', 'focus:ring-red-500', 'focus:border-red-500');
    input.classList.remove('border-gray-300', 'focus:ring-blue-500', 'focus:border-blue-500');
    input.value = '';
    input.focus();
  }
}

function toggleSecurityPassword() {
  const input = document.getElementById('security-input');
  const icon = document.getElementById('security-eye');
  if (input.type === 'password') {
    input.type = 'text';
    icon.classList.remove('fa-eye');
    icon.classList.add('fa-eye-slash');
  } else {
    input.type = 'password';
    icon.classList.remove('fa-eye-slash');
    icon.classList.add('fa-eye');
  }
}

// Check session on init
document.addEventListener('DOMContentLoaded', async () => {
  const token = sessionStorage.getItem('auth_token');
  if (token && token.startsWith('AKfy')) {
      const modal = document.getElementById('security-modal');
      if (modal) modal.classList.add('hidden');
      loadSection('dashboard');
      refreshAllData();
  }
});

function getApiUrl() {
  const token = sessionStorage.getItem('auth_token');
  if (!token) return '';
  return `https://script.google.com/macros/s/${token}/exec`;
}

let globalLoadingCount = 0;
let globalLoadingTimer = null;

const globalSearchState = {
  section: 'dashboard',
  query: ''
};

let globalSearchTimer = null;
let dashboardStatsCache = null;
let reportsCache = { loaded: false, items: [], stats: null };
let stockTrackingTimer = null;
let stockTrackingInFlight = false;
let lastInventorySignature = '';

function setStockTrackingActive(active) {
  if (!active) {
    if (stockTrackingTimer) {
      clearInterval(stockTrackingTimer);
      stockTrackingTimer = null;
    }
    stockTrackingInFlight = false;
    return;
  }

  if (stockTrackingTimer) return;
  stockTrackingTimer = setInterval(stockTrackingTick, 6000);
  stockTrackingTick();
}

function computeInventorySignature(items) {
  const base = Array.isArray(items) ? items : [];
  return base
    .map(it => `${String(it?.ID ?? '')}|${String(it?.Serial ?? '')}|${String(it?.Qty ?? '')}`)
    .sort()
    .join('~');
}

async function stockTrackingTick() {
  if (document.hidden) return;
  if (globalSearchState.section !== 'inventory') return;
  if (stockTrackingInFlight) return;

  const invModal = document.getElementById('inventoryItemModal');
  if (invModal && !invModal.classList.contains('hidden')) return;

  const deleteModal = document.getElementById('inventoryDeleteModal');
  if (deleteModal && !deleteModal.classList.contains('hidden')) return;

  const adjustModal = document.getElementById('adjustStockModal');
  if (adjustModal && !adjustModal.classList.contains('hidden')) return;

  stockTrackingInFlight = true;
  try {
    const items = await callApi('getInventory', null, { silent: true });
    if (items?.error) return;

    const normalized = Array.isArray(items) ? items.map(normalizeInventoryItem) : [];
    const sig = computeInventorySignature(normalized);
    if (sig === lastInventorySignature) return;

    lastInventorySignature = sig;
    inventoryPagination.allItems = normalized;
    populateInventoryCategoryFilterOptions();
    applyInventorySearch(String(globalSearchState.query || '').trim(), { resetPage: false });
  } catch (_) {
  } finally {
    stockTrackingInFlight = false;
  }
}

const inventoryFilterState = {
  stock: 'all',
  status: 'all',
  category: 'all'
};

const dashboardFilterState = {
  chart: 'all',
  activityType: 'all'
};

const reportsLowStockPagination = {
  currentPage: 1,
  itemsPerPage: 10,
  lastQuery: ''
};

const reportsPrintOptions = {
  summary: true,
  lowStock: true,
  activity: true
};

let reportsLowStockPrintRestore = null;

window.addEventListener('beforeprint', () => {
  const reportsRoot = document.getElementById('reports-content');
  if (!reportsRoot) return;
  if (!reportsCache?.loaded) return;

  document.body.classList.add('printing-reports');

  if (reportsPrintOptions.lowStock) {
    reportsLowStockPrintRestore = {
      currentPage: reportsLowStockPagination.currentPage,
      itemsPerPage: reportsLowStockPagination.itemsPerPage
    };

    reportsLowStockPagination.currentPage = 1;
    reportsLowStockPagination.itemsPerPage = 1000000;
    renderReportsView(String(globalSearchState.query || '').trim());
  }
});

window.addEventListener('afterprint', () => {
  const reportsRoot = document.getElementById('reports-content');
  if (!reportsRoot) return;
  if (!reportsCache?.loaded) return;

  document.body.classList.remove('printing-reports');

  if (reportsLowStockPrintRestore) {
    reportsLowStockPagination.currentPage = reportsLowStockPrintRestore.currentPage;
    reportsLowStockPagination.itemsPerPage = reportsLowStockPrintRestore.itemsPerPage;
    reportsLowStockPrintRestore = null;
  }
  renderReportsView(String(globalSearchState.query || '').trim());
});

function showGlobalLoading(title, message) {
  const el = document.getElementById('global-loading');
  if (!el) return;
  
  if (title) {
    const t = document.getElementById('loading-title');
    if (t) t.innerText = title;
  }
  if (message) {
    const m = document.getElementById('loading-message');
    if (m) m.innerText = message;
  }

  globalLoadingCount += 1;

  if (globalLoadingCount === 1) {
    if (globalLoadingTimer) clearTimeout(globalLoadingTimer);
    globalLoadingTimer = setTimeout(() => {
      if (globalLoadingCount > 0) el.classList.remove('hidden');
    }, 150);
  }
}

function hideGlobalLoading() {
  const el = document.getElementById('global-loading');
  if (!el) return;
  globalLoadingCount = Math.max(0, globalLoadingCount - 1);

  if (globalLoadingCount === 0) {
    if (globalLoadingTimer) {
      clearTimeout(globalLoadingTimer);
      globalLoadingTimer = null;
    }
    el.classList.add('hidden');
  }
}

// --- API Helper ---
async function callApi(action, data = null, options) {
  const opts = options && typeof options === 'object' ? options : {};
  const silent = opts.silent === true;
  const title = opts.title || 'Syncing Data';
  const message = opts.message || 'Please wait...';

  if (!silent) showGlobalLoading(title, message);
  try {
    let response;
    if (data) {
      // POST request
      // Use text/plain to avoid CORS preflight, handled manually in GAS doPost
      response = await fetch(`${getApiUrl()}?action=${action}`, {
        method: 'POST',
        body: JSON.stringify(data),
        headers: {
            'Content-Type': 'text/plain;charset=utf-8' 
        }
      });
    } else {
      // GET request
      response = await fetch(`${getApiUrl()}?action=${action}`);
    }
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const contentType = response.headers.get('content-type');
    if (!contentType || !contentType.includes('application/json')) {
       // Only throw if we expected JSON. GAS sometimes returns text/plain for success if we used ContentService.createTextOutput without setMimeType, 
       // but we used setMimeType(JSON) in Code.js.
       // However, error pages (Google Sign In) are text/html.
       const text = await response.text();
       if (text.includes('<!DOCTYPE html>') || text.includes('Google Accounts')) {
           throw new Error('Authentication failed or Script Deployment URL changed. Please re-login.');
       }
       // Try parsing anyway just in case
       try {
          return JSON.parse(text);
       } catch (e) {
          throw new Error('Invalid JSON response from server');
       }
    }
    
    const json = await response.json();
    if (json.error && !json.message && typeof json.error === 'string') {
      json.message = json.error;
    }
    return json;
  } catch (error) {
    if (!silent) {
      console.error('API Error:', error);
      if (error.message.includes('Authentication failed') || error.message.includes('Script Deployment URL changed')) {
          alert('Session expired or Deployment ID changed. Please enter the new Script ID.');
          sessionStorage.removeItem('auth_token');
          window.location.reload();
          return { error: true };
      }
      alert('API Error: ' + error.message);
      throw error;
    }
    return { error: true, message: error.message };
  } finally {
    if (!silent) hideGlobalLoading();
  }
}

// --- Navigation & Loading ---

async function refreshAllData() {
  showGlobalLoading('Syncing Data', 'Updating system information...');
  try {
    const [inventory, stats] = await Promise.all([
      callApi('getInventory', null, { silent: true }),
      callApi('getDashboardStats', null, { silent: true })
    ]);

    // Handle Inventory
    if (!inventory.error) {
       const normalized = Array.isArray(inventory) ? inventory.map(normalizeInventoryItem) : [];
       inventoryPagination.allItems = normalized;
    }

    // Handle Dashboard
    if (!stats.error) {
       dashboardStatsCache = stats;
    }

    // Handle Reports Cache
    if (!inventory.error && !stats.error) {
        reportsCache = {
            loaded: true,
            items: inventoryPagination.allItems,
            stats: dashboardStatsCache
        };
    }

    // Re-render current section
    const currentSection = globalSearchState.section;
    if (currentSection === 'dashboard' && dashboardStatsCache) {
        renderDashboard(dashboardStatsCache);
    } else if (currentSection === 'inventory' && inventoryPagination.allItems) {
        populateInventoryCategoryFilterOptions();
        applyInventorySearch(String(globalSearchState.query || '').trim(), { resetPage: false });
    } else if (currentSection === 'reports' && reportsCache.loaded) {
        renderReportsView(String(globalSearchState.query || '').trim());
    }

  } catch (e) {
    console.error(e);
    alert('Failed to sync data. Please check your connection.');
  } finally {
    hideGlobalLoading();
  }
}

async function loadSection(sectionId) {
    // Hide all
    document.querySelectorAll('[id^="section-"]').forEach(el => el.classList.add('hidden'));
    
    // Show wrapper
    const wrapper = document.getElementById('section-' + sectionId);
    wrapper.classList.remove('hidden');
    
    // Update Nav
    document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
    document.getElementById('nav-' + sectionId).classList.add('active');
    
    // Update Title
    const titles = {
      'dashboard': 'Dashboard',
      'inventory': 'Inventory Management',
      'reports': 'Reports'
    };
    document.getElementById('page-title').innerText = titles[sectionId];

    globalSearchState.section = sectionId;
    updateGlobalSearchPlaceholder();
    
    // Load Content dynamically if not already loaded (simple cache)
    // Note: Since we are fetching local HTML files, we assume they are in the same directory.
    // If we want to reload fresh data every time, we call the data loader.
    // But the HTML structure only needs to be loaded once.
    if (!wrapper.hasAttribute('data-loaded')) {
        try {
            const moduleName = titles[sectionId] || 'Module';
            showGlobalLoading(`Loading ${moduleName}`, 'Initializing interface...');
            // Add timestamp to prevent caching of HTML templates
            const resp = await fetch(`${sectionId}.html?v=${Date.now()}`);
            if(!resp.ok) throw new Error('Failed to load template');
            const html = await resp.text();
            wrapper.innerHTML = html;
            wrapper.setAttribute('data-loaded', 'true');
        } catch(e) {
            wrapper.innerHTML = `<p class="text-red-500">Error loading module: ${e.message}</p>`;
            return;
        } finally {
            hideGlobalLoading();
        }
    }

    initSectionFilters(sectionId);
    
    // Close mobile menu on navigation
    if (window.innerWidth < 1024) {
      const sidebar = document.getElementById('sidebar');
      if (sidebar) {
        sidebar.classList.add('-translate-x-full');
      }
    }

    applyGlobalSearch();
    
    // Trigger Render from Cache (No Auto-Fetch)
    if(sectionId === 'dashboard') {
        if(dashboardStatsCache) renderDashboard(dashboardStatsCache);
    }
    if(sectionId === 'inventory') {
       if(Array.isArray(inventoryPagination.allItems)) {
          populateInventoryCategoryFilterOptions();
          applyInventorySearch(String(globalSearchState.query || '').trim(), { resetPage: false });
      }
    }
    if(sectionId === 'reports') {
        if(reportsCache.loaded) {
            renderReportsView(String(globalSearchState.query || '').trim());
        }
    }
}

function updateGlobalSearchPlaceholder() {
  const input = document.getElementById('global-search');
  if (!input) return;
  const placeholders = {
    dashboard: 'Search recent activities...',
    inventory: 'Search inventory...',
    reports: 'Search reports...'
  };
  input.placeholder = placeholders[globalSearchState.section] || 'Search...';
}

function applyGlobalSearch() {
  const q = String(globalSearchState.query || '').trim();
  if (globalSearchState.section === 'inventory') applyInventorySearch(q);
  if (globalSearchState.section === 'reports') applyReportsSearch(q);
  if (globalSearchState.section === 'dashboard') applyDashboardSearch(q);
}

function initSectionFilters(sectionId) {
  if (sectionId === 'inventory') initInventoryFilters();
  if (sectionId === 'dashboard') initDashboardFilters();
}

function initGlobalSearch() {
  const input = document.getElementById('global-search');
  if (!input) return;

  const schedule = () => {
    if (globalSearchTimer) clearTimeout(globalSearchTimer);
    globalSearchTimer = setTimeout(() => {
      globalSearchState.query = input.value || '';
      applyGlobalSearch();
    }, 120);
  };

  input.addEventListener('input', schedule);
  input.addEventListener('search', schedule);
  updateGlobalSearchPlaceholder();
}

// --- Dashboard Logic ---
async function loadDashboard() {
  const stats = await callApi('getDashboardStats');
  if(stats.error) return; // Handled in callApi or show alert
  renderDashboard(stats);
}

function animateNumber(el, toValue, formatter, durationMs = 650) {
  if (!el) return;
  const fromRaw = el.getAttribute('data-anim-value');
  const fromValue = fromRaw ? Number(fromRaw) : 0;
  const from = Number.isFinite(fromValue) ? fromValue : 0;
  const to = Number.isFinite(toValue) ? toValue : 0;

  el.setAttribute('data-anim-value', String(to));

  const start = performance.now();
  const easeOut = t => 1 - Math.pow(1 - t, 3);

  const tick = now => {
    const t = Math.min(1, (now - start) / durationMs);
    const eased = easeOut(t);
    const current = from + (to - from) * eased;
    el.innerText = formatter(current);
    if (t < 1) requestAnimationFrame(tick);
  };

  requestAnimationFrame(tick);
}

let isLowStockDismissed = false;

function showLowStockToast(lowStockCount) {
  const toastEl = document.getElementById('low-stock-toast');
  if (!toastEl) return;

  // Don't show if dismissed in this session
  if (isLowStockDismissed) {
    toastEl.classList.add('hidden');
    return;
  }

  const count = Number(lowStockCount) || 0;
  if (count <= 0) {
    toastEl.classList.add('opacity-0', 'translate-y-2');
    toastEl.classList.remove('opacity-100', 'translate-y-0');
    const hideId = window.setTimeout(() => toastEl.classList.add('hidden'), 250);
    return;
  }

  // Show
  toastEl.classList.remove('hidden');
  // Small delay to allow transition
  window.requestAnimationFrame(() => {
    toastEl.classList.remove('opacity-0', 'translate-y-2');
    toastEl.classList.add('opacity-100', 'translate-y-0');
  });
}

function dismissLowStockToast() {
  const toastEl = document.getElementById('low-stock-toast');
  if (!toastEl) return;
  
  isLowStockDismissed = true;
  toastEl.classList.add('opacity-0', 'translate-y-2');
  toastEl.classList.remove('opacity-100', 'translate-y-0');
  setTimeout(() => toastEl.classList.add('hidden'), 300);
}

function renderDashboard(stats) {
  const fmtInt = n => Math.round(n).toLocaleString();
  const fmtCurrency = n => toPeso(n);

  dashboardStatsCache = stats;
  populateDashboardActivityFilterOptions();

  const totalItemsEl = document.getElementById('dash-total-items');
  const lowStockEl = document.getElementById('dash-low-stock');
  const outStockEl = document.getElementById('dash-out-stock');
  const totalValueEl = document.getElementById('dash-total-value');

  animateNumber(totalItemsEl, Number(stats.totalItems) || 0, fmtInt);
  animateNumber(lowStockEl, Number(stats.lowStock) || 0, fmtInt);
  animateNumber(outStockEl, Number(stats.outOfStock) || 0, fmtInt);
  animateNumber(totalValueEl, Number(stats.totalValue) || 0, fmtCurrency, 800);

  const notifEl = document.getElementById('dash-notification-count');
  if (notifEl) notifEl.innerText = stats.lowStock;
  showLowStockToast(stats.lowStock);

  const syncEl = document.getElementById('dash-last-sync');
  if (syncEl) syncEl.innerText = new Date().toLocaleString();
  
  applyDashboardSearch(String(globalSearchState.query || '').trim());
  
  // Render Chart
  renderChart(stats);
}

function applyDashboardSearch(query) {
  const activityList = document.getElementById('recent-activities-list');
  if (!activityList) return;

  const all = Array.isArray(dashboardStatsCache?.recentActivities) ? dashboardStatsCache.recentActivities : [];
  const q = String(query || '').trim().toLowerCase();
  const terms = q.split(/\s+/).filter(Boolean);
  const type = String(dashboardFilterState.activityType || 'all').trim().toLowerCase();

  const filtered = all.filter(act => {
    const actType = String(act?.Type ?? '').trim().toLowerCase();
    if (type !== 'all' && actType !== type) return false;
    if (terms.length === 0) return true;

        const hay = [
          act?.Type,
          act?.ItemName,
          act?.Notes,
          act?.User,
          act?.Quantity,
          act?.Date,
          act?.Timestamp
        ]
          .map(v => String(v ?? ''))
          .join(' ')
          .toLowerCase();
        return terms.every(term => hay.includes(term));
      });

  activityList.innerHTML = '';
  if (filtered.length === 0) {
    activityList.innerHTML = '<li class="py-3 text-sm text-gray-500">No recent activity.</li>';
    return;
  }

  filtered.forEach(act => {
    const li = document.createElement('li');
    li.className = 'py-3';
    li.innerHTML = `
      <div class="flex space-x-3">
        <div class="flex-1 space-y-1">
          <div class="flex items-center justify-between">
            <h3 class="text-sm font-medium text-gray-900">${escapeHtml(act.Type)} - ${escapeHtml(act.ItemName)}</h3>
            <p class="text-sm text-gray-500">${new Date(act.Date).toLocaleDateString()}</p>
          </div>
          <p class="text-sm text-gray-500">${escapeHtml(act.Notes)} (Qty: ${escapeHtml(act.Quantity)}) by ${escapeHtml(act.User)}</p>
        </div>
      </div>
    `;
    activityList.appendChild(li);
  });
}

function populateDashboardActivityFilterOptions() {
  const select = document.getElementById('dash-activity-filter');
  if (!select) return;

  const all = Array.isArray(dashboardStatsCache?.recentActivities) ? dashboardStatsCache.recentActivities : [];
  const current = select.value || 'all';

  const types = Array.from(
    new Set(
      all
        .map(a => String(a?.Type ?? '').trim())
        .filter(Boolean)
    )
  ).sort((a, b) => a.localeCompare(b));

  select.innerHTML = '<option value="all">All</option>' + types.map(t => `<option value="${t}">${t}</option>`).join('');
  select.value = types.includes(current) ? current : 'all';
  dashboardFilterState.activityType = select.value === 'all' ? 'all' : String(select.value).trim();
}

function initDashboardFilters() {
  const chartSelect = document.getElementById('dash-chart-filter');
  if (chartSelect && !chartSelect.dataset.bound) {
    chartSelect.dataset.bound = '1';
    chartSelect.addEventListener('change', () => {
      dashboardFilterState.chart = chartSelect.value || 'all';
      applyDashboardChartFilter();
    });
  }

  const activitySelect = document.getElementById('dash-activity-filter');
  if (activitySelect && !activitySelect.dataset.bound) {
    activitySelect.dataset.bound = '1';
    activitySelect.addEventListener('change', () => {
      dashboardFilterState.activityType = activitySelect.value || 'all';
      applyDashboardSearch(String(globalSearchState.query || '').trim());
    });
  }
}

function applyDashboardChartFilter() {
  const chart = window.myChart instanceof Chart ? window.myChart : null;
  if (!chart) return;

  const mode = String(dashboardFilterState.chart || 'all');
  const show = idx => chart.setDatasetVisibility(idx, true);
  const hide = idx => chart.setDatasetVisibility(idx, false);

  if (mode === 'sales') {
    show(0); hide(1); hide(2);
  } else if (mode === 'restocks') {
    hide(0); show(1); hide(2);
  } else if (mode === 'trend') {
    hide(0); hide(1); show(2);
  } else {
    show(0); show(1); show(2);
  }

  chart.update('none');
}

function renderChart(stats) {
    const mainCanvas = document.getElementById('mainChart');
    if(!mainCanvas) return;

    const mainCtx = mainCanvas.getContext('2d');
    const salesGradient = mainCtx.createLinearGradient(0, 0, 0, mainCanvas.height || 280);
    salesGradient.addColorStop(0, 'rgba(59, 130, 246, 0.55)');
    salesGradient.addColorStop(1, 'rgba(59, 130, 246, 0.10)');

    const restockGradient = mainCtx.createLinearGradient(0, 0, 0, mainCanvas.height || 280);
    restockGradient.addColorStop(0, 'rgba(16, 185, 129, 0.50)');
    restockGradient.addColorStop(1, 'rgba(16, 185, 129, 0.10)');

    if(window.myChart instanceof Chart) window.myChart.destroy();
    if(window.stockChart instanceof Chart) window.stockChart.destroy();

    Chart.defaults.font.family = 'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial';
    Chart.defaults.color = '#475569';

    const wa = stats && stats.weeklyActivity ? stats.weeklyActivity : null;
    const asNums = arr => Array.isArray(arr) ? arr.map(v => Number(v)) : [];
    const isValidNums = arr => Array.isArray(arr) && arr.length === 7 && arr.every(n => Number.isFinite(n));
    const labels = Array.isArray(wa?.labels) && wa.labels.length === 7 ? wa.labels.map(s => String(s)) : ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
    const sales = (() => {
      const a = asNums(wa?.sales);
      return isValidNums(a) ? a : [12, 19, 3, 5, 2, 3, 10];
    })();
    const restocks = (() => {
      const a = asNums(wa?.restocks);
      return isValidNums(a) ? a : [2, 3, 20, 5, 1, 4, 2];
    })();
    const trend = (() => {
      const a = asNums(wa?.trend);
      return isValidNums(a) ? a : labels.map((_, i) => (sales[i] + restocks[i]) / 2);
    })();

    window.myChart = new Chart(mainCtx, {
        data: {
            labels,
            datasets: [
              {
                type: 'bar',
                label: 'Stock Out',
                data: sales,
                backgroundColor: salesGradient,
                borderColor: 'rgba(59, 130, 246, 0.95)',
                borderWidth: 1,
                borderRadius: 14,
                borderSkipped: false,
                barThickness: 18,
                maxBarThickness: 20
              },
              {
                type: 'bar',
                label: 'Stock In',
                data: restocks,
                backgroundColor: restockGradient,
                borderColor: 'rgba(16, 185, 129, 0.95)',
                borderWidth: 1,
                borderRadius: 14,
                borderSkipped: false,
                barThickness: 18,
                maxBarThickness: 20
              },
              {
                type: 'line',
                label: 'Average',
                data: trend,
                borderColor: 'rgba(99, 102, 241, 0.95)',
                backgroundColor: 'rgba(99, 102, 241, 0.12)',
                pointRadius: 0,
                borderWidth: 2,
                tension: 0.42,
                fill: true
              }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            animation: { duration: 900, easing: 'easeOutQuart' },
            plugins: {
                legend: { position: 'bottom', align: 'end', labels: { boxWidth: 12, usePointStyle: true, padding: 20 } },
                tooltip: {
                    backgroundColor: 'rgba(15, 23, 42, 0.9)',
                    padding: 12,
                    cornerRadius: 12,
                    titleFont: { size: 13, weight: 600 },
                    bodyFont: { size: 12 },
                    displayColors: true,
                    boxPadding: 4,
                    callbacks: {
                      label: ctx => {
                        const raw = (ctx.parsed && typeof ctx.parsed.y !== 'undefined') ? ctx.parsed.y : ctx.parsed;
                        const n = Number(raw);
                        const v = Number.isFinite(n) ? (Number.isInteger(n) ? n : n.toFixed(1)) : raw;
                        const label = String(ctx.dataset && ctx.dataset.label ? ctx.dataset.label : '').trim();
                        return label ? `${label}: ${v}` : String(v);
                      }
                    }
                }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: '#f1f5f9', drawBorder: false }, ticks: { padding: 10 } },
                x: { grid: { display: false, drawBorder: false }, ticks: { padding: 10 } }
            }
        }
    });

    const stockCanvas = document.getElementById('stockChart');
    if (!stockCanvas) return;

    const stockCtx = stockCanvas.getContext('2d');
    const totalItems = Number(stats?.totalItems) || 0;
    const lowStock = Number(stats?.lowStock) || 0;
    const outOfStock = Number(stats?.outOfStock) || 0;
    const inStock = Math.max(0, totalItems - lowStock - outOfStock);

    window.stockChart = new Chart(stockCtx, {
        type: 'doughnut',
        data: {
            labels: ['In Stock', 'Low Stock', 'Out of Stock'],
            datasets: [
              {
                data: [inStock, lowStock, outOfStock],
                backgroundColor: ['rgba(59, 130, 246, 0.75)', 'rgba(245, 158, 11, 0.75)', 'rgba(244, 63, 94, 0.75)'],
                borderColor: ['rgba(59, 130, 246, 1)', 'rgba(245, 158, 11, 1)', 'rgba(244, 63, 94, 1)'],
                borderWidth: 1,
                hoverOffset: 6
              }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '68%',
            plugins: {
              legend: { 
                position: 'bottom', 
                labels: { 
                  boxWidth: 12, 
                  usePointStyle: true, 
                  padding: 18,
                  generateLabels: (chart) => {
                    const data = chart.data;
                    if (data.labels.length && data.datasets.length) {
                      return data.labels.map((label, i) => {
                        const val = data.datasets[0].data[i];
                        const meta = chart.getDatasetMeta(0);
                        const style = meta.controller.getStyle(i);
                        return {
                          text: `${label}: ${val}`,
                          fillStyle: style.backgroundColor,
                          strokeStyle: style.borderColor,
                          lineWidth: style.borderWidth,
                          hidden: !chart.getDataVisibility(i),
                          index: i
                        };
                      });
                    }
                    return [];
                  }
                } 
              },
              tooltip: {
                  backgroundColor: 'rgba(15, 23, 42, 0.9)',
                  padding: 12,
                  cornerRadius: 12,
                  titleFont: { size: 13, weight: 600 },
                  bodyFont: { size: 12 },
                  displayColors: true,
                  boxPadding: 4,
                  callbacks: {
                    label: ctx => {
                      const raw = ctx.parsed;
                      const n = Number(raw);
                      const v = Number.isFinite(n) ? (Number.isInteger(n) ? n : n.toFixed(1)) : raw;
                      const label = String(ctx.label || '').trim();
                      return label ? `${label}: ${v}` : String(v);
                    }
                  }
              }
            }
        }
    });
}

// --- Inventory Logic ---
async function loadInventory() {
  setStockTrackingActive(true);
  const items = await callApi('getInventory');
  
  if(items.error) return; // Handled in callApi

  const normalized = Array.isArray(items) ? items.map(normalizeInventoryItem) : [];
  inventoryPagination.allItems = normalized;
  
  populateInventoryCategoryFilterOptions();
  applyInventorySearch(String(globalSearchState.query || '').trim(), { resetPage: true });
}

function normalizeInventoryItem(item) {
  // Ensure consistent keys
  return {
    ID: item.ID,
    Project: item.Project,
    Category: item.Category,
    Item: item.Item,
    BrandModel: item.BrandModel,
    Serial: item.Serial,
    Qty: item.Qty,
    Unit: item.Unit,
    UnitCost: item.UnitCost,
    DateAcquired: item.DateAcquired,
    ProcurementProject: item.ProcurementProject,
    PersonInCharge: item.PersonInCharge,
    Location: item.Location,
    Status: item.Status,
    Remarks: item.Remarks,
    ...item
  };
}

const inventoryPagination = {
  currentPage: 1,
  itemsPerPage: 10,
  allItems: [],
  currentFiltered: []
};

function applyInventorySearch(query, options) {
  const resetPage = options?.resetPage ?? true;
  const q = String(query || '').trim().toLowerCase();
  
  // Filters
  const stockFilter = inventoryFilterState.stock;
  const statusFilter = String(inventoryFilterState.status || 'all').toLowerCase();
  const categoryFilter = inventoryFilterState.category;

  const terms = q.split(/\s+/).filter(Boolean);
  const matchesTerms = hay => terms.length === 0 || terms.every(t => hay.includes(t));

  const filtered = inventoryPagination.allItems.filter(item => {
    // 1. Text Search
    const hay = Object.values(item)
      .map(v => String(v ?? ''))
      .join(' ')
      .toLowerCase();
    
    if (!matchesTerms(hay)) return false;

    // 2. Stock Filter
    const qty = Number(item.Qty);
    if (stockFilter === 'low') {
      if (qty >= 10 || qty <= 0) return false;
    } else if (stockFilter === 'out') {
      if (qty > 0) return false;
    } else if (stockFilter === 'instock') {
      if (qty <= 0) return false;
    }

    // 3. Status Filter
    if (statusFilter !== 'all') {
      const s = String(item.Status || '').toLowerCase();
      if (s !== statusFilter) return false;
    }

    // 4. Category Filter
    if (categoryFilter && categoryFilter !== 'all') {
      const c = String(item.Category || '');
      if (c !== categoryFilter) return false;
    }

    return true;
  });

  inventoryPagination.currentFiltered = filtered;
  if (resetPage) inventoryPagination.currentPage = 1;
  
  renderInventoryTable();
}

function renderInventoryTable() {
  const { currentFiltered, currentPage, itemsPerPage } = inventoryPagination;
  const total = currentFiltered.length;
  const totalPages = Math.ceil(total / itemsPerPage) || 1;
  
  // Clamp page
  if (inventoryPagination.currentPage > totalPages) inventoryPagination.currentPage = totalPages;
  if (inventoryPagination.currentPage < 1) inventoryPagination.currentPage = 1;

  const startIdx = (inventoryPagination.currentPage - 1) * itemsPerPage;
  const endIdx = startIdx + itemsPerPage;
  const pageItems = currentFiltered.slice(startIdx, endIdx);

  renderInventory(pageItems);
  updateInventoryPaginationControls(total, totalPages, startIdx + 1, Math.min(endIdx, total));
}

function updateInventoryPaginationControls(total, totalPages, start, end) {
  const info = document.getElementById('inv-page-info');
  const prevBtn = document.getElementById('inv-prev-btn');
  const nextBtn = document.getElementById('inv-next-btn');

  if (info) info.innerText = total === 0 ? 'No items' : `Showing ${start} to ${end} of ${total} results`;
  
  if (prevBtn) {
    prevBtn.disabled = inventoryPagination.currentPage <= 1;
    prevBtn.onclick = () => {
      if (inventoryPagination.currentPage > 1) {
        inventoryPagination.currentPage--;
        renderInventoryTable();
      }
    };
  }

  if (nextBtn) {
    nextBtn.disabled = inventoryPagination.currentPage >= totalPages;
    nextBtn.onclick = () => {
      if (inventoryPagination.currentPage < totalPages) {
        inventoryPagination.currentPage++;
        renderInventoryTable();
      }
    };
  }
}

function toPeso(val) {
  const n = parseFloat(val);
  if (isNaN(n)) return '₱0.00';
  return '₱' + n.toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// --- Modal Helpers ---
function openModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) {
    modal.classList.remove('hidden');
    // Prevent body scroll
    document.body.style.overflow = 'hidden';
  }
}

function closeModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) {
    modal.classList.add('hidden');
    // Restore body scroll
    document.body.style.overflow = '';
  }
}

function openInventoryCreateModal() {
  document.getElementById('inventory-modal-title').innerText = 'Add New Item';
  
  // Clear inputs
  const fields = ['project', 'category', 'item', 'brandmodel', 'serial', 'qty', 'unit', 'unitcost', 'dateacquired', 'procurementproject', 'personincharge', 'location', 'status', 'remarks'];
  fields.forEach(f => {
      const el = document.getElementById('inv-' + f);
      if(el) el.value = '';
  });
  
  document.getElementById('inv-item-id').value = '';
  openModal('inventoryItemModal');
}

function openInventoryEditModal(id) {
  const item = inventoryPagination.allItems.find(i => String(i.ID) === String(id));
  if (!item) return;

  document.getElementById('inventory-modal-title').innerText = 'Edit Item';
  document.getElementById('inv-item-id').value = item.ID;
  
  // Fill form
  const fields = ['Project', 'Category', 'Item', 'BrandModel', 'Serial', 'Qty', 'Unit', 'UnitCost', 'DateAcquired', 'ProcurementProject', 'PersonInCharge', 'Location', 'Status', 'Remarks'];
  fields.forEach(f => {
      const el = document.getElementById('inv-' + f.toLowerCase());
      if(el) el.value = item[f] || '';
  });

  openModal('inventoryItemModal');
}

async function submitInventoryItem() {
  const data = {};
  const fields = ['Project', 'Category', 'Item', 'BrandModel', 'Serial', 'Qty', 'Unit', 'UnitCost', 'DateAcquired', 'ProcurementProject', 'PersonInCharge', 'Location', 'Status', 'Remarks'];
  
  fields.forEach(f => {
      const el = document.getElementById('inv-' + f.toLowerCase());
      if(el) data[f] = el.value;
  });

  const idEl = document.getElementById('inv-item-id');
  const id = idEl ? idEl.value : '';
  if (id) data.id = id;

  const action = id ? 'editItem' : 'addItem';
  const res = await callApi(action, data);
  
  if (res.success) {
    closeModal('inventoryItemModal');
    loadInventory();
  } else {
    alert('Error: ' + res.message);
  }
}

let itemToDelete = null;
function deleteInventoryItem(id) {
  itemToDelete = id;
  
  // Update UI
  const item = inventoryPagination.allItems.find(i => String(i.ID) === String(id));
  const nameEl = document.getElementById('inv-delete-name');
  if (nameEl) nameEl.innerText = item ? (item.Item || 'Unknown Item') : 'Item';
  
  const input = document.getElementById('inv-delete-confirm');
  if (input) input.value = '';
  
  const btn = document.getElementById('inventory-delete-confirm-btn');
  if (btn) btn.disabled = true;

  openModal('inventoryDeleteModal');
}

function inventoryDeleteValidationChanged() {
  const input = document.getElementById('inv-delete-confirm');
  const btn = document.getElementById('inventory-delete-confirm-btn');
  if (!input || !btn) return;
  
  if (input.value === 'DELETE') {
    btn.disabled = false;
    btn.classList.remove('opacity-50', 'cursor-not-allowed');
  } else {
    btn.disabled = true;
    btn.classList.add('opacity-50', 'cursor-not-allowed');
  }
}

async function confirmInventoryDelete() {
  if (!itemToDelete) return;
  
  const btn = document.getElementById('inventory-delete-confirm-btn');
  if(btn) {
      btn.disabled = true;
      btn.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin"></i> Deleting...';
  }

  const res = await callApi('deleteItem', { id: itemToDelete });
  
  if (res.success) {
    closeModal('inventoryDeleteModal');
    loadInventory();
  } else {
    alert('Error: ' + res.message);
  }
  
  if(btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="inline-flex items-center space-x-2"><i class="fa-solid fa-trash"></i><span>Delete</span></span>';
  }
}

function populateInventoryCategoryFilterOptions() {
  const select = document.getElementById('inv-filter-category');
  if (!select) return;
  const current = select.value || 'all';

  const categories = Array.from(
    new Set(
      (Array.isArray(inventoryPagination.allItems) ? inventoryPagination.allItems : [])
        .map(it => String(it?.Category ?? '').trim())
        .filter(Boolean)
    )
  ).sort((a, b) => a.localeCompare(b));

  select.innerHTML = '<option value="all">All Categories</option>' + categories.map(c => `<option value="${c}">${c}</option>`).join('');
  select.value = categories.includes(current) ? current : 'all';
  inventoryFilterState.category = select.value;
}

function initInventoryFilters() {
  const stockEl = document.getElementById('inv-filter-stock');
  if (stockEl && !stockEl.dataset.bound) {
    stockEl.dataset.bound = '1';
    stockEl.addEventListener('change', () => {
      inventoryFilterState.stock = stockEl.value || 'all';
      applyInventorySearch(String(globalSearchState.query || '').trim());
    });
  }

  const statusEl = document.getElementById('inv-filter-status');
  if (statusEl && !statusEl.dataset.bound) {
    statusEl.dataset.bound = '1';
    statusEl.addEventListener('change', () => {
      inventoryFilterState.status = String(statusEl.value || 'all').toLowerCase();
      applyInventorySearch(String(globalSearchState.query || '').trim());
    });
  }

  const categoryEl = document.getElementById('inv-filter-category');
  if (categoryEl && !categoryEl.dataset.bound) {
    categoryEl.dataset.bound = '1';
    categoryEl.addEventListener('change', () => {
      inventoryFilterState.category = categoryEl.value || 'all';
      applyInventorySearch(String(globalSearchState.query || '').trim());
    });
  }

  populateInventoryCategoryFilterOptions();
}

function getDisposalCountdown(dateStr) {
  if (!dateStr) return '<span class="text-gray-400 text-xs">N/A</span>';
  
  let acquiredDate;
  // Handle Year-only input (e.g. "2025")
  if (/^\d{4}$/.test(String(dateStr).trim())) {
    acquiredDate = new Date(`${dateStr}-01-01`);
  } else {
    acquiredDate = new Date(dateStr);
  }

  if (isNaN(acquiredDate.getTime())) return '<span class="text-gray-400 text-xs">Invalid Date</span>';
  
  // 5 Years Disposal Logic
  const disposalDate = new Date(acquiredDate);
  disposalDate.setFullYear(disposalDate.getFullYear() + 5);

  
  const now = new Date();
  // Set to midnight for accurate day calculation
  now.setHours(0, 0, 0, 0);
  disposalDate.setHours(0, 0, 0, 0);
  
  const diffTime = disposalDate - now;
  
  if (diffTime < 0) {
    return '<span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-red-100 text-red-800">Expired</span>';
  }
  
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  const years = Math.floor(diffDays / 365);
  const days = diffDays % 365;
  
  let text = '';
  if (years > 0) text += `${years}y `;
  text += `${days}d`;
  
  // Color coding
  // < 3 months (approx 90 days) = Warning
  // < 1 year = Notice
  // > 1 year = Good
  
  if (diffDays < 90) {
    return `<span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-amber-100 text-amber-800" title="Disposal soon">${text}</span>`;
  } else if (diffDays < 365) {
     return `<span class="text-amber-600 font-medium text-xs">${text}</span>`;
  }
  
  return `<span class="text-emerald-600 font-medium text-xs">${text}</span>`;
}

function escapeHtml(text) {
  if (text === null || text === undefined) return '';
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function renderInventory(items) {
  const tbody = document.getElementById('inventory-table-body');
  if(!tbody) return;
  tbody.innerHTML = '';
  
  if (items.length === 0) {
      tbody.innerHTML = `
        <tr>
            <td colspan="16" class="px-6 py-8 text-center text-sm text-gray-500">
                <div class="flex flex-col items-center justify-center">
                    <i class="fa-solid fa-box-open text-gray-300 text-4xl mb-3"></i>
                    <p>No inventory items found.</p>
                </div>
            </td>
        </tr>
      `;
      return;
  }

  items.forEach(item => {
     const tr = document.createElement('tr');
     tr.className = 'hover:bg-gray-50';
     const hasId = Boolean(item.ID);
     const qtyNum = Number(item.Qty);
     const canTake = hasId && Number.isFinite(qtyNum) && qtyNum > 0;
     const actionBaseClass = 'p-2 rounded-lg border border-gray-200 text-gray-700';
     const actionEnabledClass = ' hover:bg-gray-50';
     const actionDisabledClass = ' opacity-40 cursor-not-allowed';
     
     // XSS Prevention: Use escapeHtml for all user-provided content
     tr.innerHTML = `
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">${escapeHtml(item.Project)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${escapeHtml(item.Category)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${escapeHtml(item.Item)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${escapeHtml(item.BrandModel)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${escapeHtml(item.Serial)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900 font-bold">${escapeHtml(item.Qty)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${escapeHtml(item.Unit)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${escapeHtml(toPeso(item.UnitCost))}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${escapeHtml(item.DateAcquired)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${getDisposalCountdown(item.DateAcquired)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${escapeHtml(item.ProcurementProject)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${escapeHtml(item.PersonInCharge)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${escapeHtml(item.Location)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${escapeHtml(item.Status)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${escapeHtml(item.Remarks)}</td>
       <td class="px-6 py-4 whitespace-nowrap text-right text-sm">
         <div class="inline-flex items-center justify-end space-x-2">
           <button class="p-2 rounded-lg border border-gray-200 text-emerald-600${canTake ? ' hover:bg-emerald-50' : ' opacity-40 cursor-not-allowed'}" ${canTake ? `onclick='takeOneInventoryItem("${escapeHtml(item.ID)}")'` : 'disabled'} title="Take 1">
             <i class="fa-solid fa-minus"></i>
           </button>
           <button class="${actionBaseClass}${hasId ? actionEnabledClass : actionDisabledClass}" ${hasId ? `onclick='openInventoryEditModal("${escapeHtml(item.ID)}")'` : 'disabled'} title="Edit">
             <i class="fa-solid fa-pen-to-square"></i>
           </button>
           <button class="p-2 rounded-lg border border-gray-200 text-red-600${hasId ? ' hover:bg-red-50' : ' opacity-40 cursor-not-allowed'}" ${hasId ? `onclick='deleteInventoryItem("${escapeHtml(item.ID)}")'` : 'disabled'} title="Delete">
             <i class="fa-solid fa-trash"></i>
           </button>
         </div>
       </td>
     `;
     tbody.appendChild(tr);
  });
}

const takeOneInFlightIds = new Set();

async function takeOneInventoryItem(id) {
  const itemId = String(id || '').trim();
  if (!itemId) return;
  if (takeOneInFlightIds.has(itemId)) return;

  const current = (Array.isArray(inventoryPagination.allItems) ? inventoryPagination.allItems : []).find(it => it.ID === itemId);
  const qtyNum = Number(current?.Qty);
  if (Number.isFinite(qtyNum) && qtyNum <= 0) return;

  takeOneInFlightIds.add(itemId);
  try {
    const res = await callApi('adjustStock', { id: itemId, amount: -1, reason: 'Taken' });
    if (res?.success) await loadInventory();
  } finally {
    takeOneInFlightIds.delete(itemId);
  }
}

function openAdjustStockModal(id, name) {
  document.getElementById('adjust-item-id').value = id;
  document.getElementById('adjust-item-name').innerText = name;
  document.getElementById('adjust-amount').value = '';
  document.getElementById('adjust-reason').value = '';
  openModal('adjustStockModal');
}

async function submitStockAdjustment() {
  const id = document.getElementById('adjust-item-id').value;
  const amount = document.getElementById('adjust-amount').value;
  const reason = document.getElementById('adjust-reason').value;
  
  const res = await callApi('adjustStock', { id, amount, reason });
  if(res.success) {
    closeModal('adjustStockModal');
    loadInventory();
  } else {
    alert('Error: ' + res.message);
  }
}

// --- Reports Logic ---
async function loadReports() {
  const root = document.getElementById('reports-content');
  if (!root) return;

  const [items, stats] = await Promise.all([
    callApi('getInventory'),
    callApi('getDashboardStats')
  ]);
  if (items?.error || stats?.error) return;

  reportsCache = {
    loaded: true,
    items: Array.isArray(items) ? items.map(normalizeInventoryItem) : [],
    stats
  };

  renderReportsView(String(globalSearchState.query || '').trim());
}

function renderReportsView(query) {
  const root = document.getElementById('reports-content');
  if (!root) return;
  if (!reportsCache?.loaded) return;

  const initPrintFilters = () => {
    const s = document.getElementById('reports-print-summary');
    const l = document.getElementById('reports-print-lowstock');
    const a = document.getElementById('reports-print-activity');
    const reset = document.getElementById('reports-print-reset');

    if (s && !s.dataset.bound) {
      s.dataset.bound = '1';
      s.addEventListener('change', () => {
        reportsPrintOptions.summary = s.checked;
        renderReportsView(String(globalSearchState.query || '').trim());
      });
    }
    if (l && !l.dataset.bound) {
      l.dataset.bound = '1';
      l.addEventListener('change', () => {
        reportsPrintOptions.lowStock = l.checked;
        renderReportsView(String(globalSearchState.query || '').trim());
      });
    }
    if (a && !a.dataset.bound) {
      a.dataset.bound = '1';
      a.addEventListener('change', () => {
        reportsPrintOptions.activity = a.checked;
        renderReportsView(String(globalSearchState.query || '').trim());
      });
    }
    if (reset && !reset.dataset.bound) {
      reset.dataset.bound = '1';
      reset.addEventListener('click', () => {
        reportsPrintOptions.summary = true;
        reportsPrintOptions.lowStock = true;
        reportsPrintOptions.activity = true;
        if (s) s.checked = true;
        if (l) l.checked = true;
        if (a) a.checked = true;
        renderReportsView(String(globalSearchState.query || '').trim());
      });
    }

    if (s) s.checked = reportsPrintOptions.summary;
    if (l) l.checked = reportsPrintOptions.lowStock;
    if (a) a.checked = reportsPrintOptions.activity;
  };

  initPrintFilters();

  const toFiniteNumber = value => {
    const cleaned = String(value ?? '').replace(/[^0-9.\-]+/g, '');
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : 0;
  };

  const items = Array.isArray(reportsCache.items) ? reportsCache.items : [];
  const stats = reportsCache.stats || null;

  const fmtInt = n => Math.round(n).toLocaleString();
  const fmtCurrency = n => toPeso(n);

  const totalItems = items.length;
  const totalValue = items.reduce((sum, it) => sum + toFiniteNumber(it.Qty) * toFiniteNumber(it.UnitCost), 0);
  const lowStockAll = items.filter(it => {
    const qty = toFiniteNumber(it.Qty);
    return qty > 0 && qty < 10;
  });
  const outStockAll = items.filter(it => toFiniteNumber(it.Qty) <= 0);

  const q = String(query || '').trim().toLowerCase();
  const terms = q.split(/\s+/).filter(Boolean);
  const matchesTerms = hay => terms.length === 0 || terms.every(t => hay.includes(t));

  const lowStockFiltered = lowStockAll.filter(it => {
    const hay = Object.values(it)
      .map(v => String(v ?? ''))
      .join(' ')
      .toLowerCase();
    return matchesTerms(hay);
  });

  const allActivities = Array.isArray(stats?.recentActivities) ? stats.recentActivities : [];
  const activitiesFiltered = allActivities.filter(act => {
    const hay = [
      act?.Type,
      act?.ItemName,
      act?.Notes,
      act?.User,
      act?.Quantity,
      act?.Date,
      act?.Timestamp
    ]
      .map(v => String(v ?? ''))
      .join(' ')
      .toLowerCase();
    return matchesTerms(hay);
  });

  const syncEl = document.getElementById('reports-last-sync');
  if (syncEl) syncEl.innerText = new Date().toLocaleString();

  const printDateEl = document.getElementById('print-date');
  if (printDateEl) printDateEl.innerText = new Date().toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' });

  const totalItemsEl = document.getElementById('reports-total-items');
  if (totalItemsEl) totalItemsEl.innerText = fmtInt(totalItems);

  const totalValueEl = document.getElementById('reports-total-value');
  if (totalValueEl) totalValueEl.innerText = fmtCurrency(totalValue);

  const lowStockEl = document.getElementById('reports-low-stock');
  if (lowStockEl) lowStockEl.innerText = fmtInt(lowStockAll.length);

  const outStockEl = document.getElementById('reports-out-stock');
  if (outStockEl) outStockEl.innerText = fmtInt(outStockAll.length);

  const summaryLabel = document.getElementById('reports-summary-label');
  if (summaryLabel) summaryLabel.classList.toggle('hidden', !reportsPrintOptions.summary);

  const summaryCard = document.getElementById('reports-summary-card');
  if (summaryCard) summaryCard.classList.toggle('hidden', !reportsPrintOptions.summary);

  const lowCard = document.getElementById('reports-lowstock-card');
  if (lowCard) lowCard.classList.toggle('hidden', !reportsPrintOptions.lowStock);

  const actCard = document.getElementById('reports-activity-card');
  if (actCard) actCard.classList.toggle('hidden', !reportsPrintOptions.activity);

  const lowCountEl = document.getElementById('reports-lowstock-count');
  if (lowCountEl) lowCountEl.innerText = fmtInt(lowStockFiltered.length);

  const initLowStockPager = () => {
    const prevBtn = document.getElementById('reports-lowstock-prev-btn');
    if (prevBtn && !prevBtn.dataset.bound) {
      prevBtn.dataset.bound = '1';
      prevBtn.addEventListener('click', () => {
        if (reportsLowStockPagination.currentPage <= 1) return;
        reportsLowStockPagination.currentPage -= 1;
        renderReportsView(String(globalSearchState.query || '').trim());
      });
    }

    const nextBtn = document.getElementById('reports-lowstock-next-btn');
    if (nextBtn && !nextBtn.dataset.bound) {
      nextBtn.dataset.bound = '1';
      nextBtn.addEventListener('click', () => {
        reportsLowStockPagination.currentPage += 1;
        renderReportsView(String(globalSearchState.query || '').trim());
      });
    }
  };

  if (reportsPrintOptions.lowStock) initLowStockPager();

  const lowBody = document.getElementById('reports-lowstock-body');
  if (lowBody && reportsPrintOptions.lowStock) {
    const qKey = q;
    if (reportsLowStockPagination.lastQuery !== qKey) {
      reportsLowStockPagination.lastQuery = qKey;
      reportsLowStockPagination.currentPage = 1;
    }

    const ordered = lowStockFiltered
      .slice()
      .sort((a, b) => toFiniteNumber(a.Qty) - toFiniteNumber(b.Qty));

    const total = ordered.length;
    const totalPages = Math.ceil(total / reportsLowStockPagination.itemsPerPage) || 1;
    if (reportsLowStockPagination.currentPage > totalPages) reportsLowStockPagination.currentPage = totalPages;
    if (reportsLowStockPagination.currentPage < 1) reportsLowStockPagination.currentPage = 1;

    const startIdx = (reportsLowStockPagination.currentPage - 1) * reportsLowStockPagination.itemsPerPage;
    const endIdx = startIdx + reportsLowStockPagination.itemsPerPage;
    const shown = ordered.slice(startIdx, endIdx);

    lowBody.innerHTML = '';
    if (shown.length === 0) {
      lowBody.innerHTML = `
        <tr>
          <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500" colspan="4">No low stock items.</td>
        </tr>
      `;
    } else {
      shown.forEach(it => {
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-gray-50';
        tr.innerHTML = `
          <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${escapeHtml(it.Item)}</td>
          <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${escapeHtml(it.Category)}</td>
          <td class="px-6 py-4 whitespace-nowrap text-sm text-amber-700 font-semibold">${fmtInt(toFiniteNumber(it.Qty))}</td>
          <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${escapeHtml(it.Location)}</td>
        `;
        lowBody.appendChild(tr);
      });
    }

    const pageInfo = document.getElementById('reports-lowstock-page-info');
    if (pageInfo) {
      const start = total === 0 ? 0 : startIdx + 1;
      const end = Math.min(endIdx, total);
      pageInfo.innerText = `Page ${reportsLowStockPagination.currentPage} of ${totalPages} • Showing ${start}-${end} of ${total}`;
    }

    const prevBtn = document.getElementById('reports-lowstock-prev-btn');
    if (prevBtn) prevBtn.disabled = reportsLowStockPagination.currentPage <= 1;
    const nextBtn = document.getElementById('reports-lowstock-next-btn');
    if (nextBtn) nextBtn.disabled = reportsLowStockPagination.currentPage >= totalPages;
  }

  const actList = document.getElementById('reports-activity-list');
  if (actList && reportsPrintOptions.activity) {
    actList.innerHTML = '';
    if (activitiesFiltered.length === 0) {
      actList.innerHTML = '<li class="py-3 text-sm text-gray-500">No recent activity.</li>';
    } else {
      activitiesFiltered.forEach(act => {
        const li = document.createElement('li');
        li.className = 'py-3';
        const label = escapeHtml(String(act?.Type ?? '').trim());
        const name = escapeHtml(String(act?.ItemName ?? '').trim());
        const qty = escapeHtml(String(act?.Quantity ?? '').trim());
        const when = act?.Timestamp ? new Date(act.Timestamp).toLocaleString() : (act?.Date ? String(act.Date) : '');
        li.innerHTML = `
          <div class="flex items-start justify-between gap-3">
            <div class="min-w-0">
              <div class="text-sm font-medium text-gray-900 truncate">${label || 'Activity'}${name ? ` — ${name}` : ''}</div>
              <div class="text-xs text-gray-500 truncate">${qty ? `Qty: ${qty}` : ''}${act?.User ? ` • ${escapeHtml(act.User)}` : ''}</div>
            </div>
            <div class="text-xs text-gray-400 whitespace-nowrap">${when || ''}</div>
          </div>
        `;
        actList.appendChild(li);
      });
    }
  }
}

function applyReportsSearch(query) {
  if (!reportsCache?.loaded) return;
  renderReportsView(query);
}

// --- Special Reports Logic ---

function openSpecialReportModal() {
  const modal = document.getElementById('specialReportModal');
  if (!modal) return;

  // Populate Options
  const items = Array.isArray(inventoryPagination.allItems) ? inventoryPagination.allItems : [];
  
  // People
  const people = [...new Set(items.map(i => i.PersonInCharge).filter(Boolean))].sort();
  const personSel = document.getElementById('report-person');
  if(personSel) {
    personSel.innerHTML = '<option value="">Select Person...</option>' + 
        people.map(p => `<option value="${p}">${p}</option>`).join('');
  }

  // Locations
  const locations = [...new Set(items.map(i => i.Location).filter(Boolean))].sort();
  const locSel = document.getElementById('report-location');
  if(locSel) {
    locSel.innerHTML = '<option value="all">All Locations</option>' + 
        locations.map(l => `<option value="${l}">${l}</option>`).join('');
  }

  // Projects
  const projects = [...new Set(items.map(i => i.Project).filter(Boolean))].sort();
  const projSel = document.getElementById('report-project');
  if(projSel) {
    projSel.innerHTML = '<option value="">Select Project...</option>' + 
        projects.map(p => `<option value="${p}">${p}</option>`).join('');
  }

  updateReportOptions();
  modal.classList.remove('hidden');
}

function updateReportOptions() {
  const type = document.getElementById('report-type').value;
  const descEl = document.getElementById('report-description');
  document.querySelectorAll('.report-option').forEach(el => el.classList.add('hidden'));

  if (type === 'par') {
    document.getElementById('report-option-person').classList.remove('hidden');
    if(descEl) descEl.innerHTML = '<strong>Purpose:</strong> Used to assign accountability to an employee.<br>Generates a receipt for equipment issued to a specific person.';
  } else if (type === 'count') {
    document.getElementById('report-option-location').classList.remove('hidden');
    if(descEl) descEl.innerHTML = '<strong>Purpose:</strong> Used for physical inventory auditing.<br>Generates a checklist with blank columns for "Quantity per Count" and "Remarks".';
  } else if (type === 'valuation') {
    if(descEl) descEl.innerHTML = '<strong>Purpose:</strong> Used for accounting and financial reporting.<br>Summarizes the total value of assets based on acquisition cost.';
  } else if (type === 'project') {
    document.getElementById('report-option-project').classList.remove('hidden');
    if(descEl) descEl.innerHTML = '<strong>Purpose:</strong> Used to track assets deployed to specific projects.<br>Shows total items and value allocated to a project.';
  }
}

function generateSpecialReport() {
  const type = document.getElementById('report-type').value;
  const officer = escapeHtml(document.getElementById('report-officer').value || '__________________');
  const items = Array.isArray(inventoryPagination.allItems) ? inventoryPagination.allItems : [];
  
  let html = '';
  let title = '';

  const toCurrency = (val) => '₱' + (Number(val) || 0).toLocaleString('en-PH', {minimumFractionDigits: 2});

  if (type === 'par') {
    const personRaw = document.getElementById('report-person').value;
    if (!personRaw) { alert('Please select a Person In Charge'); return; }
    const person = escapeHtml(personRaw);
    
    title = 'Property Acknowledgement Receipt';
    const filtered = items.filter(i => i.PersonInCharge === personRaw);
    
    html = `
      <div class="text-center mb-8">
        <h1 class="text-2xl font-bold uppercase mb-2">Property Acknowledgement Receipt</h1>
        <p class="text-sm text-gray-600">Department of Justice - NBP/GovNet</p>
      </div>
      
      <div class="mb-6 flex justify-between text-sm">
        <div><strong>Entity Name:</strong> DOJ - NBP</div>
        <div><strong>Par No.:</strong> ________</div>
      </div>

      <table class="w-full border-collapse border border-gray-300 text-sm mb-8">
        <thead>
          <tr class="bg-gray-100">
            <th class="border border-gray-300 p-2 text-center">Qty</th>
            <th class="border border-gray-300 p-2 text-center">Unit</th>
            <th class="border border-gray-300 p-2 text-left">Description</th>
            <th class="border border-gray-300 p-2 text-left">Property Number/Serial</th>
            <th class="border border-gray-300 p-2 text-center">Date Acquired</th>
            <th class="border border-gray-300 p-2 text-right">Amount</th>
          </tr>
        </thead>
        <tbody>
          ${filtered.map(i => `
            <tr>
              <td class="border border-gray-300 p-2 text-center">${escapeHtml(i.Qty)}</td>
              <td class="border border-gray-300 p-2 text-center">${escapeHtml(i.Unit)}</td>
              <td class="border border-gray-300 p-2">
                <div class="font-semibold">${escapeHtml(i.Item)}</div>
                <div class="text-xs text-gray-500">${escapeHtml(i.BrandModel)}</div>
              </td>
              <td class="border border-gray-300 p-2">${escapeHtml(i.Serial)}</td>
              <td class="border border-gray-300 p-2 text-center">${escapeHtml(i.DateAcquired)}</td>
              <td class="border border-gray-300 p-2 text-right">${toCurrency(i.UnitCost)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>

      <div class="grid grid-cols-2 gap-12 mt-12 page-break-inside-avoid">
        <div>
          <div class="text-sm font-semibold mb-8">Received by:</div>
          <div class="border-b border-black font-bold text-center uppercase py-1">${person}</div>
          <div class="text-center text-xs mt-1">Signature over Printed Name of End User</div>
          <div class="text-center text-xs mt-4">Position/Office: __________________</div>
          <div class="text-center text-xs">Date: __________________</div>
        </div>
        <div>
          <div class="text-sm font-semibold mb-8">Issued by:</div>
          <div class="border-b border-black font-bold text-center uppercase py-1">${officer}</div>
          <div class="text-center text-xs mt-1">Signature over Printed Name of Supply Officer</div>
          <div class="text-center text-xs mt-4">Position/Office: Property/Supply Officer</div>
          <div class="text-center text-xs">Date: __________________</div>
        </div>
      </div>
    `;

  } else if (type === 'count') {
    const locRaw = document.getElementById('report-location').value;
    title = 'Physical Inventory Count Sheet';
    
    let filtered = items;
    if (locRaw && locRaw !== 'all') {
      filtered = items.filter(i => i.Location === locRaw);
    }
    
    // Group by Location
    const grouped = filtered.reduce((acc, item) => {
      const l = item.Location || 'Unassigned';
      if (!acc[l]) acc[l] = [];
      acc[l].push(item);
      return acc;
    }, {});

    html = `
      <div class="text-center mb-8">
        <h1 class="text-2xl font-bold uppercase mb-2">Report on the Physical Count of Property, Plant and Equipment</h1>
        <p class="text-sm text-gray-600">As of ${new Date().toLocaleDateString()}</p>
      </div>

      ${Object.entries(grouped).map(([location, groupItems]) => `
        <div class="mb-8 page-break-inside-avoid">
          <h3 class="font-bold text-lg mb-2 border-b border-gray-200 pb-1">Location: ${escapeHtml(location)}</h3>
          <table class="w-full border-collapse border border-gray-300 text-sm">
            <thead>
              <tr class="bg-gray-100">
                <th class="border border-gray-300 p-2 text-left">Article/Description</th>
                <th class="border border-gray-300 p-2 text-left">Property No./Serial</th>
                <th class="border border-gray-300 p-2 text-center">Unit of Measure</th>
                <th class="border border-gray-300 p-2 text-center">Unit Value</th>
                <th class="border border-gray-300 p-2 text-center">Quantity per Card</th>
                <th class="border border-gray-300 p-2 text-center">Quantity per Count</th>
                <th class="border border-gray-300 p-2 text-left w-32">Remarks</th>
              </tr>
            </thead>
            <tbody>
              ${groupItems.map(i => `
                <tr>
                  <td class="border border-gray-300 p-2">
                    <div class="font-semibold">${escapeHtml(i.Item)}</div>
                    <div class="text-xs text-gray-500">${escapeHtml(i.BrandModel)}</div>
                  </td>
                  <td class="border border-gray-300 p-2">${escapeHtml(i.Serial)}</td>
                  <td class="border border-gray-300 p-2 text-center">${escapeHtml(i.Unit)}</td>
                  <td class="border border-gray-300 p-2 text-center">${toCurrency(i.UnitCost)}</td>
                  <td class="border border-gray-300 p-2 text-center">${escapeHtml(i.Qty)}</td>
                  <td class="border border-gray-300 p-2"></td> <!-- Blank for manual count -->
                  <td class="border border-gray-300 p-2"></td> <!-- Blank for remarks -->
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      `).join('')}
      
      <div class="mt-12">
        <div class="text-sm font-semibold mb-8">Certified Correct by:</div>
        <div class="flex gap-12">
          <div class="flex-1">
             <div class="border-b border-black font-bold text-center uppercase py-1">${officer}</div>
             <div class="text-center text-xs mt-1">Inventory Committee Chair</div>
          </div>
          <div class="flex-1">
             <div class="border-b border-black h-6"></div>
             <div class="text-center text-xs mt-1">COA Representative</div>
          </div>
        </div>
      </div>
    `;

  } else if (type === 'valuation') {
    title = 'Asset Valuation Report';
    
    let totalVal = 0;
    const rows = items.map(i => {
      const cost = Number(i.UnitCost) || 0;
      const qty = Number(i.Qty) || 0;
      const sub = cost * qty;
      totalVal += sub;
      return { ...i, sub };
    });

    html = `
      <div class="text-center mb-8">
        <h1 class="text-2xl font-bold uppercase mb-2">Asset Depreciation & Valuation Report</h1>
        <p class="text-sm text-gray-600">Summary of Asset Values</p>
      </div>

      <table class="w-full border-collapse border border-gray-300 text-sm mb-8">
        <thead>
          <tr class="bg-gray-100">
            <th class="border border-gray-300 p-2 text-left">Item Description</th>
            <th class="border border-gray-300 p-2 text-left">Serial No.</th>
            <th class="border border-gray-300 p-2 text-center">Date Acquired</th>
            <th class="border border-gray-300 p-2 text-center">Qty</th>
            <th class="border border-gray-300 p-2 text-right">Unit Cost</th>
            <th class="border border-gray-300 p-2 text-right">Total Cost</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map(i => `
            <tr>
              <td class="border border-gray-300 p-2">
                <div class="font-semibold">${escapeHtml(i.Item)}</div>
                <div class="text-xs text-gray-500">${escapeHtml(i.BrandModel)}</div>
              </td>
              <td class="border border-gray-300 p-2">${escapeHtml(i.Serial)}</td>
              <td class="border border-gray-300 p-2 text-center">${escapeHtml(i.DateAcquired)}</td>
              <td class="border border-gray-300 p-2 text-center">${escapeHtml(i.Qty)}</td>
              <td class="border border-gray-300 p-2 text-right">${toCurrency(i.UnitCost)}</td>
              <td class="border border-gray-300 p-2 text-right font-medium">${toCurrency(i.sub)}</td>
            </tr>
          `).join('')}
        </tbody>
        <tfoot>
          <tr class="bg-gray-100 font-bold">
            <td class="border border-gray-300 p-2 text-right" colspan="5">GRAND TOTAL</td>
            <td class="border border-gray-300 p-2 text-right">${toCurrency(totalVal)}</td>
          </tr>
        </tfoot>
      </table>
      
      <div class="mt-8 text-xs text-gray-500 italic">
        * Depreciation values are subject to standard accounting rules based on Date Acquired.
      </div>
    `;

  } else if (type === 'project') {
    const projRaw = document.getElementById('report-project').value;
    if (!projRaw) { alert('Please select a Project'); return; }
    const proj = escapeHtml(projRaw);
    
    title = `Project Asset Allocation: ${projRaw}`;
    const filtered = items.filter(i => i.Project === projRaw);
    
    const deployed = filtered.reduce((acc, i) => acc + (Number(i.Qty) || 0), 0);

    html = `
      <div class="text-center mb-8">
        <h1 class="text-2xl font-bold uppercase mb-2">Project Asset Allocation Report</h1>
        <h2 class="text-xl font-semibold text-indigo-700">${proj}</h2>
      </div>

      <div class="mb-8 grid grid-cols-2 gap-4 bg-gray-50 p-4 rounded-xl border border-gray-200">
        <div>
          <div class="text-xs text-gray-500 uppercase tracking-wider">Total Items Allocated</div>
          <div class="text-2xl font-bold text-gray-900">${deployed}</div>
        </div>
        <div>
          <div class="text-xs text-gray-500 uppercase tracking-wider">Total Value</div>
          <div class="text-2xl font-bold text-gray-900">${toCurrency(filtered.reduce((acc, i) => acc + (Number(i.Qty)*Number(i.UnitCost)||0), 0))}</div>
        </div>
      </div>

      <table class="w-full border-collapse border border-gray-300 text-sm mb-8">
        <thead>
          <tr class="bg-gray-100">
            <th class="border border-gray-300 p-2 text-left">Item</th>
            <th class="border border-gray-300 p-2 text-left">Person In Charge</th>
            <th class="border border-gray-300 p-2 text-center">Qty</th>
            <th class="border border-gray-300 p-2 text-left">Current Location</th>
            <th class="border border-gray-300 p-2 text-left">Status</th>
          </tr>
        </thead>
        <tbody>
          ${filtered.map(i => `
            <tr>
              <td class="border border-gray-300 p-2">
                <div class="font-semibold">${escapeHtml(i.Item)}</div>
                <div class="text-xs text-gray-500">${escapeHtml(i.BrandModel)}</div>
              </td>
              <td class="border border-gray-300 p-2">${escapeHtml(i.PersonInCharge)}</td>
              <td class="border border-gray-300 p-2 text-center">${escapeHtml(i.Qty)}</td>
              <td class="border border-gray-300 p-2">${escapeHtml(i.Location)}</td>
              <td class="border border-gray-300 p-2">
                <span class="px-2 py-1 rounded-full text-xs font-semibold ${i.Qty > 0 ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'}">
                  ${escapeHtml(i.Status || (i.Qty > 0 ? 'Active' : 'Out of Stock'))}
                </span>
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;
  }

  // Inject content
  const container = document.getElementById('reports-generated-content');
  if (container) container.innerHTML = html;

  // Toggle Views
  document.getElementById('reports-default-view').classList.add('hidden');
  document.getElementById('reports-generated-view').classList.remove('hidden');
  
  // Close Modal
  closeModal('specialReportModal');
}

function closeGeneratedReport() {
  document.getElementById('reports-generated-view').classList.add('hidden');
  document.getElementById('reports-default-view').classList.remove('hidden');
}

// --- Utils ---
function openModal(modalId) {
  document.getElementById(modalId).classList.remove('hidden');
}

function closeModal(modalId) {
  document.getElementById(modalId).classList.add('hidden');
}

// --- Mobile Navigation ---
function initMobileMenu() {
  const menuBtn = document.getElementById('mobile-menu-btn');
  const sidebar = document.getElementById('sidebar');
  
  if (menuBtn && sidebar) {
    menuBtn.addEventListener('click', () => {
      sidebar.classList.toggle('-translate-x-full');
    });
    
    // Close menu when clicking outside on mobile
    document.addEventListener('click', (e) => {
      if (window.innerWidth < 1024 && 
          !sidebar.contains(e.target) && 
          !menuBtn.contains(e.target) &&
          !sidebar.classList.contains('-translate-x-full')) {
        sidebar.classList.add('-translate-x-full');
      }
    });
  }
}

function initSidebarCollapse() {
  const sidebar = document.getElementById('sidebar');
  const btn = document.getElementById('sidebar-collapse-btn');
  const icon = document.getElementById('sidebar-collapse-icon');
  if (!sidebar || !btn) return;

  const applyCollapsed = collapsed => {
    if (collapsed) {
      sidebar.classList.add('w-20', 'sidebar-collapsed');
      sidebar.classList.remove('w-64');
      if (icon) icon.classList.add('rotate-180');
      btn.setAttribute('aria-expanded', 'false');
      btn.setAttribute('title', 'Expand sidebar');
    } else {
      sidebar.classList.add('w-64');
      sidebar.classList.remove('w-20', 'sidebar-collapsed');
      if (icon) icon.classList.remove('rotate-180');
      btn.setAttribute('aria-expanded', 'true');
      btn.setAttribute('title', 'Collapse sidebar');
    }
  };

  const saved = localStorage.getItem('sidebarCollapsed') === '1';
  if (window.innerWidth >= 1024) applyCollapsed(saved);

  btn.addEventListener('click', () => {
    const willCollapse = !sidebar.classList.contains('sidebar-collapsed');
    applyCollapsed(willCollapse);
    localStorage.setItem('sidebarCollapsed', willCollapse ? '1' : '0');
  });

  window.addEventListener('resize', () => {
    if (window.innerWidth < 1024) {
      applyCollapsed(false);
      return;
    }
    const next = localStorage.getItem('sidebarCollapsed') === '1';
    applyCollapsed(next);
  });
}

// --- Chat Widget Logic ---

function autoResizeChatInput(el) {
  el.style.height = 'auto';
  el.style.height = el.scrollHeight + 'px';
}

function handleChatKeydown(e) {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    handleChatSubmit(e);
  }
}

function toggleChat() {
  const container = document.getElementById('chat-container');
  const btn = document.getElementById('chat-toggle-btn');
  
  if (container.classList.contains('hidden')) {
    // Open
    container.classList.remove('hidden');
    // Small delay to allow transition to work
    setTimeout(() => {
        container.classList.remove('scale-95', 'opacity-0');
        container.classList.add('scale-100', 'opacity-100');
    }, 10);
    // Focus input
    setTimeout(() => {
        document.getElementById('chat-input').focus();
    }, 300);
  } else {
    // Close
    container.classList.remove('scale-100', 'opacity-100');
    container.classList.add('scale-95', 'opacity-0');
    setTimeout(() => {
        container.classList.add('hidden');
    }, 300);
  }
}

function handleChatSubmit(e) {
  e.preventDefault();
  const input = document.getElementById('chat-input');
  const message = input.value.trim();
  if (!message) return;

  // Add user message
  addChatMessage(message, 'user');
  input.value = '';
  input.style.height = 'auto'; // Reset height

  // Simulate AI processing
  showChatTyping();
  
  // Simple heuristic response logic
  setTimeout(async () => {
    const response = await generateAIResponse(message);
    hideChatTyping();
    addChatMessage(response, 'ai');
  }, 1000 + Math.random() * 1000);
}

function addChatMessage(text, sender) {
  const container = document.getElementById('chat-messages');
  const isUser = sender === 'user';
  
  const div = document.createElement('div');
  div.className = `flex items-start gap-2.5 ${isUser ? 'flex-row-reverse' : ''}`;
  
  const avatar = isUser 
    ? `<div class="w-8 h-8 rounded-full bg-indigo-600 flex items-center justify-center flex-shrink-0 text-white text-xs">Me</div>`
    : `<div class="w-8 h-8 rounded-full bg-blue-100 flex items-center justify-center flex-shrink-0">
         <svg class="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
           <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9.75 17L9 20l-1 1h8l-1-1-.75-3M3 13h18M5 17h14a2 2 0 002-2V5a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"></path>
         </svg>
       </div>`;

  const bubbleClass = isUser
    ? 'bg-blue-600 text-white rounded-tr-none shadow-md shadow-blue-600/10'
    : 'bg-white rounded-tl-none shadow-sm border border-gray-100 text-gray-600';

  div.innerHTML = `
    ${avatar}
    <div class="flex flex-col gap-1 max-w-[85%] ${isUser ? 'items-end' : 'items-start'}">
      <div class="p-3 rounded-2xl text-sm ${bubbleClass}">
        ${text.replace(/\n/g, '<br>')}
      </div>
      <span class="text-[10px] text-gray-400 mx-1">Just now</span>
    </div>
  `;
  
  container.appendChild(div);
  container.scrollTop = container.scrollHeight;
}

let chatTypingEl = null;

function showChatTyping() {
  const container = document.getElementById('chat-messages');
  if (chatTypingEl) return;
  
  chatTypingEl = document.createElement('div');
  chatTypingEl.className = 'flex items-start gap-2.5';
  chatTypingEl.innerHTML = `
    <div class="w-8 h-8 rounded-full bg-blue-100 flex items-center justify-center flex-shrink-0">
      <svg class="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9.75 17L9 20l-1 1h8l-1-1-.75-3M3 13h18M5 17h14a2 2 0 002-2V5a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"></path>
      </svg>
    </div>
    <div class="flex flex-col gap-1 max-w-[85%]">
      <div class="bg-white p-3 rounded-2xl rounded-tl-none shadow-sm border border-gray-100 text-sm text-gray-600">
        <div class="flex gap-1">
          <span class="w-1.5 h-1.5 bg-gray-400 rounded-full animate-bounce"></span>
          <span class="w-1.5 h-1.5 bg-gray-400 rounded-full animate-bounce" style="animation-delay: 0.1s"></span>
          <span class="w-1.5 h-1.5 bg-gray-400 rounded-full animate-bounce" style="animation-delay: 0.2s"></span>
        </div>
      </div>
    </div>
  `;
  container.appendChild(chatTypingEl);
  container.scrollTop = container.scrollHeight;
}

function hideChatTyping() {
  if (chatTypingEl) {
    chatTypingEl.remove();
    chatTypingEl = null;
  }
}

function buildGeminiSystemPrompt_() {
  return [
    'Ikaw ay Inventory Assistant para sa web-based inventory system.',
    'Sagutin mo nang maikli, malinaw, at direktang sagot.',
    'Gamitin ang ibinigay na context kung meron.',
    'Kung kulang ang data, sabihin na kulang at kung ano ang kailangan.',
    'Huwag mag-imbento ng values.'
  ].join('\n');
}

function buildGeminiContext_() {
  const items = Array.isArray(inventoryPagination?.allItems) ? inventoryPagination.allItems : [];
  const toNum = v => {
    const n = Number(String(v ?? '').replace(/[^0-9.\-]+/g, ''));
    return Number.isFinite(n) ? n : 0;
  };

  const total = items.length;
  const lowStock = items.filter(it => {
    const qty = toNum(it?.Qty);
    return qty > 0 && qty < 10;
  }).length;
  const outStock = items.filter(it => toNum(it?.Qty) <= 0).length;

  const sample = items
    .slice(0, 25)
    .map(it => `${String(it?.Item ?? '').trim()} | Qty: ${String(it?.Qty ?? '').trim()} | Serial: ${String(it?.Serial ?? '').trim()} | Location: ${String(it?.Location ?? '').trim()}`)
    .filter(s => s.replace(/\s+/g, '').length > 0)
    .join('\n');

  return [
    `Total items: ${total}`,
    `Low stock (1-10): ${lowStock}`,
    `Out of stock (<=0): ${outStock}`,
    sample ? `Sample items (up to 25):\n${sample}` : ''
  ].filter(Boolean).join('\n');
}

async function generateAIResponse(msg) {
  let res;
  let geminiError = '';
  try {
    res = await callApi(
      'geminiChat',
      {
        message: msg,
        model: 'gemini-3-flash-preview',
        system: buildGeminiSystemPrompt_(),
        context: buildGeminiContext_(),
        includeBackendData: true
      },
      { silent: true }
    );
  } catch (e) {
    geminiError = e && e.message ? e.message : String(e);
  }

  if (res && res.success && typeof res.text === 'string' && res.text.trim()) return res.text;
  if (res && (res.message || res.error)) geminiError = String(res.message || res.error);

  const lower = msg.toLowerCase();
  
  // 1. Greetings
  if (lower.match(/^(hi|hello|hey|greetings)/)) {
    const fallback = "Hello! I'm here to help you manage your inventory. You can ask me about stock levels, low stock items, or general reports.";
    return geminiError ? `Gemini not available (${geminiError}).\n\n${fallback}` : fallback;
  }
  
  // 2. Audit & History Analysis (Who / When / Activity)
  if (lower.match(/(who|when|history|log|audit|action|deleted|added|edited|modified|activity|recent)/)) {
    // Check if logs are loaded, if not, fetch them
    if (!auditLogsCache || auditLogsCache.length === 0) {
       try {
         const logs = await callApi('getAuditLogs', null, { silent: true });
         if (!logs.error && Array.isArray(logs)) {
            auditLogsCache = logs.sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp)); // Sort newest first
         }
       } catch (e) {
         return "I tried to check the records, but I couldn't access the Audit Logs right now.";
       }
    }

    if (!auditLogsCache || auditLogsCache.length === 0) return "The audit logs appear to be empty.";

    // Case A: "Who deleted [item]?" or "When was [item] added?"
    const itemMatch = inventoryPagination.allItems.find(it => lower.includes(it.Item.toLowerCase()));
    const targetWord = itemMatch ? itemMatch.Item.toLowerCase() : lower.split(' ').find(w => w.length > 4 && !['history','audit','about','check'].includes(w));
    
    if (targetWord) {
        const matches = auditLogsCache.filter(log => {
            const content = (log.Details + ' ' + log.Action).toLowerCase();
            return content.includes(targetWord);
        });

        if (matches.length > 0) {
            const top = matches.slice(0, 3).map(l => {
                const date = new Date(l.Timestamp).toLocaleDateString();
                return `- ${date}: **${l.User}** performed **${l.Action}** (${l.Details})`;
            }).join('\n');
            return `Here is the history for "${targetWord}":\n${top}`;
        }
    }

    // Case B: General "Recent Activity"
    if (lower.includes('recent') || lower.includes('latest') || lower.includes('last')) {
        const top = auditLogsCache.slice(0, 5).map(l => {
             const time = new Date(l.Timestamp).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
             return `- ${time}: **${l.User}** - ${l.Details}`;
        }).join('\n');
        return `Here are the 5 most recent activities in the system:\n${top}`;
    }
  }

  // 3. Advanced Inventory Stats (Breakdowns)
  if (lower.includes('most') || lower.includes('highest') || lower.includes('breakdown') || lower.includes('summary')) {
      if (!inventoryPagination.allItems || inventoryPagination.allItems.length === 0) {
           // Try fetch
           try {
             const items = await callApi('getInventory', null, { silent: true });
             if (!items.error) inventoryPagination.allItems = Array.isArray(items) ? items.map(normalizeInventoryItem) : [];
           } catch(e) {}
      }
      
      const items = inventoryPagination.allItems;
      if (!items || items.length === 0) return "No inventory data available to analyze.";

      // "Which category has the most items?"
      if (lower.includes('category')) {
          const counts = {};
          items.forEach(it => { const c = it.Category || 'Uncategorized'; counts[c] = (counts[c] || 0) + 1; });
          const sorted = Object.entries(counts).sort((a,b) => b[1] - a[1]);
          const top = sorted[0];
          return `**${top[0]}** is the largest category with ${top[1]} items. The breakdown is:\n` + sorted.slice(0,3).map(s => `- ${s[0]}: ${s[1]}`).join('\n');
      }
      
      // "Which location has the most items?"
      if (lower.includes('location')) {
          const counts = {};
          items.forEach(it => { const l = it.Location || 'Unknown'; counts[l] = (counts[l] || 0) + 1; });
          const sorted = Object.entries(counts).sort((a,b) => b[1] - a[1]);
          const top = sorted[0];
          return `Most items are located in **${top[0]}** (${top[1]} items). Top locations:\n` + sorted.slice(0,3).map(s => `- ${s[0]}: ${s[1]}`).join('\n');
      }
  }

  // 4. Stock / Count / How many / List / Show / Find
  if (lower.match(/(how many|stock|count|check|list|show|find|search|where|what|which)/)) {
    // Check if we have data
    if (!inventoryPagination.allItems || inventoryPagination.allItems.length === 0) {
      // Try to load it if not loaded
       try {
         const items = await callApi('getInventory', null, { silent: true });
         if (!items.error) {
            inventoryPagination.allItems = Array.isArray(items) ? items.map(normalizeInventoryItem) : [];
         }
       } catch (e) {
         return "I'm having trouble accessing the inventory data right now. Please try again later.";
       }
    }
    
    // Find item name in query
    const items = inventoryPagination.allItems;
    if (items.length === 0) return "I don't see any items in the inventory yet.";

    // General stock status
    if (lower.includes('low stock')) {
        const low = items.filter(it => Number(it.Qty) > 0 && Number(it.Qty) < 10);
        return `There are currently ${low.length} items marked as Low Stock. Check the Dashboard or Reports for details.`;
    }
    
    if (lower.includes('out of stock') || lower.includes('out stock')) {
        const out = items.filter(it => Number(it.Qty) <= 0);
        return `There are currently ${out.length} items Out of Stock.`;
    }

    if (lower.includes('all items') || lower.includes('everything') || lower.includes('total items')) {
        return `We have a total of ${items.length} unique items in the inventory.`;
    }

    // Improved Keyword Matching
    // Filter out common stop words to focus on the actual query
    const stopWords = ['what', 'where', 'is', 'the', 'have', 'check', 'stock', 'many', 'much', 'does', 'item', 'items', 'list', 'show', 'find', 'search', 'in', 'at', 'of', 'for'];
    const words = lower.split(' ').filter(w => w.length > 2 && !stopWords.includes(w));
    
    if (words.length > 0) {
        const matches = items.filter(it => {
            // Search across multiple fields for better "System-wide" analysis
            const searchableText = [
                it.Item,
                it.BrandModel,
                it.Category,
                it.Location,
                it.Status,
                it.Description,
                it.Serial
            ].map(val => String(val || '').toLowerCase()).join(' ');

            // AND logic: all words must be present in the item's data
            return words.every(w => searchableText.includes(w));
        });
        
        if (matches.length > 0) {
            if (matches.length === 1) {
                const item = matches[0];
                return `Found it: **${item.Item}** (${item.BrandModel || 'N/A'})\n` +
                       `- Stock: ${item.Qty} ${item.Unit}\n` +
                       `- Location: ${item.Location || 'Unknown'}\n` +
                       `- Category: ${item.Category || 'General'}\n` +
                       `- Status: ${item.Status || 'Active'}`;
            } else {
                // Summarize findings
                const totalQty = matches.reduce((sum, it) => sum + Number(it.Qty || 0), 0);
                const locations = [...new Set(matches.map(it => it.Location).filter(Boolean))].join(', ');
                
                // Top 5 items
                const top = matches.slice(0, 5).map(it => `- ${it.Item} (${it.BrandModel || ''}): ${it.Qty}`).join('\n');
                
                return `I found ${matches.length} items matching "${words.join(' ')}":\n` +
                       `Total Quantity: ${totalQty}\n` +
                       `Locations: ${locations || 'Various'}\n\n` +
                       `Here are the top results:\n${top}${matches.length > 5 ? `\n...and ${matches.length - 5} more.` : ''}`;
            }
        }
    }

    return "I couldn't find any items matching your description. Try searching by Item Name, Brand, Category, or Location (e.g., 'Laptops in IT Office').";
  }
  
  // 5. Reports / Value
  if (lower.includes('value') || lower.includes('worth') || lower.includes('total')) {
      if (dashboardStatsCache) {
          return `The total estimated value of the inventory is ${toPeso(dashboardStatsCache.totalValue)}. We have ${dashboardStatsCache.totalItems} unique items tracked.`;
      }
      return "I can't see the dashboard stats right now. Please open the Dashboard first.";
  }

  // 6. System Knowledge Base (Dynamic Help)
  const knowledgeMatch = findBestKnowledgeMatch(lower);
  if (knowledgeMatch) {
    return knowledgeMatch;
  }



  // 7. Help / Capabilities
  if (lower.includes('help') || lower.includes('can you do')) {
    const fallback = "I can help you with:\n- Checking stock levels (e.g., 'How many printers?')\n- Finding item locations\n- Reporting total inventory value\n- Identifying low stock items\n- Answering questions about how to use the system";
    return geminiError ? `Gemini not available (${geminiError}).\n\n${fallback}` : fallback;
  }

  const fallback = "I'm not sure I understand. I can help you check inventory stock, find items, get status reports, or answer questions about how to use the system. Try asking 'How do I add an item?' or 'What is the stock of Mouse?'";
  return geminiError ? `Gemini not available (${geminiError}).\n\n${fallback}` : fallback;
}

// --- System Knowledge Base ---
const SYSTEM_KNOWLEDGE = [
  {
    keywords: ['add', 'create', 'new', 'item', 'inventory'],
    answer: "To add a new item:\n1. Go to the **Inventory** page.\n2. Click the **Add New Item** button (top right).\n3. Fill in the details (Item Name, Brand, Category, etc.).\n4. Click **Save Item**."
  },
  {
    keywords: ['delete', 'remove', 'trash', 'dispose', 'item'],
    answer: "To delete or dispose of an item:\n1. Go to the **Inventory** page.\n2. Find the item you want to remove.\n3. Click the **Trash Icon** (Delete) on the right side of the item row.\n4. Confirm the deletion."
  },
  {
    keywords: ['edit', 'update', 'change', 'modify', 'details'],
    answer: "To edit an item:\n1. Go to the **Inventory** page.\n2. Find the item you want to edit.\n3. Click the **Pencil Icon** (Edit) on the right side of the item row.\n4. Update the details in the popup form and click **Save Item**."
  },
  {
    keywords: ['adjust', 'stock', 'quantity', 'increase', 'decrease', 'qty'],
    answer: "To adjust stock levels:\n1. Go to the **Inventory** page.\n2. Find the item.\n3. Click the **Plus/Minus Icon** (Adjust Stock).\n4. Select 'Add Stock' or 'Reduce Stock', enter the quantity, and provide a reason.\n5. Click **Submit Adjustment**."
  },
  {
    keywords: ['print', 'export', 'pdf', 'excel', 'report', 'download'],
    answer: "To print or export reports:\n1. Go to the **Reports** page.\n2. Use the filters to select the data you need.\n3. Click the **Print / Export PDF** button at the top right of the report table."
  },
  {
    keywords: ['audit', 'log', 'history', 'track', 'who'],
    answer: "To view the audit logs (history of actions):\n1. Go to the **Audit Logs** page.\n2. You can see a list of all actions (Login, Add, Edit, Delete) with timestamps and user emails.\n3. Use the search bar to find specific events."
  },
  {
    keywords: ['dashboard', 'stats', 'summary', 'overview'],
    answer: "The **Dashboard** gives you a real-time overview of your inventory health, including:\n- Total Items & Total Value\n- Low Stock Alerts\n- Recent Activities\n- Stock Trends Chart"
  },
  {
    keywords: ['search', 'find', 'filter', 'lookup'],
    answer: "You can search for items using the **Global Search Bar** at the top of the screen. It works across all pages. You can also use specific filters (Category, Status, Stock Level) within the Inventory page."
  },
  {
    keywords: ['password', 'security', 'access', 'login'],
    answer: "The system is secured using a **Script ID** hash. If you are asked for security verification, enter the valid Script ID provided by your administrator. This ensures only authorized users can access the data."
  }
];

function findBestKnowledgeMatch(query) {
  const words = query.toLowerCase().split(/\s+/).filter(w => w.length > 2);
  let bestMatch = null;
  let maxScore = 0;

  for (const topic of SYSTEM_KNOWLEDGE) {
    let score = 0;
    // Calculate score based on keyword matches
    topic.keywords.forEach(keyword => {
      if (query.includes(keyword)) score += 2; // Exact phrase match bonus
      if (words.some(w => w.includes(keyword) || keyword.includes(w))) score += 1; // Partial match
    });

    if (score > maxScore && score >= 2) { // Minimum threshold
      maxScore = score;
      bestMatch = topic.answer;
    }
  }

  return bestMatch;
}



// Initialize mobile menu on page load
document.addEventListener('DOMContentLoaded', function() {
  initGlobalSearch();
  initMobileMenu();
  initSidebarCollapse();
  
  // Close print options dropdown when clicking outside
  document.addEventListener('click', (e) => {
    const menu = document.getElementById('print-options-menu');
    if (!menu || menu.classList.contains('hidden')) return;
    
    // Check if click is inside the menu
    if (menu.contains(e.target)) return;
    
    // Check if click is on the toggle button (assumed to be the button immediately preceding the menu container's parent or similar)
    // Actually, simpler: check if the click target is NOT the toggle button.
    // The toggle button has an onclick that handles the toggle.
    // If we click the toggle button, we don't want this listener to immediately hide it if the toggle just showed it.
    // But since the toggle is inline onclick, it runs first.
    // If it was hidden -> onclick removes hidden -> this listener sees it not hidden -> hides it again?
    // Yes, that's a race condition if propagation isn't stopped or logic isn't careful.
    
    // Let's look at the HTML structure again.
    // <button onclick="...">...</button>
    // <div id="print-options-menu">...</div>
    
    // If I click the button:
    // 1. onclick runs: toggles class. (Hidden -> Visible)
    // 2. document click runs:
    //    menu is Visible.
    //    click is NOT inside menu.
    //    -> menu.classList.add('hidden') -> Hides it immediately.
    
    // So I need to prevent this.
    // I can check if the click target is the button.
    const btn = e.target.closest('button');
    if (btn && btn.querySelector('.fa-sliders')) return;
    
    menu.classList.add('hidden');
  });
});
