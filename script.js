// Purpose: Main script for the inventory system with dynamic configuration.
// Description: This script handles the main functionality of the inventory system with real-time data fetching, centralized configuration, and comprehensive error handling.
// version: 3.0 - Dynamic Configuration Edition

// System Configuration (loaded dynamically)
let SYSTEM_CONFIG = {
    scriptId: 'AKfycbxoVWJXHOMAwulbuUoquiE-sSDkVyWKwkgprgOMVyvb7eLUpjnoN8G4YwX6O9YoNv1F',
    apiBaseUrl: 'https://script.google.com/macros/s/AKfycbxoVWJXHOMAwulbuUoquiE-sSDkVyWKwkgprgOMVyvb7eLUpjnoN8G4YwX6O9YoNv1F/exec',
    batchSize: 10,
    cacheTimeout: 300000,
    enableLogging: true,
    enableCaching: true,
    enableChat: false,
    enableReports: true,
    enableAudit: true,
    appName: 'Inventory Management System',
    appVersion: '1.0.0',
    companyName: 'NBP/GovNet',
    externalApis: {
        inventory: '',
        suppliers: '',
        reports: ''
    }
};

// Configuration Cache
let CONFIG_LOADED = false;
let CONFIG_LOAD_PROMISE = null;

// Caching mechanism
const inventoryCache = {
    items: [],
    filteredItems: [], // Added for client-side filtering
    columns: [],
    sheets: [],
    page: 1,
    pageCount: 1,
    hasMore: true,
    loaded: false,
    sheetsLoaded: false,
    columnsLoaded: false,
    categories: [],
    categoriesLoaded: false,
    searchQuery: '',
    filters: {
        stock: 'all',
        status: 'all',
        category: 'all'
    }
};

const dashboardCache = {
    stats: {},
    recentItems: [],
    loaded: false
};

const reportsCache = {
    items: [],
    filteredItems: [],
    currentPage: 1,
    itemsPerPage: 5,
    view: 'card',
    loaded: false,
    filters: {
        query: '',
        category: '',
        status: '',
        sortBy: 'newest'
    }
};

const globalSearchState = {
    query: '',
    timeoutId: null
};

// Debounce API calls
let debounceTimeout;

/**
 * LoadingManager - Unified Application Feedback System
 */
const LoadingManager = {
    _timer: null,
    _progressBar: null,
    _overlay: null,
    _syncToast: null,

    init() {
        this._progressBar = document.getElementById('nprogress-bar');
        this._overlay = document.getElementById('global-loader');
        this._syncToast = document.getElementById('sync-toast');
    },

    // 1. Global Page Transitions (Progress Bar at Top)
    startProgress() {
        if (!this._progressBar) return;
        this._progressBar.style.transform = 'translateX(-70%)';
        this._progressBar.style.opacity = '1';
    },

    finishProgress() {
        if (!this._progressBar) return;
        this._progressBar.style.transform = 'translateX(0%)';
        setTimeout(() => {
            this._progressBar.style.opacity = '0';
            setTimeout(() => {
                this._progressBar.style.transform = 'translateX(-100%)';
            }, 400);
        }, 300);
    },

    // 2. Full-Screen Blocking Loader (Used for critical initial loads)
    showOverlay(text = 'Fetching latest records...', subtext = 'GovNet Infrastructure Sync') {
        if (!this._overlay) return;
        const textEl = document.getElementById('loader-text');
        const subtextEl = document.getElementById('loader-subtext');
        const errorEl = document.getElementById('loader-error');
        
        if (textEl) textEl.textContent = text;
        if (subtextEl) subtextEl.textContent = subtext;
        if (errorEl) errorEl.classList.add('hidden');

        this._overlay.classList.remove('opacity-0', 'pointer-events-none');
        this._overlay.classList.add('opacity-100');

        // Set safety timeout
        clearTimeout(this._timer);
        this._timer = setTimeout(() => {
            if (this._overlay.classList.contains('opacity-100') && errorEl) {
                errorEl.classList.remove('hidden');
            }
        }, 15000);
    },

    hideOverlay() {
        if (!this._overlay) return;
        clearTimeout(this._timer);
        this._overlay.classList.add('opacity-0', 'pointer-events-none');
        this._overlay.classList.remove('opacity-100');
    },

    // 3. Micro-feedback Syncing Toast (Top Right)
    showSyncToast(text = 'Syncing with Google Sheets...') {
        if (!this._syncToast) return;
        const textEl = document.getElementById('sync-toast-text');
        if (textEl) textEl.textContent = text;
        this._syncToast.classList.remove('translate-x-full');
    },

    hideSyncToast() {
        if (!this._syncToast) return;
        this._syncToast.classList.add('translate-x-full');
    },

    // 4. Inline Button Loading
    setBtnLoading(btn, isLoading, originalHtml = '') {
        if (!btn) return;
        if (isLoading) {
            btn.disabled = true;
            btn.dataset.original = btn.innerHTML;
            btn.innerHTML = `<i class="fa-solid fa-circle-notch fa-spin mr-2"></i> ${originalHtml || 'Processing...'}`;
        } else {
            btn.disabled = false;
            btn.innerHTML = btn.dataset.original || originalHtml;
        }
    }
};

/**
 * Toast Notifications
 * Shows a modern, non-intrusive popup for success/error messages
 */
function showToast(message, type = 'success', duration = 3500) {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `
        pointer-events-auto flex items-center gap-3 px-5 py-4 rounded-2xl shadow-2xl border
        transform translate-y-12 opacity-0 transition-all duration-300 ease-out min-w-[320px] max-w-md
        ${type === 'success' ? 'bg-white border-emerald-100 text-emerald-900 shadow-emerald-900/10' : 'bg-white border-rose-100 text-rose-900 shadow-rose-900/10'}
    `;

    const icon = type === 'success' ? 'fa-circle-check text-emerald-500' : 'fa-circle-exclamation text-rose-500';
    
    toast.innerHTML = `
        <div class="flex-shrink-0 w-10 h-10 rounded-xl ${type === 'success' ? 'bg-emerald-50' : 'bg-rose-50'} flex items-center justify-center">
            <i class="fa-solid ${icon} text-lg"></i>
        </div>
        <div class="flex-1 min-w-0">
            <p class="text-sm font-semibold">${type === 'success' ? 'Success' : 'Attention'}</p>
            <p class="text-xs text-gray-600 truncate">${message}</p>
        </div>
        <button class="p-2 text-gray-300 hover:text-gray-500 transition-colors" onclick="this.parentElement.remove()">
            <i class="fa-solid fa-xmark text-sm"></i>
        </button>
    `;

    container.appendChild(toast);

    // Animate in
    setTimeout(() => {
        toast.classList.remove('translate-y-12', 'opacity-0');
    }, 10);

    // Auto remove
    setTimeout(() => {
        toast.classList.add('translate-y-12', 'opacity-0');
        setTimeout(() => toast.remove(), 300);
    }, duration);
}

/**
 * Centralized Skeleton Loader Manager
 * Handles both overlay-style and inline-style skeletons for different sections.
 */
function showSkeleton(section, show = true) {
    const tableBody = document.getElementById(`${section}-table-body`);
    const overlaySkeleton = document.getElementById(`${section}-skeleton`);
    
    // Start top progress bar when skeleton shows
    if (show) LoadingManager.startProgress();
    else LoadingManager.finishProgress();

    // 1. Dashboard skeleton handling
    if (section === 'dashboard') {
        const setSkeleton = (id) => {
            const el = document.getElementById(id);
            if (el) {
                if (show) {
                    el.innerHTML = '<div class="h-8 bg-gray-200/80 rounded animate-pulse w-24"></div>';
                    el.setAttribute('aria-busy', 'true');
                } else {
                    el.removeAttribute('aria-busy');
                }
            }
        };
        ['dash-total-items', 'dash-low-stock', 'dash-out-stock', 'dash-total-value'].forEach(setSkeleton);
        
        const recentItemsContainer = document.getElementById('recent-items-list');
        if (recentItemsContainer) {
            if (show) {
                recentItemsContainer.setAttribute('aria-busy', 'true');
                recentItemsContainer.innerHTML = Array(5).fill(0).map((_, i) => `
                    <div class="flex items-center justify-between p-4 border-b last:border-b-0 animate-pulse" style="animation-delay: ${i * 100}ms">
                        <div class="space-y-2 w-full max-w-[60%]">
                            <div class="h-4 bg-gray-200 rounded w-3/4"></div>
                            <div class="h-3 bg-gray-100 rounded w-1/2"></div>
                        </div>
                        <div class="text-right space-y-2 w-full max-w-[30%]">
                            <div class="h-5 bg-gray-200 rounded-full w-16 ml-auto"></div>
                            <div class="h-3 bg-gray-100 rounded w-12 ml-auto"></div>
                        </div>
                    </div>
                `).join('');
            } else {
                recentItemsContainer.removeAttribute('aria-busy');
            }
        }
    } 
    
    // 2. Inventory skeleton handling
    else if (section === 'inventory') {
        if (!tableBody && !overlaySkeleton) return;
        
        if (show) {
            // If we have an overlay skeleton DIV, use it
            if (overlaySkeleton) {
                if (tableBody) tableBody.classList.add('hidden');
                overlaySkeleton.classList.remove('hidden');
            } 
            // Otherwise, populate the table body with skeleton rows
            else if (tableBody) {
                const colCount = inventoryCache.columns.length || 7;
                tableBody.innerHTML = Array(10).fill(0).map((_, i) => `
                    <tr class="border-b animate-pulse" style="animation-delay: ${i * 50}ms">
                        ${Array(colCount).fill(0).map(() => `<td class="py-4 px-4"><div class="h-4 bg-gray-200 rounded w-full opacity-50"></div></td>`).join('')}
                        <td class="py-4 px-4 text-right"><div class="flex justify-end gap-2"><div class="h-8 w-8 bg-gray-100 rounded"></div><div class="h-8 w-8 bg-gray-100 rounded"></div></div></td>
                    </tr>
                `).join('');
            }
        } else {
            if (overlaySkeleton) overlaySkeleton.classList.add('hidden');
            if (tableBody) tableBody.classList.remove('hidden');
        }
    }

    // 3. Reports skeleton handling
    else if (section === 'reports') {
        const container = document.getElementById('report-cards-container');
        if (container) {
            if (show) {
                container.setAttribute('aria-busy', 'true');
                container.innerHTML = `<div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-6">
                    ${Array(8).fill(0).map((_, i) => `
                        <div class="bg-white rounded-2xl shadow-lg p-6 space-y-4 animate-pulse" style="animation-delay: ${i * 100}ms">
                            <div class="h-6 bg-gray-200 rounded w-3/4"></div>
                            <div class="h-4 bg-gray-100 rounded w-full"></div>
                        </div>
                    `).join('')}
                </div>`;
            } else {
                container.removeAttribute('aria-busy');
            }
        }
    }
}

// DOM Ready
document.addEventListener('DOMContentLoaded', () => {
    checkSecurityAccess();
});

function checkSecurityAccess() {
    const isAuthenticated = sessionStorage.getItem('auth_token') || localStorage.getItem('auth_token');
    const securityModal = document.getElementById('security-modal');
    
    if (isAuthenticated) {
        if (securityModal) securityModal.classList.add('hidden');
        loadSystemConfiguration().then(() => {
            initializePage();
        }).catch(error => {
            console.error('Failed to load system configuration:', error);
            // Fallback to initialize page without config
            initializePage();
        });
    } else {
        if (securityModal) {
            securityModal.classList.remove('hidden');
            const input = document.getElementById('security-input');
            if (input) input.focus();
        }
    }
}

// Configuration Management
async function loadSystemConfiguration() {
    if (CONFIG_LOADED && CONFIG_LOAD_PROMISE) {
        return CONFIG_LOAD_PROMISE;
    }
    
    if (CONFIG_LOAD_PROMISE) {
        return CONFIG_LOAD_PROMISE;
    }
    
    CONFIG_LOAD_PROMISE = fetchSystemConfiguration();
    
    try {
        const prev = SYSTEM_CONFIG;
        const config = await CONFIG_LOAD_PROMISE;
        SYSTEM_CONFIG = { ...prev, ...config };
        const hasVal = v => v !== undefined && v !== null && String(v).trim() !== '';
        SYSTEM_CONFIG.apiBaseUrl = hasVal(config.apiBaseUrl) ? config.apiBaseUrl : prev.apiBaseUrl;
        SYSTEM_CONFIG.scriptId = hasVal(config.scriptId) ? config.scriptId : prev.scriptId;
        CONFIG_LOADED = true;
        
        // Update UI with dynamic config
        updateUIWithConfiguration();
        
        return config;
    } catch (error) {
        CONFIG_LOAD_PROMISE = null;
        throw error;
    }
}

async function fetchSystemConfiguration() {
    // Try to get configuration from backend
    try {
        const response = await callApi('?action=getSystemConfig');
        if (response && !response.error) {
            return response;
        }
    } catch (error) {
        console.warn('Failed to fetch system configuration from backend:', error);
    }
    
    // Fallback: Try to get from localStorage/sessionStorage
    const storedConfig = localStorage.getItem('system_config');
    if (storedConfig) {
        try {
            return JSON.parse(storedConfig);
        } catch (error) {
            console.warn('Failed to parse stored configuration:', error);
        }
    }
    
    // Return default config
    return SYSTEM_CONFIG;
}

function updateUIWithConfiguration() {
    // Update app title
    document.title = SYSTEM_CONFIG.appName;
    
    // Update company name in sidebar
    const companyElements = document.querySelectorAll('.company-name');
    companyElements.forEach(el => {
        el.textContent = SYSTEM_CONFIG.companyName;
    });
    
    // Update app name
    const appNameElements = document.querySelectorAll('.app-name');
    appNameElements.forEach(el => {
        el.textContent = SYSTEM_CONFIG.appName;
    });
    
    // Update version
    const versionElements = document.querySelectorAll('.app-version');
    versionElements.forEach(el => {
        el.textContent = `v${SYSTEM_CONFIG.appVersion}`;
    });
    
    // Show/hide features based on feature flags
    if (!SYSTEM_CONFIG.enableChat) {
        const chatElements = document.querySelectorAll('.chat-feature');
        chatElements.forEach(el => el.style.display = 'none');
    }
    
    if (!SYSTEM_CONFIG.enableReports) {
        const reportElements = document.querySelectorAll('.reports-nav');
        reportElements.forEach(el => el.style.display = 'none');
    }
}

function initializePage() {
    setupSidebar();
    updateActiveNav();
    const currentSection = getSectionFromHash();
    // Pass false to prevent pushing the same state again on initial load
    loadSection(currentSection, false);
    setupGlobalSearch();
    startAutoRefresh();
}

async function handleSecurityCheck(event) {
    event.preventDefault();
    const securityInput = document.getElementById('security-input');
    const password = securityInput.value;

    if (!password) {
        showToast('Please enter the access script.', 'error');
        return;
    }

    try {
        const hashedPassword = await sha256(password);
        let response = null;
        try {
            response = await callApi(`?action=verifyAccess&hash=${encodeURIComponent(hashedPassword)}`);
        } catch (error) {
            response = { success: false, error: error.message };
        }

        const localExpectedHash = await sha256('1Vt_jqc3vo0Z_YMlkSTJVFGDNjB9efBC1075DVu0qbt9p_-0rZ1qfDNYC');
        const localMatch = hashedPassword === localExpectedHash;

        if (response && response.success) {
            localStorage.setItem('auth_token', hashedPassword);
            sessionStorage.setItem('auth_token', hashedPassword);
            document.getElementById('security-modal').classList.add('hidden');
            initializePage();
            return;
        }

        if (localMatch && response && (response.error || !response.success)) {
            localStorage.setItem('auth_token', hashedPassword);
            sessionStorage.setItem('auth_token', hashedPassword);
            document.getElementById('security-modal').classList.add('hidden');
            initializePage();
            return;
        }

        showToast(response && response.message ? response.message : 'Invalid access script.', 'error');
        securityInput.value = '';
    } catch (error) {
        showToast(`An error occurred: ${error.message}`, 'error');
    }
}

async function sha256(message) {
    const msgBuffer = new TextEncoder().encode(message);
    const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

function getSectionFromHash() {
    return window.location.hash.substring(1) || 'dashboard';
}

function setSectionInHash(sectionId) {
    // Only update if it's different to avoid redundant history entries
    if (window.location.hash.substring(1) !== sectionId) {
        if (history.pushState) {
            history.pushState(null, null, `#${sectionId}`);
        } else {
            window.location.hash = sectionId;
        }
    }
}

function setupGlobalSearch() {
    const searchInput = document.getElementById('global-search-input');
    if (searchInput) {
        searchInput.addEventListener('keyup', handleGlobalSearch);
    }
}

function handleGlobalSearch(event) {
    const query = event.target.value;
    clearTimeout(globalSearchState.timeoutId);

    globalSearchState.timeoutId = setTimeout(() => {
        globalSearchState.query = query;
        const currentSection = getSectionFromHash();

        switch (currentSection) {
            case 'dashboard':
                // Dashboard doesn't have a dedicated search view, but you might want to refresh data
                break;
            case 'inventory':
                inventoryCache.searchQuery = query;
                inventoryCache.page = 1;
                inventoryCache.items = [];
                inventoryCache.hasMore = true;
                loadInventory(true);
                break;
            case 'reports':
                renderReportsView(query);
                break;
        }
    }, 300); // 300ms debounce
}


// Navigation
function updateActiveNav() {
    const section = getSectionFromHash();
    const navLinks = document.querySelectorAll('.nav-item');
    navLinks.forEach(link => {
        if (link.id === `nav-${section}`) {
            link.classList.add('active', 'bg-white/10', 'text-white');
            link.classList.remove('text-gray-400');
        } else {
            link.classList.remove('active', 'bg-white/10', 'text-white');
            link.classList.add('text-gray-400');
        }
    });
}

function setupSidebar() {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar) return;

    const collapsed = localStorage.getItem('sidebar_collapsed') === '1';
    setSidebarCollapsed(collapsed, false);

    const collapseBtn = document.getElementById('sidebar-collapse-btn');
    if (collapseBtn) {
        collapseBtn.onclick = () => setSidebarCollapsed(!sidebar.classList.contains('collapsed'), true);
    }

    const mobileBtn = document.getElementById('mobile-menu-btn');
    if (mobileBtn) {
        mobileBtn.onclick = () => {
            if (sidebar.classList.contains('-translate-x-full')) {
                openMobileSidebar();
            } else {
                closeMobileSidebar();
            }
        };
    }

    window.addEventListener('resize', () => {
        if (window.innerWidth >= 1024) closeMobileSidebar();
    });
}

function setSidebarCollapsed(collapsed, persist) {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar) return;
    if (collapsed) {
        sidebar.classList.add('collapsed');
    } else {
        sidebar.classList.remove('collapsed');
    }

    const icon = document.getElementById('sidebar-collapse-icon');
    if (icon) {
        if (collapsed) {
            icon.classList.remove('fa-chevron-left');
            icon.classList.add('fa-chevron-right');
        } else {
            icon.classList.remove('fa-chevron-right');
            icon.classList.add('fa-chevron-left');
        }
    }

    if (persist) localStorage.setItem('sidebar_collapsed', collapsed ? '1' : '0');
}

function handleLogout() {
    sessionStorage.removeItem('auth_token');
    localStorage.removeItem('auth_token');
    window.location.reload();
}

function openMobileSidebar() {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar) return;
    sidebar.classList.remove('-translate-x-full');
}

function closeMobileSidebar() {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar) return;
    if (window.innerWidth < 1024) sidebar.classList.add('-translate-x-full');
}

// Section Loading
async function loadSection(sectionId, pushState = true) {
    if (pushState) {
        setSectionInHash(sectionId);
    }
    updateActiveNav();
    closeMobileSidebar();

    const mainContent = document.getElementById('main-content');
    mainContent.innerHTML = '<div class="flex justify-center items-center h-full"><div class="loader"></div></div>'; // Loading spinner

    // Update Header Title and Subtitle
    const pageTitle = document.getElementById('page-title');
    const pageSubtitle = pageTitle ? pageTitle.nextElementSibling : null;
    
    if (pageTitle) {
        switch (sectionId) {
            case 'dashboard':
                pageTitle.textContent = 'Dashboard Overview';
                if (pageSubtitle) pageSubtitle.textContent = "Welcome back! Here's your inventory summary";
                break;
            case 'inventory':
                pageTitle.textContent = 'Inventory Records';
                if (pageSubtitle) pageSubtitle.textContent = 'View and track all equipment and supplies';
                break;
            case 'reports':
                pageTitle.textContent = 'Reports & Forms';
                if (pageSubtitle) pageSubtitle.textContent = 'Generate and manage property documents';
                break;
            case 'par':
                pageTitle.textContent = 'Property Acknowledgement Receipt';
                if (pageSubtitle) pageSubtitle.textContent = 'Create and print PAR documents';
                break;
            default:
                pageTitle.textContent = sectionId.charAt(0).toUpperCase() + sectionId.slice(1);
                if (pageSubtitle) pageSubtitle.textContent = '';
        }
    }

    try {
        const response = await fetch(`${sectionId}.html`);
        if (!response.ok) throw new Error(`Page not found: ${sectionId}.html`);
        const sectionHtml = await response.text();
        mainContent.innerHTML = sectionHtml;

        // Post-load actions
        switch (sectionId) {
            case 'dashboard':
                if (dashboardCache.loaded) {
                    renderDashboard(dashboardCache.stats, dashboardCache.recentItems);
                } else {
                    await loadDashboard();
                }
                break;
            case 'inventory':
                if (inventoryCache.loaded && !inventoryCache.searchQuery) {
                    renderInventoryItems();
                } else {
                    await loadInventory(true);
                }
                break;
            case 'reports':
                if (reportsCache.loaded) {
                    populateReportFilterCategories();
                    applyReportsFilters();
                    renderKPIWidgets();
                } else {
                    await loadReports();
                }
                break;
            case 'par':
                initializeParPage();
                break;
        }
    } catch (error) {
        mainContent.innerHTML = `<div class="text-center text-red-500 p-8">Error loading section: ${error.message}</div>`;
        console.error(`Error in loadSection for ${sectionId}:`, error);
    }
}


// API Call Abstraction
async function callApi(endpoint, options = {}) {
    const config = SYSTEM_CONFIG;
    
    // Ensure configuration is loaded
    if (!CONFIG_LOADED && !endpoint.includes('getSystemConfig')) {
        await loadSystemConfiguration();
    }
    
    const urlParams = new URLSearchParams(window.location.search);
    const useRealData = urlParams.get('use_real_data') === 'true';
    const isLocalhost = ['localhost', '127.0.0.1'].includes(window.location.hostname);
    
    // Log API call if logging is enabled
    if (config.enableLogging) {
        console.log(`[API] ${endpoint} - Localhost: ${isLocalhost}, Use Real Data: ${useRealData}`);
    }
    
    if (isLocalhost && !useRealData && 1 === 2) { // disabled for now
        console.warn('Running on localhost: Using MOCK DATA for ' + endpoint);
        return getMockData(endpoint);
    }

    // Use dynamic script ID from configuration
    const scriptId = config.scriptId || 'AKfycbx90Go7FnXhMOP5-FWblg_usGbEv4ZMMbrHcbeYc_B-h98Ljk-YLNbAZB6pP8ybZy3l';
    const baseUrl = config.apiBaseUrl || `https://script.google.com/macros/s/${scriptId}/exec`;
    const url = `${baseUrl}${endpoint}`;
    
    // For POST requests to Google Apps Script Web App, we need to use 'text/plain' or 'application/x-www-form-urlencoded'
    // to avoid CORS preflight (OPTIONS) requests which are often not handled correctly by GAS.
    // However, if we use 'text/plain', the body is just text. We need to handle this in backend or use a specific pattern.
    // The standard workaround for GAS CORS is to use 'application/x-www-form-urlencoded' or just rely on 'text/plain' and parse manually.
    // BUT, since we are sending JSON, let's try to be consistent.
    // If method is POST, we ensure headers are set correctly for GAS.
    
    const fetchOptions = { ...options };
    
    if (fetchOptions.method === 'POST') {
        // Force Content-Type to text/plain to avoid CORS preflight issues with Google Apps Script
        // The backend `doPost` should parse the postData.contents regardless of Content-Type header if it expects JSON.
        fetchOptions.headers = {
            ...fetchOptions.headers,
            'Content-Type': 'text/plain;charset=utf-8',
        };
    }

    try {
        const response = await fetch(url, fetchOptions);
        const responseText = await response.text();
        
        if (!response.ok) {
            let errorData = { message: `HTTP error! status: ${response.status}` };
            try {
                errorData = JSON.parse(responseText);
            } catch (e) { /* Ignore if response is not JSON */ }
            throw new Error(errorData.message || `HTTP error! status: ${response.status}`);
        }
        
        const json = JSON.parse(responseText);
        
        // Log the raw response for debugging
        console.log('[API Response]', endpoint, json);
        
        if (json && typeof json === 'object' && 'error' in json && !json.success) {
            throw new Error(String(json.error || 'Request failed'));
        }
        return json;
    } catch (error) {
        console.error(`API call to ${endpoint} failed:`, error);
        
        // Log error if logging is enabled
        if (config.enableLogging) {
            console.error(`[API Error] ${endpoint}: ${error.message}`);
        }
        
        throw error;
    }
}

function getMockData(endpoint) {
    const config = SYSTEM_CONFIG;
    
    // Check for simulation scenarios from URL
    const urlParams = new URLSearchParams(window.location.search);
    const scenario = urlParams.get('scenario');
    
    const delay = scenario === 'slow' ? 3000 : 800;
    
    return new Promise((resolve, reject) => {
        setTimeout(() => {
            // Simulate network failure
            if (scenario === 'error' || scenario === 'network-error') {
                reject(new Error('Simulated network error'));
                return;
            }
            
            // Simulate empty data
            if (scenario === 'empty') {
                if (endpoint.includes('action=getDashboardStats')) {
                    resolve({ totalItems: 0, lowStock: 0, outOfStock: 0, totalValue: 0, weeklyActivity: { labels: [], sales: [], restocks: [], trend: [] } });
                } else if (endpoint.includes('action=getItems')) {
                    resolve({ items: [], hasMore: false });
                } else if (endpoint.includes('action=getInventory')) {
                    resolve([]);
                } else {
                    resolve({});
                }
                return;
            }

            // Dynamic mock data based on configuration
            if (endpoint.includes('action=getDashboardStats')) {
                resolve({
                    totalItems: Math.floor(Math.random() * 200) + 50,
                    lowStock: Math.floor(Math.random() * 30) + 5,
                    outOfStock: Math.floor(Math.random() * 10) + 1,
                    totalValue: Math.floor(Math.random() * 1000000) + 100000,
                    weeklyActivity: {
                        labels: ["Wed","Thu","Fri","Sat","Sun","Mon","Tue"],
                        sales: Array.from({length: 7}, () => Math.floor(Math.random() * 15)),
                        restocks: Array.from({length: 7}, () => Math.floor(Math.random() * 20)),
                        trend: Array.from({length: 7}, () => Math.random() * 2)
                    }
                });
            } else if (endpoint.includes('action=getItems')) {
                // Simulate large dataset if requested
                const itemCount = scenario === 'large' ? 100 : (config.batchSize || 20);
                
                const items = Array.from({length: itemCount}, (_, i) => ({
                    ID: `${100 + i}`,
                    Item: `Dynamic Item ${i + 1}`,
                    Category: ['Electronics', 'Furniture', 'Supplies', 'Equipment'][Math.floor(Math.random() * 4)],
                    Qty: Math.floor(Math.random() * 50) + 1,
                    UnitCost: Math.floor(Math.random() * 5000) + 100,
                    Status: Math.random() > 0.7 ? 'Low Stock' : (Math.random() > 0.9 ? 'Out of Stock' : 'In Stock'),
                    DateAcquired: new Date(Date.now() - Math.random() * 365 * 24 * 60 * 60 * 1000).toISOString().split('T')[0],
                    Remarks: `Mock data for ${config.appName}`
                }));
                resolve({
                    items: items,
                    hasMore: scenario === 'large' ? true : Math.random() > 0.5
                });
            } else if (endpoint.includes('action=getInventory')) {
                const items = Array.from({length: 10}, (_, i) => ({
                    ID: `${200 + i}`,
                    Item: `External Item ${i + 1}`,
                    Category: ['Electronics', 'Furniture', 'Supplies'][Math.floor(Math.random() * 3)],
                    Qty: Math.floor(Math.random() * 30) + 1,
                    UnitCost: Math.floor(Math.random() * 2000) + 50,
                    Status: Math.random() > 0.8 ? 'Low Stock' : 'In Stock',
                    DateAcquired: new Date(Date.now() - Math.random() * 180 * 24 * 60 * 60 * 1000).toISOString().split('T')[0],
                    Remarks: `External data for ${config.companyName}`
                }));
                resolve(items);
            } else if (endpoint.includes('action=getSystemConfig')) {
                resolve(config);
            } else {
                resolve({ success: true, message: 'Dynamic mock success', config: config });
            }
        }, delay); // Simulate network delay
    });
}

/**
 * Reusable Data Normalization Functions
 * Handles robust parsing of text, numbers, dates, and arrays from API responses.
 */

function parseNumber(value, defaultValue = 0) {
    if (value === null || value === undefined) return defaultValue;
    if (typeof value === 'number') return value;
    const str = String(value).replace(/[^0-9.\-]/g, ''); // Remove currency symbols, commas, etc.
    const num = parseFloat(str);
    return Number.isFinite(num) ? num : defaultValue;
}

function parseDate(value) {
    if (!value) return '';
    
    // If it's already a Date object
    if (value instanceof Date) {
        if (isNaN(value.getTime())) return '';
        return value.toISOString();
    }

    // Handle 4-digit numbers (years)
    if (typeof value === 'number' && value >= 1900 && value <= 2100) {
        return new Date(value, 0, 1).toISOString();
    }

    // Handle Excel Serial Dates (numbers like 44561)
    if (typeof value === 'number') {
        if (value > 25569 && value < 60000) { // Typical range for 1970-2060
            const ms = (value - 25569) * 86400 * 1000;
            const d = new Date(ms);
            return isNaN(d.getTime()) ? '' : d.toISOString();
        }
        // If it's a timestamp
        if (value > 1000000000) {
            const d = new Date(value);
            return isNaN(d.getTime()) ? '' : d.toISOString();
        }
    }

    // Handle String Dates
    if (typeof value === 'string') {
        const s = value.trim();
        if (!s) return '';

        // If it's a 4-digit year string
        if (/^\d{4}$/.test(s)) {
            const year = parseInt(s);
            if (year >= 1900 && year <= 2100) {
                return new Date(year, 0, 1).toISOString();
            }
        }

        // Try standard parsing first
        const d = new Date(s);
        if (!isNaN(d.getTime())) return d.toISOString();

        // Handle DD/MM/YYYY or MM/DD/YYYY
        const parts = s.split(/[\/\-]/);
        if (parts.length === 3) {
            if (parts[0].length === 4) { // YYYY-MM-DD
                const d2 = new Date(parts[0], parts[1] - 1, parts[2]);
                if (!isNaN(d2.getTime())) return d2.toISOString();
            } else if (parts[2].length === 4) { // DD/MM/YYYY
                const d2 = new Date(parts[2], parts[1] - 1, parts[0]);
                if (!isNaN(d2.getTime())) return d2.toISOString();
            }
        }
    }

    return String(value || ''); // Fallback to raw string if parsing fails
}

function normalizeItemForUi(rawItem) {
    if (!rawItem) return {};
    
    // Handle quantity with robust parsing
    const qty = parseNumber(rawItem.Qty !== undefined ? rawItem.Qty : rawItem.Quantity);
    
    // Handle status logic
    const rawStatus = String(rawItem.Status || '').trim();
    const status = rawStatus ? rawStatus : (qty <= 0 ? 'Out of Stock' : 'In Stock');

    return {
        ...rawItem, // Preserve all raw fields
        ID: String(rawItem.ID || '').trim(),
        ItemName: String(rawItem.Item || rawItem.ItemName || '').trim(),
        Category: String(rawItem.Category || '').trim(),
        Status: status,
        Quantity: qty,
        UnitCost: parseNumber(rawItem.UnitCost || rawItem.Cost),
        DateAdded: parseDate(rawItem.DateAcquired || rawItem.DateAdded),
        Description: String(rawItem.Remarks || rawItem.BrandModel || rawItem.Description || '').trim()
    };
}

function formatDate(value) {
    if (!value) return '-';
    
    // If value is a 4-digit number or string that looks like a year
    const strVal = String(value).trim();
    if (/^\d{4}$/.test(strVal)) {
        const year = parseInt(strVal);
        if (year >= 1900 && year <= 2100) return strVal;
    }

    const d = new Date(value);
    if (isNaN(d.getTime())) return value; // Return original string if parse fails

    // If the date is Jan 1st and it's from a year-only input, we might want to just show the year
    if (d.getMonth() === 0 && d.getDate() === 1 && d.getHours() === 0 && d.getMinutes() === 0) {
        // If it was parsed from a simple year string or number
        if (strVal.length === 4 || (!isNaN(value) && value < 3000)) {
            return String(d.getFullYear());
        }
    }

    return d.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
    });
}
function toYmd(value) {
    const iso = parseDate(value);
    return iso ? iso.split('T')[0] : '';
}

function formatCurrency(value) {
    const num = parseNumber(value);
    return new Intl.NumberFormat('en-PH', {
        style: 'currency',
        currency: 'PHP'
    }).format(num);
}

// Dashboard
async function loadDashboard(forceReload = false) {
    if (dashboardCache.loaded && !forceReload) {
        console.log('[Dashboard] Using cached frontend data');
        renderDashboard(dashboardCache.stats, dashboardCache.recentItems);
        return;
    }
    
    showSkeleton('dashboard');
    
    try {
        const [stats, recentItemsPayload] = await Promise.all([
            callApi('?action=getDashboardStats'),
            callApi('?action=getItems&limit=5&page=1&order=desc')
        ]);
        const recentItems = Array.isArray(recentItemsPayload && recentItemsPayload.items)
            ? recentItemsPayload.items.map(normalizeItemForUi)
            : [];
        dashboardCache.stats = stats;
        dashboardCache.recentItems = recentItems;
        dashboardCache.loaded = true;
        renderDashboard(stats, recentItems);
    } catch (error) {
        console.error('Error loading dashboard:', error);
        // Show error state in dashboard elements
        const errorHtml = '<span class="text-red-500 text-sm">Error</span>';
        ['dash-total-items', 'dash-low-stock', 'dash-out-stock', 'dash-total-value'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.innerHTML = errorHtml;
        });
    }
}

function renderDashboard(stats, recentItems) {
    console.log("renderDashboard stats:", stats);

    const totalItems = Number(stats?.totalItems ?? stats?.total_items ?? 0);
    const outOfStock = Number(stats?.outOfStock ?? stats?.out_of_stock ?? 0);
    const lowStock = Number(stats?.lowStock ?? stats?.low_stock ?? 0);
    const totalValue = Number(stats?.totalValue ?? stats?.total_value ?? 0).toLocaleString('en-PH', { style: 'currency', currency: 'PHP' });
    
    const setElementText = (id, text) => {
        const el = document.getElementById(id);
        if (el) el.textContent = text;
    };

    setElementText('dash-total-items', totalItems);
    setElementText('dash-low-stock', lowStock);
    setElementText('dash-out-stock', outOfStock);
    setElementText('dash-total-value', totalValue);
    
    const timeStr = new Date().toLocaleTimeString();
    setElementText('dash-last-sync', timeStr);
    setElementText('dash-last-sync-v2', timeStr);

    // Pulse animation for Out of Stock
    const outStockCard = document.getElementById('out-stock-card');
    if (outStockCard) {
        if (outOfStock > 0) {
            outStockCard.classList.add('ring-2', 'ring-rose-500', 'animate-pulse');
        } else {
            outStockCard.classList.remove('ring-2', 'ring-rose-500', 'animate-pulse');
        }
    }

    if (typeof Chart !== 'undefined') {
        updateDashboardCharts(stats);
    }

    // Activity Log / Recent Items
    const activityContainer = document.getElementById('recent-activities-list');
    if (activityContainer) {
        if (!recentItems || recentItems.length === 0) {
            activityContainer.innerHTML = '<div class="flex justify-center items-center h-full py-20 text-gray-400 text-sm italic">No recent activity logs found.</div>';
        } else {
            activityContainer.innerHTML = recentItems.map(item => `
                <div class="group flex items-start gap-4 p-4 rounded-2xl hover:bg-slate-50 transition-all border border-transparent hover:border-gray-100">
                    <div class="h-10 w-10 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center shrink-0">
                        <i class="fa-solid fa-box-open text-sm"></i>
                    </div>
                    <div class="min-w-0 flex-1">
                        <p class="text-sm font-bold text-gray-900 truncate">${item.ItemName}</p>
                        <p class="text-xs text-gray-500 mt-0.5 uppercase tracking-wider font-semibold">${item.Category}</p>
                        <div class="flex items-center gap-2 mt-2">
                            <span class="px-2 py-0.5 text-[10px] rounded-md font-bold border ${getStatusColorClass(item.Status)}">${item.Status}</span>
                            <span class="text-[10px] text-gray-400 font-medium">${formatDate(item.DateAdded)}</span>
                        </div>
                    </div>
                </div>
            `).join('');
        }
    }
}

let DASHBOARD_CHART_TIMEFRAME = 7;

function updateDashChartTimeframe(days) {
    DASHBOARD_CHART_TIMEFRAME = days;
    document.querySelectorAll('.dash-time-btn').forEach(btn => {
        const isActive = btn.textContent.includes(days === 365 ? 'YTD' : days + 'D');
        if (isActive) {
            btn.classList.add('active', 'bg-white', 'text-indigo-600', 'shadow-sm');
            btn.classList.remove('text-gray-500');
        } else {
            btn.classList.remove('active', 'bg-white', 'text-indigo-600', 'shadow-sm');
            btn.classList.add('text-gray-500');
        }
    });
    
    // In a real app, you'd fetch data here. For now, we'll just re-render with existing data
    if (dashboardCache.stats) updateDashboardCharts(dashboardCache.stats);
}

function updateDashboardCharts(stats) {
    const recreateChart = (canvasId, config) => {
        const canvas = document.getElementById(canvasId);
        if (!canvas) return;
        const existingChart = Chart.getChart(canvas);
        if (existingChart) existingChart.destroy();
        return new Chart(canvas, config);
    };

    // Stock Status Donut with Center Label
    const totalItems = stats?.totalItems || 0;
    recreateChart('stockChart', {
        type: 'doughnut',
        data: {
            labels: ['In Stock', 'Low Stock', 'Out of Stock'],
            datasets: [{
                data: [
                    Math.max(0, (stats?.totalItems || 0) - (stats?.outOfStock || 0) - (stats?.lowStock || 0)),
                    stats?.lowStock || 0,
                    stats?.outOfStock || 0
                ],
                backgroundColor: ['#10b981', '#f59e0b', '#f43f5e'],
                hoverOffset: 10,
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '80%',
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: (ctx) => ` ${ctx.label}: ${ctx.raw} units`
                    }
                }
            },
            // Custom plugin for center text
            onAfterRender: (chart) => {
                const { ctx, width, height } = chart;
                ctx.restore();
                ctx.font = "bold 24px sans-serif";
                ctx.textBaseline = "middle";
                ctx.textAlign = "center";
                ctx.fillStyle = "#1e293b";
                ctx.fillText(totalItems, width / 2, height / 2 - 5);
                ctx.font = "bold 10px sans-serif";
                ctx.fillStyle = "#94a3b8";
                ctx.fillText("TOTAL ASSETS", width / 2, height / 2 + 15);
                ctx.save();
            }
        }
    });

    // Main Movement Chart with Crosshairs/Custom Tooltips
    const weeklyData = stats?.weeklyActivity || { labels: [], sales: [], restocks: [] };
    recreateChart('mainChart', {
        type: 'line',
        data: {
            labels: weeklyData.labels.length ? weeklyData.labels : ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
            datasets: [
                {
                    label: 'Stock In',
                    data: weeklyData.restocks.length ? weeklyData.restocks : [5, 12, 8, 15, 7, 10, 12],
                    borderColor: '#3b82f6',
                    backgroundColor: 'rgba(59, 130, 246, 0.1)',
                    fill: true,
                    tension: 0.4,
                    borderWidth: 3,
                    pointRadius: 4,
                    pointHoverRadius: 6
                },
                {
                    label: 'Stock Out',
                    data: weeklyData.sales.length ? weeklyData.sales : [3, 8, 12, 5, 9, 4, 15],
                    borderColor: '#ef4444',
                    backgroundColor: 'rgba(239, 68, 68, 0.1)',
                    fill: true,
                    tension: 0.4,
                    borderWidth: 3,
                    pointRadius: 4,
                    pointHoverRadius: 6
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                intersect: false,
                mode: 'index'
            },
            scales: {
                y: { 
                    beginAtZero: true, 
                    grid: { color: '#f1f5f9' },
                    ticks: { font: { size: 10, weight: 'bold' }, color: '#94a3b8' }
                },
                x: { 
                    grid: { display: false },
                    ticks: { font: { size: 10, weight: 'bold' }, color: '#94a3b8' }
                }
            },
            plugins: {
                legend: {
                    position: 'top',
                    align: 'end',
                    labels: { usePointStyle: true, pointStyle: 'circle', boxWidth: 6, font: { size: 11, weight: 'bold' } }
                },
                tooltip: {
                    padding: 12,
                    backgroundColor: '#0f172a',
                    titleFont: { size: 13, weight: 'bold' },
                    bodyFont: { size: 12 },
                    usePointStyle: true,
                    callbacks: {
                        label: (ctx) => ` ${ctx.dataset.label}: ${ctx.raw} units`
                    }
                }
            }
        }
    });
}

function filterDashboardItems(type) {
    const modal = document.getElementById('dashboardQuickViewModal');
    const title = document.getElementById('quickview-title');
    const content = document.getElementById('quickview-content');
    if (!modal || !content) return;

    modal.classList.remove('hidden');
    content.innerHTML = '<div class="col-span-full py-10 flex flex-col items-center gap-3"><i class="fa-solid fa-circle-notch fa-spin text-2xl text-blue-500"></i><p class="text-sm text-gray-500 font-bold uppercase tracking-widest">Fetching Data...</p></div>';

    // Simulated filtering from inventoryCache
    const items = inventoryCache.items.filter(item => {
        const qty = parseNumber(item.Quantity || item.Qty);
        if (type === 'low') return qty > 0 && qty <= 10;
        if (type === 'out') return qty <= 0;
        return true;
    });

    title.textContent = type === 'low' ? 'Low Stock Assets' : 'Out of Stock Assets';
    
    setTimeout(() => {
        if (items.length === 0) {
            content.innerHTML = '<div class="col-span-full py-10 text-center text-gray-400 italic">No items found matching this criteria.</div>';
        } else {
            content.innerHTML = items.map(item => `
                <div class="p-4 rounded-2xl bg-white border border-gray-100 shadow-sm flex items-center gap-4">
                    <div class="h-12 w-12 rounded-xl bg-slate-50 flex items-center justify-center shrink-0">
                        <i class="fa-solid fa-box text-slate-400"></i>
                    </div>
                    <div class="min-w-0 flex-1">
                        <p class="text-sm font-bold text-gray-900 truncate">${item.ItemName || item.Item}</p>
                        <p class="text-xs text-rose-600 font-black tabular-nums mt-0.5">STOCK: ${item.Quantity || item.Qty}</p>
                    </div>
                    <button onclick="openEditModal('${item.ID}')" class="p-2 hover:bg-slate-100 rounded-lg transition-colors">
                        <i class="fa-solid fa-pen-to-square text-blue-500"></i>
                    </button>
                </div>
            `).join('');
        }
    }, 400);
}

function closeDashboardQuickView() {
    const modal = document.getElementById('dashboardQuickViewModal');
    if (modal) modal.classList.add('hidden');
}

// Polling for data refresh
function startAutoRefresh() {
    setInterval(async () => {
        if (document.hidden) return;
        try {
            const currentSection = getSectionFromHash();
            if (currentSection === 'dashboard') {
                const [stats, recentItemsPayload] = await Promise.all([
                    callApi('?action=getDashboardStats'),
                    callApi('?action=getItems&limit=5&page=1&order=desc')
                ]);
                const recentItems = Array.isArray(recentItemsPayload && recentItemsPayload.items)
                    ? recentItemsPayload.items.map(normalizeItemForUi)
                    : [];
                if (JSON.stringify(stats) !== JSON.stringify(dashboardCache.stats) || JSON.stringify(recentItems) !== JSON.stringify(dashboardCache.recentItems)) {
                    dashboardCache.stats = stats;
                    dashboardCache.recentItems = recentItems;
                    renderDashboard(stats, recentItems);
                }
            }
        } catch (error) {
            console.error('Periodic data refresh failed:', error);
        }
    }, 30000);
}

// Skeleton Loading
async function refreshAllData() {
    const icon = document.getElementById('refresh-icon');
    if (icon) icon.classList.add('fa-spin');

    const currentSection = getSectionFromHash();
    
    // Show skeleton before loading
    showSkeleton(currentSection);

    try {
        switch (currentSection) {
            case 'dashboard':
                await loadDashboard(true);
                break;
            case 'inventory':
                await loadInventory(true);
                break;
            case 'reports':
                await loadReports(true);
                break;
        }
    } catch (error) {
        console.error('Manual refresh failed:', error);
        showInventoryAlert('Failed to refresh data.', 'Error');
    } finally {
        if (icon) setTimeout(() => icon.classList.remove('fa-spin'), 500);
    }
}


/**
 * Global State for Table Density and Selection
 */
let currentTableDensity = localStorage.getItem('inventory_density') || 'compact';

/**
 * Toggle Table Density (Compact vs Relaxed)
 */
function setTableDensity(density) {
    currentTableDensity = density;
    localStorage.setItem('inventory_density', density);
    
    // Update UI Buttons
    const compactBtn = document.getElementById('density-compact');
    const relaxedBtn = document.getElementById('density-relaxed');
    
    if (compactBtn && relaxedBtn) {
        if (density === 'compact') {
            compactBtn.classList.add('bg-white', 'shadow-sm', 'text-indigo-600');
            compactBtn.classList.remove('text-gray-500');
            relaxedBtn.classList.remove('bg-white', 'shadow-sm', 'text-indigo-600');
            relaxedBtn.classList.add('text-gray-500');
        } else {
            relaxedBtn.classList.add('bg-white', 'shadow-sm', 'text-indigo-600');
            relaxedBtn.classList.remove('text-gray-500');
            compactBtn.classList.remove('bg-white', 'shadow-sm', 'text-indigo-600');
            compactBtn.classList.add('text-gray-500');
        }
    }
    
    renderInventoryItems();
}

/**
 * Update Dropdowns (Category & Status) from Data
 */
function updateInventoryFiltersFromData(items) {
    const catDropdown = document.getElementById('inv-filter-category');
    const statusDropdown = document.getElementById('inv-filter-status');
    if (!catDropdown && !statusDropdown) return;

    // Categories
    if (catDropdown && (!inventoryCache.categoriesLoaded || inventoryCache.categories.length === 0)) {
        const categories = [...new Set(items.map(i => i.Category || i.category).filter(Boolean))];
        if (categories.length > 0) {
            inventoryCache.categories = categories;
            inventoryCache.categoriesLoaded = true;
            const currentVal = catDropdown.value;
        catDropdown.innerHTML = '<option value="all" class="text-gray-800">All Categories</option>' + 
            categories.sort().map(c => `<option value="${c}" ${c === currentVal ? 'selected' : ''} class="text-gray-800">${c}</option>`).join('');
        }
    }

    // Statuses
    if (statusDropdown && (!inventoryCache.statusesLoaded || inventoryCache.statuses.length === 0)) {
        const statuses = [...new Set(items.map(i => i.Status || i.status).filter(Boolean))];
        if (statuses.length > 0) {
            inventoryCache.statuses = statuses;
            inventoryCache.statusesLoaded = true;
            const currentVal = statusDropdown.value;
            statusDropdown.innerHTML = '<option value="all" class="text-gray-800">All Status</option>' + 
                statuses.sort().map(s => `<option value="${s}" ${s === currentVal ? 'selected' : ''} class="text-gray-800">${s}</option>`).join('');
        }
    }
}

/**
 * Enhanced Status Badges with Color-coding & Health Indicators
 */
function getStatusBadgeHtml(status, qty) {
    status = String(status || '').toLowerCase();
    const q = parseInt(qty || 0);
    
    let colorClass = 'bg-blue-50 text-blue-700 border-blue-100';
    let dotClass = 'bg-blue-400';
    
    if (status.includes('out of stock') || q <= 0) {
        colorClass = 'bg-rose-50 text-rose-700 border-rose-100';
        dotClass = 'bg-rose-500 animate-pulse';
    } else if (status.includes('low stock') || (q > 0 && q <= 10)) {
        colorClass = 'bg-amber-50 text-amber-700 border-amber-100';
        dotClass = 'bg-amber-500';
    } else if (status.includes('in stock') || status.includes('available')) {
        colorClass = 'bg-emerald-50 text-emerald-700 border-emerald-100';
        dotClass = 'bg-emerald-500';
    }
    
    return `
        <div class="flex items-center gap-2 px-2.5 py-1 rounded-full border ${colorClass} w-fit shadow-sm">
            <span class="relative flex h-2 w-2">
                <span class="animate-ping absolute inline-flex h-full w-full rounded-full ${dotClass} opacity-20"></span>
                <span class="relative inline-flex rounded-full h-2 w-2 ${dotClass}"></span>
            </span>
            <span class="text-[10px] font-black uppercase tracking-wider whitespace-nowrap">${status || 'Unknown'}</span>
        </div>
    `;
}

/**
 * Progress Bar for Stock Levels
 */
function getStockProgressHtml(qty) {
    const q = parseInt(qty || 0);
    const max = 100; // Assume 100 as base for visualization
    const percentage = Math.min(100, (q / max) * 100);
    
    let barColor = 'bg-indigo-500';
    if (q <= 0) barColor = 'bg-rose-500';
    else if (q <= 10) barColor = 'bg-amber-500';
    else if (q > 50) barColor = 'bg-emerald-500';

    return `
        <div class="flex flex-col gap-1 min-w-[100px]">
            <div class="flex items-center text-[10px] font-bold">
                <span class="text-gray-900 tabular-nums">${q}</span>
            </div>
        </div>
    `;
}

/**
 * Updated handleInventorySearch to include Active Tags
 */
function updateActiveFilterTags(filters) {
    const container = document.getElementById('active-filters-container');
    const noFilters = document.getElementById('no-filters-tag');
    if (!container) return;

    // Remove existing tags except prefix
    const tags = container.querySelectorAll('.filter-tag');
    tags.forEach(t => t.remove());

    let hasActive = false;
    
    Object.entries(filters).forEach(([key, value]) => {
        if (value && value !== 'all') {
            hasActive = true;
            const tag = document.createElement('div');
            tag.className = 'filter-tag flex items-center gap-1.5 px-2 py-1 bg-indigo-50 text-indigo-700 rounded-lg border border-indigo-100 text-[10px] font-bold animate-in zoom-in-95 duration-200';
            tag.innerHTML = `
                <span class="opacity-60 uppercase">${key}:</span>
                <span>${value}</span>
                <button onclick="clearFilter('${key}')" class="hover:text-indigo-900 transition-colors">
                    <i class="fa-solid fa-circle-xmark"></i>
                </button>
            `;
            container.appendChild(tag);
        }
    });

    if (noFilters) {
        if (hasActive) noFilters.classList.add('hidden');
        else noFilters.classList.remove('hidden');
    }
}

function clearFilter(key) {
    const elId = `inv-filter-${key === 'query' ? 'search' : key}`;
    const el = document.getElementById(elId) || document.getElementById(`inv-search-input`);
    if (el) {
        el.value = (key === 'query' ? '' : 'all');
        handleInventorySearch({ target: el });
    }
}

/**
 * Enhanced Status Badges with Color-coding
 */
function getStatusColorClass(status) {
    status = String(status || '').toLowerCase();
    if (status.includes('out of stock') || status.includes('disposal') || status.includes('damaged') || status.includes('missing')) 
        return 'bg-rose-50 text-rose-700 border-rose-100';
    if (status.includes('low stock') || status.includes('warning') || status.includes('repair')) 
        return 'bg-amber-50 text-amber-700 border-amber-100';
    if (status.includes('in stock') || status.includes('available') || status.includes('good')) 
        return 'bg-emerald-50 text-emerald-700 border-emerald-100';
    return 'bg-blue-50 text-blue-700 border-blue-100';
}

/**
 * Handle Search & Filters - Real-time client-side filtering
 */
function handleInventorySearch(event) {
    const searchInput = document.getElementById('inv-search-input');
    const stockFilter = document.getElementById('inv-filter-stock');
    const statusFilter = document.getElementById('inv-filter-status');
    const categoryFilter = document.getElementById('inv-filter-category');

    const query = (searchInput ? searchInput.value : '').toLowerCase().trim();
    const stock = stockFilter ? stockFilter.value : 'all';
    const status = statusFilter ? statusFilter.value : 'all';
    const category = categoryFilter ? categoryFilter.value : 'all';

    inventoryCache.searchQuery = query;
    
    // Update Tags
    updateActiveFilterTags({
        query: query,
        stock: stock,
        status: status,
        category: category
    });

    // Instant client-side filtering logic
    const filtered = inventoryCache.items.filter(item => {
        const matchesQuery = !query || Object.values(item).some(val => 
            String(val).toLowerCase().includes(query)
        );

        let matchesStock = true;
        const qty = parseInt(item.Qty || item.Quantity || 0);
        if (stock === 'in') matchesStock = qty > 10;
        else if (stock === 'low') matchesStock = qty > 0 && qty <= 10;
        else if (stock === 'out') matchesStock = qty <= 0;

        const matchesStatus = status === 'all' || 
            String(item.Status || '').toLowerCase() === status.toLowerCase();

        const matchesCategory = category === 'all' || 
            String(item.Category || '').toLowerCase() === category.toLowerCase();

        return matchesQuery && matchesStock && matchesStatus && matchesCategory;
    });

    renderInventoryItems(filtered);
}

/**
 * Render Inventory Items with High Density UX
 */
function renderInventoryItems(itemsToRender = null) {
    const tableBody = document.getElementById('inventory-table-body');
    const cardContainer = document.getElementById('inventory-card-view');
    if (!tableBody || !cardContainer) return;

    // Filter items based on selected filters (Stock, Status, Category)
    const allItems = itemsToRender || inventoryCache.items;
    const { stock, status, category } = inventoryCache.filters || { stock: 'all', status: 'all', category: 'all' };

    const filtered = allItems.filter(item => {
        // 1. Stock Filter
        let matchesStock = true;
        const qty = parseInt(item.Qty || item.Quantity || 0);
        if (stock === 'in') matchesStock = qty > 10;
        else if (stock === 'low') matchesStock = qty > 0 && qty <= 10;
        else if (stock === 'out') matchesStock = qty <= 0;

        // 2. Status Filter
        const matchesStatus = status === 'all' || 
            String(item.Status || '').toLowerCase() === status.toLowerCase();

        // 3. Category Filter
        const matchesCategory = category === 'all' || 
            String(item.Category || '').toLowerCase() === category.toLowerCase();

        return matchesStock && matchesStatus && matchesCategory;
    });

    // Store filtered items for pagination count
    inventoryCache.filteredItems = filtered;
    
    // Calculate page count based on filtered items
    const pageSize = 10;
    inventoryCache.pageCount = Math.ceil(filtered.length / pageSize) || 1;
    
    // Validation: Ensure precisely 10 records per page maximum
    const startIndex = (inventoryCache.page - 1) * pageSize;
    const items = filtered.slice(startIndex, startIndex + pageSize);
    
    // Update pagination UI with current state
    updateInventoryPaginationUI(inventoryCache.page, inventoryCache.pageCount, (startIndex + pageSize) < filtered.length);

    const densityClass = currentTableDensity === 'compact' ? 'py-2 px-4' : 'py-4 px-6';

    if (items.length === 0) {
        const query = inventoryCache.searchQuery;
        tableBody.innerHTML = `<tr><td colspan="100%" class="text-center py-20 bg-white">
            <div class="flex flex-col items-center justify-center space-y-4">
                <div class="w-16 h-16 bg-gray-50 rounded-full flex items-center justify-center">
                    <i class="fa-solid fa-magnifying-glass text-gray-200 text-2xl"></i>
                </div>
                <div>
                    <p class="text-gray-900 font-bold">No items found</p>
                    <p class="text-gray-400 text-sm mt-1">Try adjusting your filters or search terms</p>
                </div>
                <button onclick="loadInventory(true)" class="text-indigo-600 font-bold text-xs uppercase tracking-widest hover:text-indigo-700">Clear all filters</button>
            </div>
        </td></tr>`;
        cardContainer.innerHTML = `<div class="col-span-full text-center py-12 text-gray-500 italic">No matching records found</div>`;
        return;
    }

    const columns = inventoryCache.columns.length > 0 ? inventoryCache.columns : ['ID', 'Item', 'Category', 'Status', 'Qty', 'DateAcquired', 'Remarks'];
    
    // Table Rendering (Desktop)
    tableBody.innerHTML = items.map((item, idx) => {
        const id = String(item.ID || item.id || '');
        
        const cells = columns.map(col => {
             const val = item[col];
             const lowerCol = String(col).toLowerCase();
             
             if (lowerCol === 'status') {
                 return `<td class="${densityClass} border-b border-gray-100">${getStatusBadgeHtml(val, item.Qty || item.Quantity)}</td>`;
             }
             
             if (lowerCol === 'qty' || lowerCol === 'quantity') {
                 return `<td class="${densityClass} border-b border-gray-100">${getStockProgressHtml(val)}</td>`;
             }

             if (lowerCol === 'id' || lowerCol.includes('serial') || lowerCol.includes('cost')) {
                 const displayVal = lowerCol.includes('cost') ? formatCurrency(val) : (val || '-');
                 return `<td class="${densityClass} border-b border-gray-100 font-mono text-[11px] text-gray-600 tracking-tight">${displayVal}</td>`;
             }

             if (lowerCol.includes('date')) {
                 return `<td class="${densityClass} border-b border-gray-100 text-xs text-gray-500 font-medium">${formatDate(val)}</td>`;
             }

             return `<td class="${densityClass} border-b border-gray-100 text-sm font-semibold text-gray-700 truncate max-w-[200px]" title="${val || ''}">${val !== undefined && val !== null ? val : '-'}</td>`;
        }).join('');
        
        const actions = `
            <td class="${densityClass} sticky right-0 z-20 bg-white group-hover:bg-slate-50 transition-colors border-b border-gray-100 text-right">
                <div class="flex items-center justify-end gap-1.5">
                    <button onclick="openEditModal('${id}')" class="w-7 h-7 flex items-center justify-center rounded-lg text-gray-400 hover:bg-indigo-50 hover:text-indigo-600 transition-all" title="Edit"><i class="fas fa-pen text-[10px]"></i></button>
                    <button onclick="confirmDeleteItem('${id}')" class="w-7 h-7 flex items-center justify-center rounded-lg text-gray-400 hover:bg-rose-50 hover:text-rose-600 transition-all" title="Delete"><i class="fas fa-trash text-[10px]"></i></button>
                </div>
            </td>
        `;
        
        const rowClass = idx % 2 === 0 ? 'bg-white' : 'bg-gray-50/30';
        return `<tr class="group ${rowClass} hover:bg-slate-50 transition-all duration-150">${cells}${actions}</tr>`;
    }).join('');

    // Card Rendering (Mobile)
    cardContainer.innerHTML = items.map(item => {
        const id = item.ID || item.id;
        const statusVal = item['Status'] || item['status'] || '-';
        return `
            <div class="bg-white rounded-2xl p-5 shadow-sm border border-gray-100 flex flex-col gap-4 active:scale-[0.98] transition-transform">
                <div class="flex justify-between items-start">
                    <div class="min-w-0">
                        <h4 class="font-bold text-gray-900 truncate">${item.Item || item.ItemName || 'Unnamed Item'}</h4>
                        <p class="text-[10px] text-gray-400 mt-0.5 tracking-widest uppercase font-black">${item.Category || 'Uncategorized'}</p>
                    </div>
                    ${getStatusBadgeHtml(statusVal, item.Qty || item.Quantity)}
                </div>
                <div class="py-3 border-y border-gray-50 my-1">
                    ${getStockProgressHtml(item.Qty || item.Quantity)}
                </div>
                <div class="flex items-center justify-between mt-auto pt-2">
                    <p class="text-[10px] text-gray-400 font-mono">${item.ID || '-'}</p>
                    <div class="flex gap-2">
                        <button onclick="openEditModal('${id}')" class="w-9 h-9 flex items-center justify-center rounded-xl bg-indigo-50 text-indigo-600 shadow-sm active:bg-indigo-600 active:text-white transition-colors"><i class="fas fa-pen text-xs"></i></button>
                        <button onclick="confirmDeleteItem('${id}')" class="w-9 h-9 flex items-center justify-center rounded-xl bg-rose-50 text-rose-600 shadow-sm active:bg-rose-600 active:text-white transition-colors"><i class="fas fa-trash text-xs"></i></button>
                    </div>
                </div>
            </div>
        `;
    }).join('');

    syncResponsiveView();
}

/**
 * Sync View State based on device
 */
function syncResponsiveView() {
    const tableWrapper = document.querySelector('.overflow-x-auto');
    const cardView = document.getElementById('inventory-card-view');
    if (!tableWrapper || !cardView) return;

    if (window.innerWidth < 1024) { // lg breakpoint
        tableWrapper.classList.add('hidden');
        cardView.classList.remove('hidden');
    } else {
        tableWrapper.classList.remove('hidden');
        cardView.classList.add('hidden');
    }
}

// Add resize listener
window.addEventListener('resize', syncResponsiveView);

/**
 * Updated showInventoryAlert to use showToast
 */
function showInventoryAlert(message, title = 'Notification', type = 'info') {
    // Map internal types to toast types
    let toastType = 'success';
    const lowerTitle = String(title).toLowerCase();
    const lowerType = String(type).toLowerCase();
    
    if (lowerTitle.includes('error') || lowerType.includes('error') || lowerTitle.includes('fail')) {
        toastType = 'error';
    }
    
    showToast(message, toastType);
}

async function loadInventory(forceReload = false) {
    if (inventoryCache.loaded && !forceReload) {
        renderInventoryItems();
        return;
    }

    // Reset local items only if we are specifically fetching new results
    if (forceReload) {
        inventoryCache.page = 1;
        inventoryCache.items = [];
    }
    
    showSkeleton('inventory', true);
    
    try {
        const promises = [fetchInventoryBatch(forceReload)];
        
        // Parallelize column/sheet fetch only if not yet loaded
        if (!inventoryCache.sheetsLoaded) promises.push(loadSheets());
        if (!inventoryCache.columnsLoaded) promises.push(fetchInventoryColumns());
        
        await Promise.all(promises);
        inventoryCache.loaded = true;
    } catch (error) {
        console.error('Failed to load inventory:', error);
        showToast('Failed to load inventory data', 'error');
    } finally {
        showSkeleton('inventory', false);
    }
}

async function loadSheets(forceReload = false) {
    const sheetSelect = document.getElementById('inv-sheet-select');
    if (!sheetSelect) return;

    // Use cached sheets if available and not force reloading
    if (inventoryCache.sheetsLoaded && !forceReload && inventoryCache.sheets.length > 0) {
        renderSheetOptions();
        return;
    }

    try {
        const cacheBuster = forceReload ? `&_t=${Date.now()}` : '';
        const response = await callApi(`?action=getSheets${cacheBuster}`);
        const sheets = response && response.sheets ? response.sheets : (Array.isArray(response) ? response : []);
        
        if (sheets.length > 0) {
            inventoryCache.sheets = sheets;
            inventoryCache.sheetsLoaded = true;
            inventoryCache.currentSheet = response.currentSheet || 'Inventory';
            renderSheetOptions();
        } else {
            sheetSelect.innerHTML = '<option value="">No sheets found</option>';
        }
    } catch (error) {
        console.error('Error loading sheets:', error);
        sheetSelect.innerHTML = '<option value="">Error loading</option>';
    }
}

function renderSheetOptions() {
    const sheetSelect = document.getElementById('inv-sheet-select');
    if (!sheetSelect) return;
    
    const sheets = inventoryCache.sheets;
    const currentSheet = inventoryCache.currentSheet || sheetSelect.value || 'Inventory';
    
    sheetSelect.innerHTML = sheets.map(sheet => 
        `<option value="${sheet}" ${sheet === currentSheet ? 'selected' : ''}>${sheet}</option>`
    ).join('');
    
    sheetSelect.classList.remove('animate-pulse');
    
    // Set up onchange if not already bound
    if (!sheetSelect._onchangeBound) {
        sheetSelect._onchangeBound = true;
        sheetSelect.onchange = async (e) => {
            const newSheet = e.target.value;
            if (!newSheet) return;
            
            sheetSelect.disabled = true;
            try {
                const res = await callApi('?action=setInventorySheetName', {
                    method: 'POST',
                    body: JSON.stringify({ name: newSheet })
                });
                
                if (res.success) {
                    showToast(`Switched to sheet: ${newSheet}`);
                    inventoryCache.currentSheet = newSheet;
                    inventoryCache.columnsLoaded = false;
                    await loadInventory(true);
                } else {
                    throw new Error(res.error || 'Failed to switch sheet');
                }
            } catch (error) {
                showToast(`Error switching sheet: ${error.message}`, 'error');
            } finally {
                sheetSelect.disabled = false;
            }
        };
    }
}

function openDeleteSheetModal() {
    const sheetSelect = document.getElementById('inv-sheet-select');
    const modal = document.getElementById('deleteSheetModal');
    const nameEl = document.getElementById('delete-sheet-name');
    const targetInput = document.getElementById('delete-sheet-target');
    const confirmInput = document.getElementById('delete-sheet-confirm');
    const btn = document.getElementById('submit-delete-sheet-btn');
    if (!sheetSelect || !modal || !nameEl || !targetInput || !confirmInput || !btn) return;
    const name = sheetSelect.value;
    if (!name) {
        showInventoryAlert('No sheet selected to delete.', 'Error');
        return;
    }
    nameEl.textContent = name;
    targetInput.value = name;
    confirmInput.value = '';
    btn.disabled = true;
    modal.classList.remove('hidden');
}

function deleteSheetValidationChanged() {
    const input = document.getElementById('delete-sheet-confirm');
    const btn = document.getElementById('submit-delete-sheet-btn');
    if (!input || !btn) return;
    btn.disabled = String(input.value).trim().toUpperCase() !== 'DELETE';
}

async function submitDeleteSheet() {
    const targetInput = document.getElementById('delete-sheet-target');
    const modal = document.getElementById('deleteSheetModal');
    const name = targetInput ? targetInput.value : '';
    if (!name) {
        showInventoryAlert('Invalid sheet.', 'Error');
        return;
    }
    try {
        const res = await callApi('?action=deleteInventorySheet', {
            method: 'POST',
            body: JSON.stringify({ name })
        });
        if (res && res.success) {
            showInventoryAlert(`Deleted sheet: ${name}`, 'Success');
            closeModal('deleteSheetModal');
            inventoryCache.columns = [];
            inventoryCache.items = [];
            inventoryCache.page = 1;
            inventoryCache.hasMore = true;
            await loadSheets();
            await loadInventory(true);
        } else {
            throw new Error(res && res.error ? res.error : 'Delete failed');
        }
    } catch (error) {
        showInventoryAlert(error.message || 'Delete failed', 'Error');
        closeModal('deleteSheetModal');
    }
}

/**
 * Render Inventory Headers with Client-side Sorting
 */
function renderInventoryHeaders() {
    const headerRow = document.getElementById('inventory-table-header');
    if (!headerRow) return;
    
    const columns = inventoryCache.columns.length > 0 ? inventoryCache.columns : ['ID', 'Item', 'Category', 'Status', 'Qty', 'DateAcquired', 'Remarks'];
    
    let headerHtml = '';

    headerHtml += columns.map(col => {
        const isSortable = true; // Most columns are sortable
        const sortIcon = `<i class="fa-solid fa-sort ml-1.5 opacity-30 group-hover:opacity-100 transition-opacity"></i>`;
        
        return `
            <th scope="col" onclick="handleInventorySort('${col}')" class="px-6 py-3 text-left text-[10px] font-black text-gray-400 uppercase tracking-widest bg-gray-50 border-b border-gray-200 cursor-pointer group hover:bg-gray-100 transition-colors">
                <div class="flex items-center">
                    ${col}
                    ${isSortable ? sortIcon : ''}
                </div>
            </th>
        `;
    }).join('');

    // Sticky Action Header
    headerHtml += `
        <th scope="col" class="px-6 py-3 text-right text-[10px] font-black text-gray-400 uppercase tracking-widest bg-gray-50 border-b border-gray-200 sticky right-0 z-40 shadow-[-10px_0_10px_-10px_rgba(0,0,0,0.1)]">
            Actions
        </th>
    `;
    
    headerRow.innerHTML = headerHtml;
}

let inventorySortState = { column: null, direction: 'asc' };

function handleInventorySort(column) {
    if (inventorySortState.column === column) {
        inventorySortState.direction = inventorySortState.direction === 'asc' ? 'desc' : 'asc';
    } else {
        inventorySortState.column = column;
        inventorySortState.direction = 'asc';
    }

    // Visual feedback on headers
    const headers = document.querySelectorAll('#inventory-table-header th');
    headers.forEach(th => {
        const icon = th.querySelector('i.fa-sort, i.fa-sort-up, i.fa-sort-down');
        if (icon) {
            if (th.textContent.trim().includes(column)) {
                icon.className = `fa-solid fa-sort-${inventorySortState.direction === 'asc' ? 'up' : 'down'} ml-1.5 text-indigo-600 opacity-100`;
            } else {
                icon.className = `fa-solid fa-sort ml-1.5 opacity-30`;
            }
        }
    });

    const sorted = [...inventoryCache.items].sort((a, b) => {
        let valA = a[column];
        let valB = b[column];

        // Handle numeric values
        if (column.toLowerCase().includes('qty') || column.toLowerCase().includes('cost')) {
            valA = parseNumber(valA);
            valB = parseNumber(valB);
            return inventorySortState.direction === 'asc' ? valA - valB : valB - valA;
        }

        // Handle dates
        if (column.toLowerCase().includes('date')) {
            valA = new Date(valA || 0).getTime();
            valB = new Date(valB || 0).getTime();
            return inventorySortState.direction === 'asc' ? valA - valB : valB - valA;
        }

        // Default string sort
        valA = String(valA || '').toLowerCase();
        valB = String(valB || '').toLowerCase();
        return inventorySortState.direction === 'asc' 
            ? valA.localeCompare(valB) 
            : valB.localeCompare(valA);
    });

    renderInventoryItems(sorted);
}

async function fetchInventoryColumns(forceReload = false) {
    if (inventoryCache.columnsLoaded && !forceReload && inventoryCache.columns.length > 0) {
        renderInventoryHeaders();
        return;
    }

    try {
        const cacheBuster = forceReload ? `&_t=${Date.now()}` : '';
        const response = await callApi(`?action=getInventorySheetColumns${cacheBuster}`);
        if (response && response.columns && Array.isArray(response.columns)) {
            inventoryCache.columns = response.columns;
            inventoryCache.columnsLoaded = true;
        } else {
             console.warn('Using fallback columns');
             inventoryCache.columns = ['ID', 'Item', 'Category', 'Status', 'Qty', 'DateAcquired', 'Remarks'];
        }
        renderInventoryHeaders();
        if (inventoryCache.items.length > 0) {
            renderInventoryItems();
        }
    } catch (error) {
        console.error('Error fetching inventory columns:', error);
        inventoryCache.columns = ['ID', 'Item', 'Category', 'Status', 'Qty', 'DateAcquired', 'Remarks'];
        renderInventoryHeaders();
    }
}

// Sheet Management
function openCreateSheetModal() {
    const modal = document.getElementById('createSheetModal');
    if (!modal) return;
    
    const nameInput = document.getElementById('new-sheet-name');
    if (nameInput) nameInput.value = '';
    
    const container = document.getElementById('new-sheet-columns-container');
    if (container) {
        const defaultCols = ['Project', 'Category', 'Item', 'BrandModel', 'Serial', 'Qty', 'Unit', 'UnitCost', 'DateAcquired', 'ProcurementProject', 'PersonInCharge', 'Location', 'Status', 'Remarks'];
        container.innerHTML = defaultCols.map(col => `
            <div class="flex items-center gap-2 bg-gray-50 p-2 rounded-lg border border-gray-100 group">
                <div class="cursor-move text-gray-400 px-1"><i class="fa-solid fa-grip-vertical text-xs"></i></div>
                <input type="text" class="bg-transparent border-none p-0 text-sm focus:ring-0 w-full text-gray-700 font-medium" value="${col}" placeholder="Column Name">
                <button type="button" onclick="this.parentElement.remove()" class="text-gray-400 hover:text-red-500 transition-colors p-1 rounded-md hover:bg-red-50 opacity-0 group-hover:opacity-100"><i class="fa-solid fa-trash-can text-xs"></i></button>
            </div>
        `).join('');
    }
    
    modal.classList.remove('hidden');
}

function addNewSheetColumnInput() {
    const container = document.getElementById('new-sheet-columns-container');
    if (!container) return;
    
    const div = document.createElement('div');
    div.className = 'flex items-center gap-2 bg-indigo-50 p-2 rounded-lg border border-indigo-100 group';
    div.innerHTML = `
        <div class="cursor-move text-indigo-300 px-1"><i class="fa-solid fa-grip-vertical text-xs"></i></div>
        <input type="text" class="bg-transparent border-none p-0 text-sm focus:ring-0 w-full text-indigo-700 font-medium" placeholder="New Column Name">
        <button type="button" onclick="this.parentElement.remove()" class="text-red-400 hover:text-red-600 transition-colors p-1 rounded-md hover:bg-red-50"><i class="fa-solid fa-trash-can text-xs"></i></button>
    `;
    container.appendChild(div);
    div.querySelector('input[type="text"]').focus();
}

async function submitCreateSheet() {
    const nameInput = document.getElementById('new-sheet-name');
    const name = nameInput ? nameInput.value.trim() : '';
    
    if (!name) {
        showInventoryAlert('Please enter a sheet name.', 'Error');
        return;
    }
    
    const container = document.getElementById('new-sheet-columns-container');
    const columns = [];
    if (container) {
        container.querySelectorAll('div.flex').forEach(div => {
            const textInput = div.querySelector('input[type="text"]');
            if (textInput) {
                const colName = textInput.value.trim();
                if (colName) columns.push(colName);
            }
        });
    }

    if (columns.length === 0) {
         showInventoryAlert('Please add at least one column.', 'Error');
         return;
    }

    try {
        const response = await callApi('?action=createInventorySheet', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name, columns })
        });
        
        if (response.success) {
            showInventoryAlert(`Sheet "${response.name}" created successfully!`, 'Success');
            closeModal('createSheetModal');
            loadSheets(); // Refresh sheets dropdown
            loadInventory(true); // Reload with new sheet
        } else {
            throw new Error(response.error || 'Failed to create sheet');
        }
    } catch (error) {
        showInventoryAlert(`Error: ${error.message}`, 'Error');
    }
}

async function openEditColumnsModal() {
    const modal = document.getElementById('editColumnsModal');
    const list = document.getElementById('edit-columns-list');
    if (!modal || !list) return;
    
    modal.classList.remove('hidden');
    list.innerHTML = '<div class="text-center py-4 text-gray-500"><i class="fa-solid fa-circle-notch fa-spin"></i> Loading columns...</div>';
    
    try {
        const response = await callApi('?action=getInventorySheetColumns');
        if (response && response.columns && Array.isArray(response.columns)) {
            list.innerHTML = response.columns.map(col => `
                <div class="flex items-center justify-between bg-gray-50 p-3 rounded-xl border border-gray-100 group">
                    <span class="text-sm font-medium text-gray-700">${col}</span>
                    <div class="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                        <button onclick="promptRenameColumn('${col}')" class="p-1.5 text-indigo-600 hover:bg-indigo-100 rounded-lg transition-colors" title="Rename">
                            <i class="fa-solid fa-pen-to-square text-xs"></i>
                        </button>
                        <button onclick="confirmDeleteColumn('${col}')" class="p-1.5 text-red-600 hover:bg-red-100 rounded-lg transition-colors" title="Delete">
                            <i class="fa-solid fa-trash-can text-xs"></i>
                        </button>
                    </div>
                </div>
            `).join('');
        } else {
            list.innerHTML = '<div class="text-center py-4 text-red-500">Failed to load columns.</div>';
        }
    } catch (error) {
        list.innerHTML = `<div class="text-center py-4 text-red-500">Error: ${error.message}</div>`;
    }
}

function promptRenameColumn(oldName) {
    const modal = document.getElementById('renameColumnModal');
    const oldInput = document.getElementById('rename-col-old');
    const newInput = document.getElementById('rename-col-new');
    
    if (modal && oldInput && newInput) {
        oldInput.value = oldName;
        newInput.value = oldName;
        modal.classList.remove('hidden');
        newInput.focus();
        newInput.select();
    }
}

function closeRenameColumnModal() {
    const modal = document.getElementById('renameColumnModal');
    if (modal) modal.classList.add('hidden');
}

async function submitRenameColumn() {
    const oldName = document.getElementById('rename-col-old').value;
    const newName = document.getElementById('rename-col-new').value.trim();
    
    if (!newName || oldName === newName) {
        closeRenameColumnModal();
        return;
    }
    
    try {
        const response = await callApi('?action=renameInventorySheetColumn', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ oldName, newName })
        });
        
        if (response.success) {
            showInventoryAlert('Column renamed successfully!', 'Success');
            closeRenameColumnModal();
            openEditColumnsModal(); // Refresh list
            fetchInventoryColumns(); // Refresh table headers
        } else {
            throw new Error(response.error || 'Failed to rename column');
        }
    } catch (error) {
        showInventoryAlert(`Error: ${error.message}`, 'Error');
    }
}

function confirmDeleteColumn(colName) {
    const modal = document.getElementById('deleteColumnModal');
    const nameSpan = document.getElementById('delete-col-name');
    const targetInput = document.getElementById('delete-col-target');
    
    if (modal && nameSpan && targetInput) {
        nameSpan.textContent = colName;
        targetInput.value = colName;
        modal.classList.remove('hidden');
    }
}

async function submitDeleteColumn() {
    const name = document.getElementById('delete-col-target').value;
    
    if (!name) {
        closeModal('deleteColumnModal');
        return;
    }

    // Prevent deletion of essential columns
    const required = SYSTEM_CONFIG.requiredFields || ['ID', 'Item', 'Qty', 'Status'];
    if (required.map(c => c.toLowerCase()).includes(name.toLowerCase())) {
        showInventoryAlert(`Cannot delete a required column: "${name}".`, 'Error');
        return;
    }
    
    try {
        const response = await callApi('?action=removeInventorySheetColumn', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name })
        });
        
        if (response.success) {
            showInventoryAlert(`Column "${name}" deleted successfully!`, 'Success');
            closeModal('deleteColumnModal');
            openEditColumnsModal(); // Refresh list
            fetchInventoryColumns(); // Refresh table headers
        } else {
            throw new Error(response.error || 'Failed to delete column');
        }
    } catch (error) {
        showInventoryAlert(`Error: ${error.message}`, 'Error');
    }
}

function promptAddNewColumn() {
    const modal = document.getElementById('addColumnModal');
    const input = document.getElementById('new-column-name-input');
    if (modal && input) {
        input.value = '';
        modal.classList.remove('hidden');
        input.focus();
    }
}

async function submitAddColumn() {
    const nameInput = document.getElementById('new-column-name-input');
    const name = nameInput ? nameInput.value.trim() : '';
    
    if (!name) {
        closeModal('addColumnModal');
        return;
    }
    
    try {
        const response = await callApi('?action=addInventorySheetColumn', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name })
        });
        
        if (response.success) {
            showInventoryAlert(`Column "${name}" added successfully!`, 'Success');
            closeModal('addColumnModal');
            openEditColumnsModal(); // Refresh list
            fetchInventoryColumns(); // Refresh table headers
        } else {
            throw new Error(response.error || 'Failed to add column');
        }
    } catch (error) {
        showInventoryAlert(`Error: ${error.message}`, 'Error');
    }
}



async function fetchInventoryBatch(forceReload = false, targetPage = null) {
    if (targetPage !== null) {
        inventoryCache.page = targetPage;
    } else if (!inventoryCache.hasMore && !forceReload) {
        showSkeleton('inventory', false);
        return;
    }

    // Show loading indicators for filtering/fetching
    showSkeleton('inventory', true);
    if (forceReload) {
        LoadingManager.showSyncToast('Filtering records...');
    }

    try {
        const query = String(inventoryCache.searchQuery || '').trim();
        
        // Use a more balanced limit for fetching to speed up Google Sheets response
        const page = 1; 
        const limit = 50; // Further reduced for maximum speed
        
        const cacheBuster = forceReload ? '&nocache=true' : '';
        const params = `?action=getItems&limit=${limit}&page=${page}&search=${encodeURIComponent(query)}${cacheBuster}`;
        const payload = await callApi(params);
        
        if (!payload || (payload.error && !payload.items)) {
            throw new Error(payload.error || 'Invalid data received');
        }

        const items = Array.isArray(payload.items) ? payload.items.map(normalizeItemForUi) : [];
        
        // Update Filter Dropdowns (Categories/Statuses) if needed
        if (items.length > 0) {
            updateInventoryFiltersFromData(items);
        }

        // Store items for client-side filtering and pagination
        inventoryCache.items = items;
        inventoryCache.hasMore = false; // We fetch all we can in one go now
        
        renderInventoryItems();

        // Scroll to top of table on page change
        const tableContainer = document.getElementById('inventory-content');
        if (tableContainer && targetPage !== null) {
            tableContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    } catch (error) {
        console.error('Error fetching inventory batch:', error);
        showToast(`Failed to load items: ${error.message}`, 'error');
    } finally {
        showSkeleton('inventory', false);
        LoadingManager.hideSyncToast();
    }
}

/**
 * Handle Search & Filters - Refreshes or Re-renders based on filter type
 */
function handleInventorySearch(event) {
    const searchInput = document.getElementById('inv-search-input');
    const stockFilter = document.getElementById('inv-filter-stock');
    const statusFilter = document.getElementById('inv-filter-status');
    const categoryFilter = document.getElementById('inv-filter-category');

    const query = (searchInput ? searchInput.value : '').toLowerCase().trim();
    const stock = stockFilter ? stockFilter.value : 'all';
    const status = statusFilter ? statusFilter.value : 'all';
    const category = categoryFilter ? categoryFilter.value : 'all';

    inventoryCache.searchQuery = query;
    inventoryCache.filters = { stock, status, category };
    
    // Update Tags
    updateActiveFilterTags({
        query: query,
        stock: stock,
        status: status,
        category: category
    });

    // Reset to page 1 for new search/filter
    inventoryCache.page = 1;
    
    // Use debounce for text search, instant for dropdowns
    const isTextSearch = event && event.target && event.target.id === 'inv-search-input';
    
    if (isTextSearch) {
        clearTimeout(debounceTimeout);
        debounceTimeout = setTimeout(async () => {
            await fetchInventoryBatch(true);
        }, 300);
    } else {
        // If it's a dropdown, we can re-render immediately if we have the data
        // or fetch if we need fresh data from server
        fetchInventoryBatch(true);
    }
}

function updateInventoryPaginationUI(page, pageCount, hasMore) {
    const info = document.getElementById('inv-page-info');
    const prevBtn = document.getElementById('inv-prev-btn');
    const nextBtn = document.getElementById('inv-next-btn');
    const pageNumbersContainer = document.getElementById('inv-page-numbers');
    
    if (info) {
        const totalText = Number.isFinite(pageCount) ? ` of ${pageCount}` : '';
        info.textContent = `Page ${page}${totalText}`;
    }
    
    if (prevBtn) {
        prevBtn.disabled = page <= 1;
        prevBtn.onclick = () => {
            inventoryCache.page = Math.max(1, page - 1);
            renderInventoryItems();
            window.scrollTo({ top: 0, behavior: 'smooth' });
        };
    }
    
    if (nextBtn) {
        nextBtn.disabled = !hasMore;
        nextBtn.onclick = () => {
            inventoryCache.page = Math.min(pageCount, page + 1);
            renderInventoryItems();
            window.scrollTo({ top: 0, behavior: 'smooth' });
        };
    }

    if (pageNumbersContainer && Number.isFinite(pageCount)) {
        let pages = [];
        const maxVisible = 5;
        
        if (pageCount <= maxVisible) {
            for (let i = 1; i <= pageCount; i++) pages.push(i);
        } else {
            if (page <= 3) {
                pages = [1, 2, 3, 4, '...', pageCount];
            } else if (page >= pageCount - 2) {
                pages = [1, '...', pageCount - 3, pageCount - 2, pageCount - 1, pageCount];
            } else {
                pages = [1, '...', page - 1, page, page + 1, '...', pageCount];
            }
        }

        pageNumbersContainer.innerHTML = pages.map(p => {
            if (p === '...') return '<span class="px-2 text-gray-400">...</span>';
            const isActive = p === page;
            return `
                <button onclick="inventoryCache.page = ${p}; renderInventoryItems(); window.scrollTo({top:0, behavior:'smooth'});" 
                    class="w-8 h-8 flex items-center justify-center rounded-lg font-bold text-xs transition-all
                    ${isActive ? 'bg-indigo-600 text-white shadow-md' : 'text-gray-500 hover:bg-gray-100 hover:text-indigo-600'}">
                    ${p}
                </button>
            `;
        }).join('');
    } else if (pageNumbersContainer) {
        pageNumbersContainer.innerHTML = '';
    }
}

// Inventory Modals (Add/Edit)
function openInventoryCreateModal() {
    const modal = document.getElementById('inventoryItemModal');
    const title = document.getElementById('inventory-modal-title');
    const saveBtnLabel = document.getElementById('inventory-save-btn-label');
    const idInput = document.getElementById('inv-item-id');
    const fieldsContainer = document.getElementById('inventory-dynamic-fields');

    if (!modal || !fieldsContainer) return;

    // Reset
    if (idInput) idInput.value = '';
    if (title) title.innerText = 'Add New Item';
    if (saveBtnLabel) saveBtnLabel.innerText = 'Save';

    // Build dynamic fields based on current columns
    const columns = inventoryCache.columns.length > 0 ? inventoryCache.columns : ['ID', 'Item', 'Category', 'Status', 'Qty', 'DateAcquired', 'Remarks'];
    
    fieldsContainer.innerHTML = columns.map(col => {
        if (col === 'ID') return ''; // Skip ID
        
        const colLower = String(col).toLowerCase();
        let type = 'text';
        let step = '';
        let placeholder = `Enter ${col}`;
        
        if (colLower.includes('qty') || colLower.includes('quantity')) {
            type = 'number';
            step = '1';
        } else if (colLower.includes('cost') || colLower.includes('price') || colLower.includes('amount')) {
            type = 'number';
            step = '0.01';
        } else if (colLower.includes('date')) {
            placeholder = 'YYYY-MM-DD or Year (YYYY)';
        }

        return `
            <div class="space-y-1">
                <label class="block text-xs font-semibold text-gray-500 uppercase tracking-wider">${col}</label>
                <input type="${type}" ${step ? `step="${step}"` : ''} data-field="${col}" class="w-full border rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500/30 focus:border-indigo-500 transition-all" placeholder="${placeholder}">
            </div>
        `;
    }).join('');

    modal.classList.remove('hidden');
}

async function openEditModal(itemId) {
    try {
        const rawItem = await callApi(`?action=getItem&id=${encodeURIComponent(String(itemId))}`);
        if (!rawItem || (rawItem && rawItem.error)) throw new Error(rawItem && rawItem.error ? rawItem.error : 'Item not found');
        
        // Open the dynamic modal first to build fields
        openInventoryCreateModal();
        
        const modal = document.getElementById('inventoryItemModal');
        const title = document.getElementById('inventory-modal-title');
        const saveBtnLabel = document.getElementById('inventory-save-btn-label');
        const idInput = document.getElementById('inv-item-id');
        
        if (title) title.innerText = 'Edit Item';
        if (saveBtnLabel) saveBtnLabel.innerText = 'Update';
        if (idInput) idInput.value = itemId;

        // Fill fields
        const inputs = document.querySelectorAll('#inventory-dynamic-fields input[data-field]');
        inputs.forEach(input => {
            const field = input.getAttribute('data-field');
            if (rawItem[field] !== undefined) {
                input.value = rawItem[field];
            }
        });

    } catch (error) {
        showInventoryAlert(`Error fetching item details: ${error.message}`, 'Error');
    }
}

function submitInventoryItem() {
    const idInput = document.getElementById('inv-item-id');
    const itemId = idInput ? idInput.value : '';
    const isEditing = !!itemId;
    
    const inputs = document.querySelectorAll('#inventory-dynamic-fields input[data-field]');
    const itemData = {};
    const errors = [];

    inputs.forEach(input => {
        const field = input.getAttribute('data-field');
        const val = input.value.trim();
        itemData[field] = val;

        // Basic Validation
        const fieldLower = field.toLowerCase();
        
        // 1. Required Fields check (Item Name is usually mandatory)
        if ((fieldLower === 'item' || fieldLower === 'itemname') && !val) {
            errors.push('Item Name is required');
            input.classList.add('border-red-500');
        } else if (fieldLower === 'qty' || fieldLower === 'quantity') {
            // 2. Quantity Validation
            if (!val || isNaN(val) || parseInt(val) < 0) {
                errors.push('Quantity must be a positive number');
                input.classList.add('border-red-500');
            } else {
                input.classList.remove('border-red-500');
            }
        } else if (fieldLower.includes('cost') && val && (isNaN(val) || parseFloat(val) < 0)) {
            // 3. Unit Cost Validation (if provided)
            errors.push('Unit Cost must be a valid price');
            input.classList.add('border-red-500');
        } else {
            input.classList.remove('border-red-500');
        }
    });

    if (errors.length > 0) {
        showToast(errors[0], 'error'); // Show the first error found
        return;
    }

    // Add ID for editing
    if (isEditing) {
        itemData.id = itemId;
    }

    const action = isEditing ? 'editItem' : 'addItem';
    
    // Show syncing toast for background operation
    LoadingManager.showSyncToast(isEditing ? 'Updating record...' : 'Adding new item...');
    
    // Show loading state on button
    const btn = document.getElementById('inventory-save-btn');
    if (btn) LoadingManager.setBtnLoading(btn, true, isEditing ? 'Update' : 'Save');

    callApi(`?action=${action}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(itemData)
    }).then(response => {
        if (response.success) {
            showToast(isEditing ? 'Item updated successfully!' : 'Item added successfully!');
            closeModal('inventoryItemModal');
            loadInventory(true);
        } else {
            throw new Error(response.error || 'Operation failed');
        }
    }).catch(error => {
        showToast(`Operation failed: ${error.message}`, 'error');
    }).finally(() => {
        LoadingManager.hideSyncToast();
        if (btn) LoadingManager.setBtnLoading(btn, false, isEditing ? 'Update' : 'Save');
    });
}

function closeModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) modal.classList.add('hidden');
}


function confirmDeleteItem(itemId) {
    const modal = document.getElementById('inventoryDeleteModal');
    if (!modal) return;
    const idInput = document.getElementById('inv-delete-id');
    const nameEl = document.getElementById('inv-delete-name');
    const confirmInput = document.getElementById('inv-delete-confirm');
    const confirmBtn = document.getElementById('inventory-delete-confirm-btn');
    idInput.value = itemId;
    const it = inventoryCache.items.find(i => String(i.ID) === String(itemId));
    const label = it ? (it.Item || it.ItemName || it.Description || it.Serial || itemId) : itemId;
    nameEl.textContent = label;
    if (confirmInput) confirmInput.value = '';
    if (confirmBtn) confirmBtn.disabled = true;
    modal.classList.remove('hidden');
}

function inventoryDeleteValidationChanged() {
    const input = document.getElementById('inv-delete-confirm');
    const btn = document.getElementById('inventory-delete-confirm-btn');
    if (!input || !btn) return;
    btn.disabled = String(input.value).trim().toUpperCase() !== 'DELETE';
}

async function confirmInventoryDelete() {
    const idInput = document.getElementById('inv-delete-id');
    const id = idInput ? idInput.value : '';
    try {
        await callApi('?action=deleteItem', {
            method: 'POST',
            body: JSON.stringify({ id })
        });
        showInventoryAlert('Item deleted successfully!', 'Success');
        closeModal('inventoryDeleteModal');
        inventoryCache.items = [];
        inventoryCache.page = 1;
        inventoryCache.hasMore = true;
        await loadInventory(true);
        dashboardCache.loaded = false;
    } catch (error) {
        showInventoryAlert(`Delete failed: ${error.message}`, 'Error');
        closeModal('inventoryDeleteModal');
    }
}

// Notification System
function showInventoryAlert(message, title = 'Notification', type = 'info') {
    const container = document.getElementById('notification-container');
    if (!container) return;
    
    // Infer type from title if default
    if (type === 'info') {
        const lowerTitle = String(title).toLowerCase();
        if (lowerTitle.includes('error') || lowerTitle.includes('failed')) type = 'error';
        else if (lowerTitle.includes('success')) type = 'success';
        else if (lowerTitle.includes('warning')) type = 'warning';
    }

    const id = Date.now().toString(36) + Math.random().toString(36).substr(2);
    
    // Colors and Icons
    let bgClass = 'bg-white';
    let borderClass = 'border-l-4 border-blue-500';
    let iconHtml = '<i class="fa-solid fa-circle-info text-blue-500 text-xl"></i>';
    let titleColor = 'text-gray-900';
    
    if (type === 'success') {
        borderClass = 'border-l-4 border-green-500';
        iconHtml = '<i class="fa-solid fa-circle-check text-green-500 text-xl"></i>';
    } else if (type === 'error') {
        borderClass = 'border-l-4 border-red-500';
        iconHtml = '<i class="fa-solid fa-circle-xmark text-red-500 text-xl"></i>';
    } else if (type === 'warning') {
        borderClass = 'border-l-4 border-orange-500';
        iconHtml = '<i class="fa-solid fa-triangle-exclamation text-orange-500 text-xl"></i>';
    }
    
    const toast = document.createElement('div');
    toast.id = `toast-${id}`;
    toast.className = `${bgClass} ${borderClass} shadow-lg rounded-r-lg p-4 pointer-events-auto transform transition-all duration-300 translate-x-full opacity-0 flex items-start gap-3 w-full`;
    toast.innerHTML = `
        <div class="shrink-0 mt-0.5">${iconHtml}</div>
        <div class="flex-1 min-w-0">
            <h4 class="text-sm font-semibold ${titleColor}">${title}</h4>
            <p class="text-sm text-gray-600 mt-1 break-words leading-relaxed">${message}</p>
        </div>
        <button onclick="dismissToast('${id}')" class="shrink-0 text-gray-400 hover:text-gray-600 transition-colors p-1 rounded-md hover:bg-gray-100">
            <i class="fa-solid fa-xmark text-sm"></i>
        </button>
    `;
    
    container.appendChild(toast);
    
    // Animate In
    requestAnimationFrame(() => {
        toast.classList.remove('translate-x-full', 'opacity-0');
    });
    
    // Auto Dismiss
    setTimeout(() => {
        dismissToast(id);
    }, 5000);
}

function dismissToast(id) {
    const toast = document.getElementById(`toast-${id}`);
    if (!toast) return;
    
    toast.classList.add('translate-x-full', 'opacity-0');
    setTimeout(() => {
        if (toast.parentElement) toast.parentElement.removeChild(toast);
    }, 300);
}

// Legacy close function (no-op now as toasts auto-dismiss)
function closeInventoryAlert() {}

// Reports Section
/**
 * Overhauled Reports Module Logic
 */
async function loadReports(forceReload = false) {
    if (reportsCache.loaded && !forceReload) {
        applyReportsFilters();
        return;
    }
    
    showSkeleton('reports', true);
    
    try {
        const items = await callApi('?action=getInventory');
        reportsCache.items = Array.isArray(items) ? items.map(normalizeItemForUi) : [];
        reportsCache.loaded = true;
        
        // Initial setup
        populateReportFilterCategories();
        applyReportsFilters();
        renderKPIWidgets();
    } catch (error) {
        console.error('Reports load error:', error);
        showToast('Error loading report data', 'error');
    } finally {
        showSkeleton('reports', false);
    }
}

function populateReportFilterCategories() {
    const catSelect = document.getElementById('report-filter-category');
    if (!catSelect) return;
    
    const categories = [...new Set(reportsCache.items.map(item => item.Category).filter(Boolean))];
    catSelect.innerHTML = '<option value="">All Categories</option>' + 
        categories.sort().map(cat => `<option value="${cat}">${cat}</option>`).join('');
}

function renderKPIWidgets() {
    const container = document.getElementById('reports-kpi-container');
    if (!container) return;

    const items = reportsCache.items;
    
    // 1. Total Asset Value
    const totalValue = items.reduce((sum, item) => sum + (parseNumber(item.UnitCost) * parseNumber(item.Quantity)), 0);
    
    // 2. Critical Stock (Stock <= 5 or marked as Low Stock)
    const criticalCount = items.filter(item => {
        const qty = parseNumber(item.Quantity);
        return qty > 0 && (qty <= 5 || item.Status === 'Low Stock');
    }).length;

    // 3. Total Consumables
    const consumables = items.filter(item => item.Category?.toLowerCase().includes('consumable')).length;

    // 4. Out of Stock
    const outOfStock = items.filter(item => parseNumber(item.Quantity) === 0 || item.Status === 'Out of Stock').length;

    const widgets = [
        { label: 'Total Asset Value', val: formatCurrency(totalValue), icon: 'fa-money-bill-trend-up', color: 'indigo' },
        { label: 'Critical Stock Items', val: criticalCount, icon: 'fa-triangle-exclamation', color: 'amber' },
        { label: 'Total Consumables', val: consumables, icon: 'fa-vial', color: 'blue' },
        { label: 'Out of Stock Assets', val: outOfStock, icon: 'fa-boxes-packing', color: 'rose' }
    ];

    container.innerHTML = widgets.map(w => `
        <div class="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 flex items-center gap-5 group hover:shadow-md transition-shadow">
            <div class="w-14 h-14 rounded-2xl bg-${w.color}-50 flex items-center justify-center group-hover:scale-110 transition-transform">
                <i class="fa-solid ${w.icon} text-${w.color}-600 text-2xl"></i>
            </div>
            <div>
                <p class="text-xs font-bold text-gray-400 uppercase tracking-wider">${w.label}</p>
                <p class="text-2xl font-black text-gray-800">${w.val}</p>
            </div>
        </div>
    `).join('');
}

function handleReportFilterChange() {
    reportsCache.filters = {
        query: document.getElementById('report-search-input')?.value || '',
        category: document.getElementById('report-filter-category')?.value || '',
        status: document.getElementById('report-filter-status')?.value || '',
        sortBy: document.getElementById('report-sort-by')?.value || 'newest'
    };
    applyReportsFilters();
}

function applyReportsFilters() {
    const { query, category, status, sortBy } = reportsCache.filters;
    const lowerQuery = query.toLowerCase();

    let filtered = reportsCache.items.filter(item => {
        const matchSearch = !query || 
            (item.ItemName?.toLowerCase().includes(lowerQuery)) ||
            (item.Description?.toLowerCase().includes(lowerQuery)) ||
            (item.Category?.toLowerCase().includes(lowerQuery));
        
        const matchCategory = !category || item.Category === category;
        const matchStatus = !status || item.Status === status;

        return matchSearch && matchCategory && matchStatus;
    });

    // Sorting
    filtered.sort((a, b) => {
        if (sortBy === 'alpha') return a.ItemName.localeCompare(b.ItemName);
        if (sortBy === 'stock-low') return parseNumber(a.Quantity) - parseNumber(b.Quantity);
        if (sortBy === 'stock-high') return parseNumber(b.Quantity) - parseNumber(a.Quantity);
        // Newest First (Assuming higher ID is newer)
        return parseNumber(b.ID) - parseNumber(a.ID);
    });

    reportsCache.filteredItems = filtered;
    reportsCache.currentPage = 1; // Reset pagination
    renderReportsView();
}

function setReportView(view) {
    reportsCache.view = view;
    
    const cardBtn = document.getElementById('view-toggle-card');
    const tableBtn = document.getElementById('view-toggle-table');
    const cardView = document.getElementById('report-cards-view');
    const tableView = document.getElementById('report-table-view');

    if (view === 'card') {
        cardBtn.classList.add('text-indigo-600', 'bg-white', 'shadow-sm');
        cardBtn.classList.remove('text-gray-500');
        tableBtn.classList.add('text-gray-500');
        tableBtn.classList.remove('text-indigo-600', 'bg-white', 'shadow-sm');
        cardView.classList.remove('hidden');
        tableView.classList.add('hidden');
    } else {
        tableBtn.classList.add('text-indigo-600', 'bg-white', 'shadow-sm');
        tableBtn.classList.remove('text-gray-500');
        cardBtn.classList.add('text-gray-500');
        cardBtn.classList.remove('text-indigo-600', 'bg-white', 'shadow-sm');
        tableView.classList.remove('hidden');
        cardView.classList.add('hidden');
    }
    renderReportsView();
}

function renderReportsView() {
    const { filteredItems, currentPage, itemsPerPage, view } = reportsCache;
    
    // Calculate page range
    const startIndex = (currentPage - 1) * itemsPerPage;
    const itemsToShow = filteredItems.slice(startIndex, startIndex + itemsPerPage);
    
    const cardContainer = document.getElementById('report-cards-view');
    const tableBody = document.getElementById('report-table-body');

    if (filteredItems.length === 0) {
        const emptyState = '<div class="col-span-full py-20 text-center"><p class="text-gray-400">No assets match your current filters.</p></div>';
        if (view === 'card') cardContainer.innerHTML = emptyState;
        else tableBody.innerHTML = `<tr><td colspan="7">${emptyState}</td></tr>`;
        renderPaginationControls(0); // Show no pages
        return;
    }

    if (view === 'card') {
        cardContainer.innerHTML = itemsToShow.map(item => createReportItemCard(item)).join('');
    } else {
        tableBody.innerHTML = itemsToShow.map(item => createReportItemRow(item)).join('');
    }

    // Render pagination UI
    const totalPages = Math.ceil(filteredItems.length / itemsPerPage);
    renderPaginationControls(totalPages);
}

function renderPaginationControls(totalPages) {
    const container = document.getElementById('report-pagination-container');
    if (!container) return;

    if (totalPages <= 1) {
        container.innerHTML = '';
        container.classList.add('hidden');
        return;
    }

    container.classList.remove('hidden');
    const { currentPage } = reportsCache;

    let html = `
        <div class="flex items-center justify-center gap-2 py-8">
            <button onclick="changeReportPage(${currentPage - 1})" 
                ${currentPage === 1 ? 'disabled' : ''}
                class="px-4 py-2 bg-white border border-gray-200 rounded-xl text-gray-600 hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed transition-all font-semibold text-sm">
                <i class="fa-solid fa-chevron-left mr-2"></i>Previous
            </button>
            
            <div class="flex items-center gap-1 mx-2">
                ${generatePageNumbers(currentPage, totalPages)}
            </div>

            <button onclick="changeReportPage(${currentPage + 1})" 
                ${currentPage === totalPages ? 'disabled' : ''}
                class="px-4 py-2 bg-white border border-gray-200 rounded-xl text-gray-600 hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed transition-all font-semibold text-sm">
                Next<i class="fa-solid fa-chevron-right ml-2"></i>
            </button>
        </div>
    `;

    container.innerHTML = html;
}

function generatePageNumbers(current, total) {
    let pages = [];
    const maxVisible = 5;
    
    if (total <= maxVisible) {
        for (let i = 1; i <= total; i++) pages.push(i);
    } else {
        if (current <= 3) {
            pages = [1, 2, 3, 4, '...', total];
        } else if (current >= total - 2) {
            pages = [1, '...', total - 3, total - 2, total - 1, total];
        } else {
            pages = [1, '...', current - 1, current, current + 1, '...', total];
        }
    }

    return pages.map(p => {
        if (p === '...') return '<span class="px-3 py-2 text-gray-400">...</span>';
        const isActive = p === current;
        return `
            <button onclick="changeReportPage(${p})" 
                class="w-10 h-10 flex items-center justify-center rounded-xl font-bold text-sm transition-all
                ${isActive ? 'bg-indigo-600 text-white shadow-lg shadow-indigo-200' : 'text-gray-500 hover:bg-gray-100 hover:text-gray-800'}">
                ${p}
            </button>
        `;
    }).join('');
}

function changeReportPage(page) {
    const totalPages = Math.ceil(reportsCache.filteredItems.length / reportsCache.itemsPerPage);
    if (page < 1 || page > totalPages) return;
    
    reportsCache.currentPage = page;
    renderReportsView();
    
    // Scroll back to top of content area
    const contentArea = document.getElementById('report-content-area');
    if (contentArea) contentArea.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function createReportItemCard(item) {
    const qty = parseNumber(item.Quantity);
    const isLowStock = qty > 0 && qty <= 5;
    const isOut = qty === 0;

    return `
        <div class="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden group hover:shadow-xl transition-all duration-300">
            <div class="p-5 space-y-4">
                <div class="flex items-start justify-between">
                    <span class="px-2.5 py-1 bg-gray-100 text-gray-500 text-[10px] font-bold uppercase rounded-lg">${item.Category || 'Asset'}</span>
                    ${getStatusBadge(item.Status, qty)}
                </div>
                
                <div>
                    <h3 class="text-base font-bold text-gray-800 line-clamp-1 group-hover:text-indigo-600 transition-colors">${item.ItemName}</h3>
                    <p class="text-xs text-gray-400 line-clamp-2 mt-1 h-8">${item.Description || 'No description provided.'}</p>
                </div>

                <div class="flex items-center justify-between pt-4 border-t border-gray-50">
                    <div class="flex flex-col">
                        <span class="text-[10px] text-gray-400 font-bold uppercase">Stock Level</span>
                        <span class="text-sm font-black ${isLowStock ? 'text-amber-600' : isOut ? 'text-rose-600' : 'text-gray-800'}">
                            ${qty} ${item.Unit || 'pcs'}
                            ${isLowStock ? '<i class="fa-solid fa-circle-exclamation ml-1 animate-pulse"></i>' : ''}
                        </span>
                    </div>
                    <div class="flex flex-col text-right">
                        <span class="text-[10px] text-gray-400 font-bold uppercase">Unit Cost</span>
                        <span class="text-sm font-black text-gray-800">${formatCurrency(item.UnitCost)}</span>
                    </div>
                </div>
            </div>
        </div>
    `;
}

function createReportItemRow(item) {
    const qty = parseNumber(item.Quantity);
    return `
        <tr class="hover:bg-gray-50 transition-colors">
            <td class="px-6 py-4">
                <div class="flex flex-col">
                    <span class="font-bold text-gray-800 text-sm">${item.ItemName}</span>
                    <span class="text-xs text-gray-400 truncate max-w-[200px]">${item.Description || ''}</span>
                </div>
            </td>
            <td class="px-6 py-4 text-sm text-gray-600">${item.Category || '-'}</td>
            <td class="px-6 py-4 font-black text-sm text-gray-800">${qty} ${item.Unit || ''}</td>
            <td class="px-6 py-4 text-sm text-gray-600">${formatCurrency(item.UnitCost)}</td>
            <td class="px-6 py-4">${getStatusBadge(item.Status, qty)}</td>
            <td class="px-6 py-4 text-xs text-gray-400">${formatDate(item.DateAdded)}</td>
            <td class="px-6 py-4 text-right">
                <button onclick="exportSingleItemRow('${item.ID}')" class="p-2 text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all" title="Export Asset Details">
                    <i class="fa-solid fa-file-arrow-down"></i>
                </button>
            </td>
        </tr>
    `;
}

function getStatusBadge(status, qty) {
    let colors = 'bg-gray-100 text-gray-600';
    if (status === 'In Stock' || status === 'Available') colors = 'bg-emerald-50 text-emerald-700 border border-emerald-100';
    if (status === 'Low Stock' || (qty > 0 && qty <= 5)) colors = 'bg-amber-50 text-amber-700 border border-amber-100';
    if (status === 'Out of Stock' || qty === 0) colors = 'bg-rose-50 text-rose-700 border border-rose-100';

    return `<span class="px-2.5 py-1 rounded-lg text-[10px] font-black uppercase tracking-tight ${colors}">${status || 'Unknown'}</span>`;
}

/**
 * Drawer & Export Workflow
 */
function toggleSpecialReportDrawer(show = true) {
    const overlay = document.getElementById('report-drawer-overlay');
    const drawer = document.getElementById('report-drawer');
    
    if (show) {
        overlay.classList.remove('hidden');
        setTimeout(() => overlay.classList.add('opacity-100'), 10);
        drawer.classList.remove('translate-x-full');
        populateColumnChecklist();
    } else {
        overlay.classList.remove('opacity-100');
        setTimeout(() => overlay.classList.add('hidden'), 300);
        drawer.classList.add('translate-x-full');
    }
}

function populateColumnChecklist() {
    const container = document.getElementById('drawer-column-checklist');
    if (!container) return;
    
    const columns = ['Item ID', 'Asset Name', 'Category', 'Quantity', 'Status', 'Date Acquired', 'Unit Cost', 'Remarks'];
    container.innerHTML = columns.map(col => `
        <label class="flex items-center gap-2 cursor-pointer p-2 hover:bg-white rounded-lg transition-colors">
            <input type="checkbox" checked value="${col}" class="rounded text-indigo-600 focus:ring-indigo-500">
            <span class="text-xs text-gray-600 font-medium">${col}</span>
        </label>
    `).join('');
}

function selectReportTemplate(type) {
    document.querySelectorAll('.report-template-btn').forEach(btn => {
        btn.classList.remove('active', 'border-indigo-600', 'bg-indigo-50');
        btn.classList.add('border-gray-100');
    });
    
    const activeBtn = event.currentTarget;
    activeBtn.classList.add('active', 'border-indigo-600', 'bg-indigo-50');
    activeBtn.classList.remove('border-gray-100');
    
    const projectGroup = document.getElementById('drawer-project-group');
    if (type === 'inventory') projectGroup.classList.add('hidden');
    else projectGroup.classList.remove('hidden');
}

function handleDrawerReportGenerate() {
    const templateBtn = document.querySelector('.report-template-btn.active');
    // Check if it's PAR template by checking icon or name (more robust than toString)
    const isPar = templateBtn && (templateBtn.innerHTML.includes('fa-file-signature') || templateBtn.innerText.includes('PAR Form'));
    
    // If user selected PAR template, redirect to the actual PAR page
    if (isPar) {
        loadSection('par');
        toggleSpecialReportDrawer(false);
        return;
    }

    const template = templateBtn && templateBtn.innerText.includes('Inventory') ? 'inventory' : 'par';
    const officer = document.getElementById('drawer-report-officer').value;
    const project = document.getElementById('drawer-report-project').value;
    const selectedCols = Array.from(document.querySelectorAll('#drawer-column-checklist input:checked')).map(cb => cb.value);

    if (!officer) {
        showToast('Please enter an Officer in Charge', 'error');
        return;
    }

    const reportData = reportsCache.filteredItems.map(item => {
        const row = {};
        if (selectedCols.includes('Item ID')) row['Item ID'] = item.ID;
        if (selectedCols.includes('Asset Name')) row['Asset Name'] = item.ItemName;
        if (selectedCols.includes('Category')) row['Category'] = item.Category;
        if (selectedCols.includes('Quantity')) row['Quantity'] = `${item.Quantity} ${item.Unit || ''}`;
        if (selectedCols.includes('Status')) row['Status'] = item.Status;
        if (selectedCols.includes('Date Acquired')) row['Date Acquired'] = formatDate(item.DateAdded);
        if (selectedCols.includes('Unit Cost')) row['Unit Cost'] = formatCurrency(item.UnitCost);
        if (selectedCols.includes('Remarks')) row['Remarks'] = item.Remarks || '-';
        return row;
    });

    openReportEditorWithData(template, officer, project, reportData);
    toggleSpecialReportDrawer(false);
}



function exportSingleItemRow(id) {
    const item = reportsCache.items.find(i => String(i.ID) === String(id));
    if (!item) return;
    
    const data = [{
        'Asset': item.ItemName,
        'Description': item.Description,
        'Category': item.Category,
        'Quantity': item.Quantity,
        'Status': item.Status,
        'Unit Cost': formatCurrency(item.UnitCost),
        'Total Value': formatCurrency(parseNumber(item.Quantity) * parseNumber(item.UnitCost))
    }];
    
    openReportEditorWithData('inventory', 'In-Charge Officer', 'DICT System Export', data);
}

function openReportEditorWithData(type, officer, project, data) {
    const reportTitle = {
        inventory: 'Full Inventory Report',
        par: 'Property Acknowledgement Receipt (PAR)',
        ics: 'Inventory Custodian Slip (ICS)'
    }[type];
    const headers = Object.keys(data[0] || {});
    const headerHtml = `
        <div class="text-center mb-8 pb-4 border-b-2 border-gray-700">
            <h1 class="text-3xl font-bold">${reportTitle}</h1>
            ${type === 'ics' ? `<p class="text-lg">Project: ${project}</p>` : ''}
        </div>
    `;
    let tableHtml = '<table class="w-full border-collapse"><thead><tr>';
    headers.forEach(h => { tableHtml += `<th class="text-sm">${h}</th>`; });
    tableHtml += '</tr></thead><tbody>';
    data.forEach(row => {
        tableHtml += '<tr>';
        headers.forEach(h => { tableHtml += `<td class="text-xs">${row[h]}</td>`; });
        tableHtml += '</tr>';
    });
    tableHtml += '</tbody></table>';
    const footerHtml = `
        <div class="mt-12">
            <p class="mb-8">Received by:</p>
            <div class="footer-text">
                <p class="font-semibold">${officer}</p>
                <p class="text-sm text-gray-600">Signature over Printed Name</p>
            </div>
        </div>
    `;
    const modal = document.getElementById('reportEditorModal');
    const area = document.getElementById('reportEditorArea');
    if (!modal || !area) return;
    area.innerHTML = headerHtml + tableHtml + footerHtml;
    modal.classList.remove('hidden');
}

function applyEditorCommand(cmd, value = null) {
    try { document.execCommand(cmd, false, value); } catch (_) {}
}

function insertLogoFromFile() {
    const input = document.getElementById('reportLogoInput');
    const area = document.getElementById('reportEditorArea');
    if (!input || !input.files || input.files.length === 0 || !area) return;
    const file = input.files[0];
    const reader = new FileReader();
    reader.onload = () => {
        const img = document.createElement('img');
        img.src = reader.result;
        img.alt = 'Logo';
        img.style.maxHeight = '80px';
        img.style.maxWidth = '200px';
        area.focus();
        const sel = window.getSelection();
        if (sel && sel.rangeCount) {
            const range = sel.getRangeAt(0);
            range.insertNode(img);
        } else {
            area.appendChild(img);
        }
        input.value = '';
    };
    reader.readAsDataURL(file);
}

function printEditedReport() {
    const area = document.getElementById('reportEditorArea');
    if (!area) return;
    const html = area.innerHTML;
    const win = window.open('', '_blank');
    win.document.write('<html><head><title>Print Report</title>');
    win.document.write('<link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">');
    win.document.write(`
        <style>
            body { font-family: 'Segoe UI', sans-serif; counter-reset: page; }
            @page { size: A4; margin: 20mm; }
            .page-counter::after { content: "Page " counter(page); counter-increment: page; position: fixed; bottom: 10px; right: 20px; font-size: 12px; color: #888; }
            th, td { text-align: left; padding: 8px; border: 1px solid #ddd; }
            th { background-color: #f2f2f2; }
            tbody tr:nth-child(even) { background-color: #f9f9f9; }
            .footer-text { margin-top: 40px; border-top: 1px solid #333; padding-top: 5px; width: 250px; text-align: center; }
            img { display: inline-block; }
        </style>
    `);
    win.document.write('</head><body class="p-8"><div class="page-counter"></div>');
    win.document.write(html);
    win.document.write('</body></html>');
    win.document.close();
    win.focus();
    setTimeout(() => { win.print(); win.close(); }, 500);
}
function toggleChat() {
    const container = document.getElementById('chat-container');
    if (!container) return;

    const isHidden = container.classList.contains('hidden');
    if (isHidden) {
        container.classList.remove('hidden');
        requestAnimationFrame(() => {
            container.classList.remove('scale-95', 'opacity-0');
            container.classList.add('scale-100', 'opacity-100');
        });

        const input = document.getElementById('chat-input');
        if (input) setTimeout(() => input.focus(), 50);
        return;
    }

    container.classList.add('scale-95', 'opacity-0');
    container.classList.remove('scale-100', 'opacity-100');
    setTimeout(() => container.classList.add('hidden'), 300);
}

function autoResizeChatInput(textarea) {
    if (!textarea) return;
    textarea.style.height = 'auto';
    textarea.style.height = `${Math.min(textarea.scrollHeight, 128)}px`;
}

function handleChatKeydown(event) {
    if (!event) return;
    if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        handleChatSubmit(event);
    }
}

async function handleChatSubmit(event) {
    if (event) event.preventDefault();
    const input = document.getElementById('chat-input');
    const messages = document.getElementById('chat-messages');
    if (!input || !messages) return;

    const message = String(input.value || '').trim();
    if (!message) return;

    input.value = '';
    autoResizeChatInput(input);

    appendChatMessage_({ role: 'user', text: message });
    scrollChatToBottom_();

    const submitButton = input.closest('form') ? input.closest('form').querySelector('button[type="submit"]') : null;
    if (submitButton) submitButton.disabled = true;

    try {
        const payload = await callApi('?action=geminiChat', {
            method: 'POST',
            body: JSON.stringify({ message })
        });

        if (payload && payload.success && payload.text) {
            appendChatMessage_({ role: 'assistant', text: String(payload.text) });
        } else {
            const err = payload && (payload.error || payload.message) ? String(payload.error || payload.message) : 'Chat request failed';
            appendChatMessage_({ role: 'assistant', text: err, isError: true });
        }
    } catch (error) {
        appendChatMessage_({ role: 'assistant', text: error && error.message ? error.message : 'Chat request failed', isError: true });
    } finally {
        if (submitButton) submitButton.disabled = false;
        scrollChatToBottom_();
    }
}

function appendChatMessage_({ role, text, isError }) {
    const messages = document.getElementById('chat-messages');
    if (!messages) return;

    const safeText = String(text || '').replace(/[&<>"']/g, c => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
    }[c]));

    const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

    const isUser = role === 'user';
    const wrapperClass = isUser ? 'flex items-start gap-2.5 justify-end' : 'flex items-start gap-2.5';
    const bubbleBase = 'p-3 rounded-2xl shadow-sm border text-sm whitespace-pre-wrap';
    const bubbleClass = isUser
        ? `${bubbleBase} rounded-tr-none bg-blue-600 text-white border-blue-600/20`
        : `${bubbleBase} rounded-tl-none bg-white text-gray-700 border-gray-100 ${isError ? 'text-red-700 border-red-200 bg-red-50' : ''}`;

    const avatar = isUser
        ? `<div class="w-8 h-8 rounded-full bg-indigo-100 flex items-center justify-center flex-shrink-0">
              <span class="text-xs font-semibold text-indigo-700">You</span>
           </div>`
        : `<div class="w-8 h-8 rounded-full bg-blue-100 flex items-center justify-center flex-shrink-0">
              <svg class="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9.75 17L9 20l-1 1h8l-1-1-.75-3M3 13h18M5 17h14a2 2 0 002-2V5a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"></path>
              </svg>
           </div>`;

    const content = `
        <div class="${wrapperClass}">
            ${isUser ? '' : avatar}
            <div class="flex flex-col gap-1 max-w-[85%] ${isUser ? 'items-end' : ''}">
                <div class="${bubbleClass}">${safeText}</div>
                <span class="text-[10px] text-gray-400 ${isUser ? 'mr-1' : 'ml-1'}">${time}</span>
            </div>
            ${isUser ? avatar : ''}
        </div>
    `;

    messages.insertAdjacentHTML('beforeend', content);
}

function scrollChatToBottom_() {
    const messages = document.getElementById('chat-messages');
    if (!messages) return;
    messages.scrollTop = messages.scrollHeight;
}
(function () {
    const selector = 'button[type="submit"], .btn-save';
    const getLoadingText = btn => btn.dataset.loadingText || 'Saving...';
    const getSavedText = btn => btn.dataset.savedText || 'Saved!';
    const useLoadingIcon = btn => (btn.dataset.loadingIcon || 'true') !== 'false';
    const useSavedIcon = btn => (btn.dataset.savedIcon || 'true') !== 'false';
    const showSaved = btn => (btn.dataset.showSaved || 'true') !== 'false';
    const storeHtml = btn => { if (!btn.dataset.originalHtml) btn.dataset.originalHtml = btn.innerHTML; };
    const restoreHtml = btn => { if (btn.dataset.originalHtml !== undefined) { btn.innerHTML = btn.dataset.originalHtml; delete btn.dataset.originalHtml; } };
    const setBusy = btn => { storeHtml(btn); btn.disabled = true; btn.setAttribute('aria-busy', 'true'); btn.innerHTML = `${useLoadingIcon(btn) ? '<i class="fa-solid fa-circle-notch fa-spin mr-1"></i>' : ''}${getLoadingText(btn)}`; };
    const setIdle = (btn, success) => {
        if (success && showSaved(btn)) {
            btn.innerHTML = `${useSavedIcon(btn) ? '<i class="fa-solid fa-check mr-1"></i>' : ''}${getSavedText(btn)}`;
            setTimeout(() => { restoreHtml(btn); btn.disabled = false; btn.removeAttribute('aria-busy'); }, 1200);
        } else {
            restoreHtml(btn);
            btn.disabled = false;
            btn.removeAttribute('aria-busy');
        }
    };
    const onClick = e => {
        const btn = e.target.closest(selector);
        if (!btn) return;
        if (btn.disabled) { e.preventDefault(); return; }
        setBusy(btn);
        const form = btn.closest('form');
        if (form) {
            form.addEventListener('saving:done', ev => {
                const ok = !ev.detail || ev.detail.success !== false;
                setIdle(btn, ok);
            }, { once: true });
        }
    };
    document.addEventListener('click', onClick, true);
    window.SavingButtons = {
        start(btn) { if (btn) setBusy(btn); },
        complete(btn, success = true) { if (btn) setIdle(btn, success); }
    };
})();
let PAR_EDIT_MODE = false;
function initializeParPage() {
    const tbody = document.getElementById('par-tbody');
    if (tbody && tbody.children.length === 0) parAddRow();
    
    // Auto-fill current date
    const today = new Date().toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'short', 
        day: 'numeric' 
    });
    const dateInput = document.getElementById('par-date');
    if (dateInput && !dateInput.value) dateInput.value = today;

    parRecalc();
    parUpdateLogoPreviews();
}
function toggleParEditMode() {
    PAR_EDIT_MODE = !PAR_EDIT_MODE;
    const ids = ['par-label-entity','par-label-fund','par-label-number','par-col-qty','par-col-unit','par-col-desc','par-col-prop','par-col-date','par-col-amount'];
    ids.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        el.setAttribute('contenteditable', PAR_EDIT_MODE ? 'true' : 'false');
        el.classList.toggle('bg-yellow-50', PAR_EDIT_MODE);
        el.classList.toggle('rounded', PAR_EDIT_MODE);
        el.classList.toggle('px-1', PAR_EDIT_MODE);
    });
    const btn = document.getElementById('par-edit-toggle');
    if (btn) btn.textContent = PAR_EDIT_MODE ? 'Editing...' : 'Edit Mode';
}
function parAddRow(initial) {
    const tbody = document.getElementById('par-tbody');
    if (!tbody) return;
    const tr = document.createElement('tr');
    
    // Store original inventory quantity in a data attribute
    if (initial && initial.rawQty !== undefined) {
        tr.dataset.maxQty = initial.rawQty;
    }

    tr.innerHTML = `
      <td class="px-3 py-2"><input type="number" min="0" step="1" class="w-full border rounded-lg p-2 text-sm" value="${initial&&initial.qty||1}" oninput="parHandleQtyChange(this)"></td>
      <td class="px-3 py-2"><input type="text" class="w-full border rounded-lg p-2 text-sm" value="${initial&&initial.unit||''}" readonly></td>
      <td class="px-3 py-2"><input type="text" class="w-full border rounded-lg p-2 text-sm" value="${initial&&initial.desc||''}" readonly></td>
      <td class="px-3 py-2"><input type="text" class="w-full border rounded-lg p-2 text-sm" value="${initial&&initial.prop||''}" readonly></td>
      <td class="px-3 py-2"><input type="text" class="w-full border rounded-lg p-2 text-sm" value="${initial&&initial.date||''}" readonly></td>
      <td class="px-3 py-2"><input type="number" min="0" step="0.01" class="w-full border rounded-lg p-2 text-sm" value="${initial&&initial.amount||''}" readonly></td>
      <td class="px-3 py-2 print:hidden">
        <div class="flex gap-2">
          <button class="px-3 py-1.5 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700" onclick="parOpenItemPicker(this)">Select Item</button>
          <button class="px-3 py-1.5 bg-red-600 text-white rounded-lg hover:bg-red-700" onclick="parRemoveRow(this)">Remove</button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
    parRecalc();
}

/**
 * Handle quantity change with validation against available stock
 */
function parHandleQtyChange(input) {
    const tr = input.closest('tr');
    const maxQty = tr.dataset.maxQty ? parseInt(tr.dataset.maxQty) : Infinity;
    const currentQty = parseInt(input.value) || 0;

    if (currentQty > maxQty) {
        showToast(`Warning: Input quantity (${currentQty}) exceeds available stock (${maxQty}).`, 'error');
        input.value = maxQty;
        input.classList.add('border-red-500', 'bg-red-50');
        setTimeout(() => input.classList.remove('border-red-500', 'bg-red-50'), 2000);
    } else {
        input.classList.remove('border-red-500', 'bg-red-50');
    }
    
    parRecalc();
}
function parRemoveRow(btn) {
    const tr = btn.closest('tr');
    if (!tr) return;
    tr.remove();
    parRecalc();
}
function parClearRows() {
    const tbody = document.getElementById('par-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';
    parRecalc();
}
function parRecalc() {
    const tbody = document.getElementById('par-tbody');
    const totalEl = document.getElementById('par-total');
    if (!tbody || !totalEl) return;
    let sum = 0;
    tbody.querySelectorAll('tr').forEach(tr => {
        const qtyEl = tr.querySelector('td:nth-child(1) input');
        const amtEl = tr.querySelector('td:nth-child(6) input');
        const unitCost = Number(tr.dataset.unitCost || '0');
        const qty = Number(qtyEl ? qtyEl.value : '0');
        const rowAmount = Number.isFinite(qty * unitCost) ? (qty * unitCost) : 0;
        if (amtEl) amtEl.value = rowAmount.toFixed(2);
        const n = rowAmount;
        sum += Number.isFinite(n) ? n : 0;
    });
    totalEl.value = sum.toFixed(2);
}
function collectParData() {
    const labels = {
        entity: document.getElementById('par-label-entity') ? document.getElementById('par-label-entity').textContent : 'Entity Name',
        fundCluster: document.getElementById('par-label-fund') ? document.getElementById('par-label-fund').textContent : 'Fund Cluster',
        parNumber: document.getElementById('par-label-number') ? document.getElementById('par-label-number').textContent : 'PAR Number',
        qty: document.getElementById('par-col-qty') ? document.getElementById('par-col-qty').textContent : 'Quantity',
        unit: document.getElementById('par-col-unit') ? document.getElementById('par-col-unit').textContent : 'Unit',
        description: document.getElementById('par-col-desc') ? document.getElementById('par-col-desc').textContent : 'Description',
        propertyNumber: document.getElementById('par-col-prop') ? document.getElementById('par-col-prop').textContent : 'Property Number',
        dateAcquired: document.getElementById('par-col-date') ? document.getElementById('par-col-date').textContent : 'Date Acquired',
        amount: document.getElementById('par-col-amount') ? document.getElementById('par-col-amount').textContent : 'Amount'
    };
    const header = {
        entityName: document.getElementById('par-entity') ? document.getElementById('par-entity').value : '',
        fundCluster: document.getElementById('par-fund') ? document.getElementById('par-fund').value : '',
        parNumber: document.getElementById('par-number') ? document.getElementById('par-number').value : '',
        date: document.getElementById('par-date') ? document.getElementById('par-date').value : ''
    };
    const items = [];
    const rows = document.querySelectorAll('#par-tbody tr');
    rows.forEach(tr => {
        items.push({
            itemId: tr.dataset.itemId || '',
            qty: tr.querySelector('td:nth-child(1) input') ? tr.querySelector('td:nth-child(1) input').value : '',
            unit: tr.querySelector('td:nth-child(2) input') ? tr.querySelector('td:nth-child(2) input').value : '',
            description: tr.querySelector('td:nth-child(3) input') ? tr.querySelector('td:nth-child(3) input').value : '',
            propertyNumber: tr.querySelector('td:nth-child(4) input') ? tr.querySelector('td:nth-child(4) input').value : '',
            dateAcquired: tr.querySelector('td:nth-child(5) input') ? tr.querySelector('td:nth-child(5) input').value : '',
            amount: tr.querySelector('td:nth-child(6) input') ? tr.querySelector('td:nth-child(6) input').value : ''
        });
    });
    const signatories = {
        receivedBy: {
            name: document.getElementById('par-received-name') ? document.getElementById('par-received-name').value : '',
            position: document.getElementById('par-received-position') ? document.getElementById('par-received-position').value : ''
        },
        issuedBy: {
            name: document.getElementById('par-issued-name') ? document.getElementById('par-issued-name').value : '',
            position: document.getElementById('par-issued-position') ? document.getElementById('par-issued-position').value : ''
        }
    };
    const totalAmount = document.getElementById('par-total') ? document.getElementById('par-total').value : '0';
    return { labels, header, items, signatories, totalAmount };
}

async function savePAR(e) {
    const data = collectParData();
    if (!data.header.parNumber) {
        alert('Please enter a PAR Number.');
        return;
    }
    if (!data.items || data.items.length === 0) {
        alert('Please add at least one item.');
        return;
    }

    // Open a new window immediately to avoid popup blocker
    const printWin = window.open('', '_blank');
    if (printWin) {
        printWin.document.write('<html><body style="font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;"><div><h2 style="color:#4f46e5;">Saving PAR...</h2><p>Please wait while we update the inventory.</p></div></body></html>');
        printWin.document.close();
    }

    const btn = e ? e.target : document.activeElement;
    const originalText = btn ? btn.textContent : 'Save PAR';
    if (btn) {
        btn.textContent = 'Saving...';
        btn.disabled = true;
    }

    try {
        const res = await callApi('?action=savePAR', {
            method: 'POST',
            body: JSON.stringify({ ...data, action: 'savePAR' })
        });
        if (res && res.success) {
            alert(`PAR ${data.header.parNumber} saved and inventory updated.`);
            // Use the early-opened window for print preview
            printPAR(printWin);
            // Reload inventory and dashboard to reflect changes
            if (typeof loadInventory === 'function') loadInventory(true);
            if (typeof loadDashboard === 'function') loadDashboard(true);
        } else {
            if (printWin) printWin.close();
            alert('Error saving PAR: ' + (res.error || 'Unknown error'));
        }
    } catch (err) {
        if (printWin) printWin.close();
        console.error('Save PAR error:', err);
        alert('Failed to save PAR: ' + err.message);
    } finally {
        if (btn) {
            btn.textContent = originalText;
            btn.disabled = false;
        }
    }
}
function setParPrintLogos(leftUrl, rightUrl) {
    localStorage.setItem('par_logo_left', leftUrl || '');
    localStorage.setItem('par_logo_right', rightUrl || '');
}
function parUpdateLogoPreviews() {
    const l = localStorage.getItem('par_logo_left') || '';
    const r = localStorage.getItem('par_logo_right') || '';
    const lp = document.getElementById('par-logo-left-preview');
    const rp = document.getElementById('par-logo-right-preview');
    if (lp) lp.src = l || '';
    if (rp) rp.src = r || '';
}
function parHandleLogoFile(side, input) {
    if (!input || !input.files || input.files.length === 0) return;
    const file = input.files[0];
    const reader = new FileReader();
    reader.onload = () => {
        const dataUrl = reader.result;
        if (side === 'left') {
            localStorage.setItem('par_logo_left', dataUrl);
        } else {
            localStorage.setItem('par_logo_right', dataUrl);
        }
        parUpdateLogoPreviews();
        input.value = '';
    };
    reader.readAsDataURL(file);
}
function parClearLogos() {
    localStorage.removeItem('par_logo_left');
    localStorage.removeItem('par_logo_right');
    parUpdateLogoPreviews();
}
function printPAR(existingWindow) {
    const d = collectParData();
    const logoL = localStorage.getItem('par_logo_left') || '';
    const win = existingWindow || window.open('', '_blank');
    
    if (!win) {
        alert('Popup blocked! Please allow popups to see the print preview.');
        return;
    }

    const money = n => {
        const num = Number(String(n || 0).replace(/[^0-9.-]+/g,""));
        return new Intl.NumberFormat('en-PH', { 
            style: 'currency', 
            currency: 'PHP',
            minimumFractionDigits: 2 
        }).format(num);
    };

    win.document.open();
    win.document.write(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>PAR - ${d.header.parNumber || 'Document'}</title>
            <style>
                @page {
                    size: A4;
                    margin: 0.5in;
                }
                @media print {
                    body { 
                        -webkit-print-color-adjust: exact; 
                        print-color-adjust: exact;
                        background-color: white !important;
                        margin: 0;
                        padding: 0;
                    }
                    .no-print { display: none !important; }
                    .print-container {
                        width: 100% !important;
                        box-shadow: none !important;
                        margin: 0 !important;
                        padding: 0 !important;
                    }
                }
                
                body {
                    font-family: 'Times New Roman', Times, serif;
                    color: black;
                    line-height: 1.3;
                    background-color: #f3f4f6;
                    margin: 0;
                    padding: 20px;
                }

                .print-container {
                    width: 210mm;
                    min-height: 297mm;
                    margin: 0 auto;
                    padding: 0.5in;
                    background-color: white;
                    box-shadow: 0 0 10px rgba(0,0,0,0.1);
                    position: relative;
                    box-sizing: border-box;
                }

                .form-title {
                    text-align: center;
                    font-weight: bold;
                    font-size: 14pt;
                    text-transform: uppercase;
                    margin-bottom: 25px;
                    margin-top: 10px;
                }

                .header-table { 
                    width: 100%; 
                    margin-bottom: 15px; 
                    border-collapse: collapse;
                }
                .header-table td { 
                    padding: 3px 0; 
                    vertical-align: bottom; 
                    font-size: 11pt;
                }
                .header-label { 
                    font-weight: bold; 
                    width: 1%; 
                    white-space: nowrap; 
                    padding-right: 8px !important; 
                }
                .header-underline { 
                    border-bottom: 1px solid black !important; 
                    min-width: 120px; 
                }

                .main-table {
                    width: 100%;
                    border-collapse: collapse;
                    border: 1.5px solid black;
                    table-layout: fixed;
                }
                .main-table th {
                    border: 1px solid black;
                    padding: 6px 4px;
                    font-size: 10pt;
                    text-transform: uppercase;
                    background-color: #f8f9fa !important;
                }
                .main-table td {
                    border: 1px solid black;
                    padding: 8px 6px;
                    font-size: 11pt;
                    word-wrap: break-word;
                    vertical-align: top;
                }
                .main-table thead {
                    display: table-header-group;
                }

                .footer-table {
                    width: 100%;
                    border: 1.5px solid black;
                    border-top: none;
                    border-collapse: collapse;
                    table-layout: fixed;
                    break-inside: avoid;
                }
                .footer-table td {
                    width: 50%;
                    border: 1px solid black;
                    padding: 12px;
                    vertical-align: top;
                }
                
                .sig-label { 
                    font-weight: bold; 
                    font-style: italic; 
                    text-decoration: underline; 
                    margin-bottom: 35px; 
                    display: block;
                    font-size: 11pt;
                }
                .sig-box { 
                    width: 90%; 
                    margin: 0 auto; 
                    text-align: center; 
                }
                .sig-line { 
                    border-bottom: 1px solid black; 
                    font-weight: bold; 
                    text-transform: uppercase; 
                    margin-bottom: 2px; 
                    min-height: 1.2em;
                    font-size: 11pt;
                }
                .sig-caption { 
                    font-size: 8.5pt; 
                    font-weight: bold; 
                    text-transform: uppercase; 
                    line-height: 1.1; 
                }
                
                .logo-container {
                    position: absolute;
                    top: 0.4in;
                    left: 0.5in;
                }
                .logo-container img {
                    max-height: 65px;
                    max-width: 140px;
                    object-fit: contain;
                }

                .text-center { text-align: center; }
                .text-right { text-align: right; }
                .font-bold { font-weight: bold; }
                
                .btn-print {
                    position: fixed;
                    bottom: 20px;
                    right: 20px;
                    background: #2563eb;
                    color: white;
                    padding: 12px 24px;
                    border-radius: 9999px;
                    font-weight: bold;
                    border: none;
                    cursor: pointer;
                    box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
                    font-family: sans-serif;
                }
                .btn-print:hover { background: #1d4ed8; }
            </style>
        </head>
        <body>
            <div class="print-container">
                ${logoL ? `<div class="logo-container"><img src="${logoL}"></div>` : ''}
                
                <div class="form-title">Property Acknowledgement Receipt</div>

                <table class="header-table">
                    <tr>
                        <td class="header-label">Entity Name:</td>
                        <td class="header-underline" colspan="3">${d.header.entityName || ''}</td>
                        <td colspan="2"></td>
                    </tr>
                    <tr>
                        <td class="header-label">Fund Cluster:</td>
                        <td class="header-underline">${d.header.fundCluster || ''}</td>
                        <td class="header-label" style="text-align: right; padding-left: 15px !important;">PAR No.:</td>
                        <td class="header-underline" style="width: 18%;">${d.header.parNumber || ''}</td>
                        <td class="header-label" style="text-align: right; padding-left: 15px !important;">Date:</td>
                        <td class="header-underline" style="width: 18%;">${d.header.date || ''}</td>
                    </tr>
                </table>

                <table class="main-table">
                    <thead>
                        <tr>
                            <th style="width: 10%;">Quantity</th>
                            <th style="width: 8%;">Unit</th>
                            <th style="width: 37%;">Description</th>
                            <th style="width: 15%;">Property Number</th>
                            <th style="width: 15%;">Date Acquired</th>
                            <th style="width: 15%;">Amount</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${d.items.length > 0 ? d.items.map(it => `
                            <tr>
                                <td class="text-center">${it.qty || ''}</td>
                                <td class="text-center">${it.unit || ''}</td>
                                <td>${it.description || ''}</td>
                                <td class="text-center">${it.propertyNumber || ''}</td>
                                <td class="text-center">${it.dateAcquired || ''}</td>
                                <td class="text-right font-bold">${it.amount ? money(it.amount) : ''}</td>
                            </tr>
                        `).join('') : `
                            <tr>
                                <td class="text-center">&nbsp;</td>
                                <td class="text-center"></td>
                                <td></td>
                                <td class="text-center"></td>
                                <td class="text-center"></td>
                                <td class="text-right"></td>
                            </tr>
                        `}
                    </tbody>
                    <tfoot>
                        <tr class="font-bold">
                            <td colspan="5" class="text-right uppercase p-3" style="font-size: 10pt; letter-spacing: 1px;">Total Amount</td>
                            <td class="text-right p-3">${money(d.totalAmount)}</td>
                        </tr>
                    </tfoot>
                </table>

                <table class="footer-table">
                    <tr>
                        <td>
                            <span class="sig-label">Received by:</span>
                            <div class="sig-box" style="margin-top: 30px;">
                                <div class="sig-line">${d.signatories.receivedBy.name || '&nbsp;'}</div>
                                <div class="sig-caption">Signature over Printed Name of End User</div>
                            </div>
                            <div class="sig-box" style="margin-top: 25px;">
                                <div class="sig-line">${d.signatories.receivedBy.position || '&nbsp;'}</div>
                                <div class="sig-caption">Position/Office</div>
                            </div>
                        </td>
                        <td>
                            <span class="sig-label">Issued by:</span>
                            <div class="sig-box" style="margin-top: 30px;">
                                <div class="sig-line">${d.signatories.issuedBy.name || '&nbsp;'}</div>
                                <div class="sig-caption">Signature over Printed Name of Supply and/or Property Custodian</div>
                            </div>
                            <div class="sig-box" style="margin-top: 25px;">
                                <div class="sig-line">${d.signatories.issuedBy.position || '&nbsp;'}</div>
                                <div class="sig-caption">Position/Office</div>
                            </div>
                        </td>
                    </tr>
                </table>
            </div>
            
            <button onclick="window.print()" class="btn-print no-print">
                Confirm Print
            </button>
        </body>
        </html>
    `);
    win.document.close();
    win.focus();
}

let PAR_PICK_TARGET = null;
let PAR_ITEM_CACHE = [];
async function parOpenItemPicker(btn) {
    const tr = btn.closest('tr');
    PAR_PICK_TARGET = tr;
    const modal = document.getElementById('parItemPickerModal');
    const list = document.getElementById('par-item-list');
    if (!modal || !list) return;
    modal.classList.remove('hidden');
    list.innerHTML = '<div class="text-center py-4 text-gray-500"><i class="fa-solid fa-circle-notch fa-spin"></i> Loading items...</div>';
    try {
        const items = await callApi('?action=getInventory');
        PAR_ITEM_CACHE = Array.isArray(items) ? items.map(normalizeItemForUi) : [];
        parRenderItemOptions(PAR_ITEM_CACHE);
    } catch (error) {
        list.innerHTML = '<div class="text-center py-4 text-red-500">Failed to load items.</div>';
    }
}
function parRenderItemOptions(items) {
    const list = document.getElementById('par-item-list');
    if (!list) return;
    if (!items || items.length === 0) {
        list.innerHTML = '<div class="text-center py-6 text-gray-500">No items found.</div>';
        return;
    }
    list.innerHTML = items.map(it => {
        const unitCost = Number(it.UnitCost || 0).toFixed(2);
        const date = it.DateAcquired || it.DateAdded || '';
        const prop = it.PropertyNumber || `PROP-${it.ID}`;
        const category = it.Category || '';
        const status = it.Status || '';
        return `
        <div class="flex items-center justify-between p-3 border rounded-xl hover:bg-gray-50">
          <div class="min-w-0">
            <div class="font-semibold text-gray-800">${it.ItemName || it.Item || it.Description || '(No Name)'}</div>
            <div class="text-xs text-gray-500">ID: ${it.ID} • ${category} • ${status}</div>
            <div class="text-xs text-gray-500">Property: ${prop} • Date: ${date} • UnitCost: ₱${unitCost}</div>
          </div>
          <button class="px-3 py-1.5 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700" onclick="parChooseItem('${String(it.ID).replace(/"/g,'&quot;')}')">Select</button>
        </div>`;
    }).join('');
}
function parSearchItems(query) {
    const q = String(query || '').toLowerCase();
    const filtered = PAR_ITEM_CACHE.filter(it => {
        const name = (it.ItemName || it.Item || it.Description || '').toLowerCase();
        const cat = (it.Category || '').toLowerCase();
        const status = (it.Status || '').toLowerCase();
        return name.includes(q) || cat.includes(q) || status.includes(q) || String(it.ID).toLowerCase().includes(q);
    });
    parRenderItemOptions(filtered);
}
function parChooseItem(id) {
    const modal = document.getElementById('parItemPickerModal');
    if (!PAR_PICK_TARGET) return;
    const it = PAR_ITEM_CACHE.find(x => String(x.ID) === String(id));
    if (!it) return;
    const unitEl = PAR_PICK_TARGET.querySelector('td:nth-child(2) input');
    const descEl = PAR_PICK_TARGET.querySelector('td:nth-child(3) input');
    const propEl = PAR_PICK_TARGET.querySelector('td:nth-child(4) input');
    const dateEl = PAR_PICK_TARGET.querySelector('td:nth-child(5) input');
    const qtyEl = PAR_PICK_TARGET.querySelector('td:nth-child(1) input');
    
    // Store data for validation and calculation
    PAR_PICK_TARGET.dataset.unitCost = Number(it.UnitCost || 0);
    PAR_PICK_TARGET.dataset.itemId = String(it.ID);
    PAR_PICK_TARGET.dataset.maxQty = it.Quantity || it.Qty || 0; // Capture available stock

    if (unitEl) unitEl.value = it.Unit || 'pcs';
    if (descEl) descEl.value = it.ItemName || it.Description || '';
    if (propEl) propEl.value = it.PropertyNumber || `PROP-${it.ID}`;
    
    // Use the raw date value instead of forcing it to YYYY-MM-DD
    const dateVal = it.DateAcquired || it.DateAdded || '';
    if (dateEl) dateEl.value = formatDate(dateVal); // Use formatted date for display
    
    if (qtyEl) {
        qtyEl.value = 1;
        parHandleQtyChange(qtyEl); // Initial validation
    }
    
    parRecalc();
    if (modal) modal.classList.add('hidden');
    PAR_PICK_TARGET = null;
}

// Global Initialization
document.addEventListener('DOMContentLoaded', () => {
    LoadingManager.init();
    
    // Initial Load Overlay
    LoadingManager.showOverlay('Initializing System...', 'Connecting to GovNet Infrastructure');
    
    // Optimistic Hide: Close overlay as soon as basic stats are ready, or after 4 seconds max
    const startTime = Date.now();
    const checkInitialLoad = setInterval(() => {
        const elapsed = Date.now() - startTime;
        if (dashboardCache.loaded || inventoryCache.loaded || elapsed > 4000) {
            LoadingManager.hideOverlay();
            clearInterval(checkInitialLoad);
        }
    }, 250);
});
