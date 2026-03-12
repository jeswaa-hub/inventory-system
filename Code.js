// --- Global Cache ---
const CACHE = {
  ss: null,
  config: null,
  configTimestamp: 0
};

// --- Configuration Service ---
function getSystemConfig() {
  const now = Date.now();
  // Cache config for 5 minutes to reduce PropertiesService calls
  if (CACHE.config && (now - CACHE.configTimestamp) < 300000) {
    return CACHE.config;
  }
  
  try {
    const props = PropertiesService.getScriptProperties();
    const config = {
      // API Configuration
      scriptId: props.getProperty('SCRIPT_ID') || '',
      apiBaseUrl: props.getProperty('API_BASE_URL') || '',
      
      // Database Configuration
      inventorySheetName: props.getProperty('INVENTORY_SHEET_NAME') || 'Inventory',
      auditSheetName: props.getProperty('AUDIT_SHEET_NAME') || 'AuditLogs',
      suppliersSheetName: props.getProperty('SUPPLIERS_SHEET_NAME') || 'Suppliers',
      
      // Authentication
      securityScript: props.getProperty('SECURITY_SCRIPT') || '',
      authToken: props.getProperty('AUTH_TOKEN') || '',
      
      // External Services
      geminiApiKey: props.getProperty('GEMINI_API_KEY') || '',
      webhookUrl: props.getProperty('WEBHOOK_URL') || '',
      
      // System Settings
      batchSize: parseInt(props.getProperty('BATCH_SIZE') || '20'),
      cacheTimeout: parseInt(props.getProperty('CACHE_TIMEOUT') || '300000'),
      enableLogging: props.getProperty('ENABLE_LOGGING') === 'true',
      enableCaching: props.getProperty('ENABLE_CACHING') !== 'false',
      
      // Feature Flags
      enableChat: props.getProperty('ENABLE_CHAT') === 'true',
      enableReports: props.getProperty('ENABLE_REPORTS') !== 'false',
      enableAudit: props.getProperty('ENABLE_AUDIT') !== 'false',
      
      // UI Configuration
      appName: props.getProperty('APP_NAME') || 'Inventory Management System',
      appVersion: props.getProperty('APP_VERSION') || '1.0.0',
      companyName: props.getProperty('COMPANY_NAME') || 'NBP/GovNet',
      
      // Data Validation
      requiredFields: (props.getProperty('REQUIRED_FIELDS') || 'ID,Item,Qty,Status').split(',').map(f => f.trim()),
      uniqueFields: (props.getProperty('UNIQUE_FIELDS') || 'ID').split(',').map(f => f.trim()),
      
      // External API Endpoints
      externalApis: {
        inventory: props.getProperty('EXTERNAL_INVENTORY_API') || '',
        suppliers: props.getProperty('EXTERNAL_SUPPLIERS_API') || '',
        reports: props.getProperty('EXTERNAL_REPORTS_API') || ''
      }
    };
    
    CACHE.config = config;
    CACHE.configTimestamp = now;
    return config;
  } catch (error) {
    console.error('Error loading system config:', error);
    return getDefaultConfig();
  }
}

function getDefaultConfig() {
  return {
    scriptId: '',
    apiBaseUrl: '',
    inventorySheetName: 'Inventory',
    auditSheetName: 'AuditLogs',
    suppliersSheetName: 'Suppliers',
    securityScript: '',
    authToken: '',
    geminiApiKey: '',
    webhookUrl: '',
    batchSize: 20,
    cacheTimeout: 300000,
    enableLogging: true,
    enableCaching: true,
    enableChat: false,
    enableReports: true,
    enableAudit: true,
    appName: 'Inventory Management System',
    appVersion: '1.0.0',
    companyName: 'NBP/GovNet',
    requiredFields: ['ID', 'Item', 'Qty', 'Status'],
    uniqueFields: ['ID'],
    externalApis: {
      inventory: '',
      suppliers: '',
      reports: ''
    }
  };
}

function updateSystemConfig(updates) {
  try {
    const props = PropertiesService.getScriptProperties();
    const config = getSystemConfig();
    
    Object.keys(updates).forEach(key => {
      if (key in config) {
        props.setProperty(key.toUpperCase(), String(updates[key]));
      }
    });
    
    // Clear cache to force reload
    CACHE.config = null;
    CACHE.configTimestamp = 0;
    
    return { success: true, message: 'Configuration updated successfully' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// --- Logging Service ---
function logAction(action, details, level = 'info') {
  const config = getSystemConfig();
  if (!config.enableLogging) return;
  
  const timestamp = new Date().toISOString();
  const logEntry = {
    timestamp,
    action,
    details,
    level,
    user: Session.getActiveUser().getEmail() || 'anonymous'
  };
  
  try {
    const sheet = getSheet(config.auditSheetName);
    sheet.appendRow([timestamp, action, details, level, logEntry.user]);
  } catch (error) {
    console.error('Failed to log action:', error);
  }
}

// --- Data Access Layer ---
function getDataSource(source) {
  const config = getSystemConfig();
  
  switch (source) {
    case 'inventory':
      return config.externalApis.inventory || 'internal';
    case 'suppliers':
      return config.externalApis.suppliers || 'internal';
    case 'reports':
      return config.externalApis.reports || 'internal';
    default:
      return 'internal';
  }
}

function fetchExternalData(apiUrl, options = {}) {
  try {
    const response = UrlFetchApp.fetch(apiUrl, {
      method: options.method || 'GET',
      headers: options.headers || {},
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    const text = response.getContentText();
    
    if (code < 200 || code >= 300) {
      throw new Error(`External API error (${code}): ${text}`);
    }
    
    try {
      return JSON.parse(text);
    } catch (e) {
      return { success: true, data: text };
    }
  } catch (error) {
    logAction('External API Error', `${apiUrl}: ${error.message}`, 'error');
    throw error;
  }
}

// --- Cache Management ---
const DATA_CACHE = {};

function getCachedData(key) {
  const config = getSystemConfig();
  if (!config.enableCaching) return null;
  
  const cached = DATA_CACHE[key];
  if (cached && (Date.now() - cached.timestamp) < config.cacheTimeout) {
    logAction('Cache Hit', key, 'debug');
    return cached.data;
  }
  return null;
}

function setCachedData(key, data) {
  const config = getSystemConfig();
  if (!config.enableCaching) return;
  
  DATA_CACHE[key] = {
    data: data,
    timestamp: Date.now()
  };
  logAction('Cache Set', key, 'debug');
}

function clearCache(key) {
  if (key) {
    delete DATA_CACHE[key];
    logAction('Cache Cleared', key, 'debug');
  } else {
    Object.keys(DATA_CACHE).forEach(k => delete DATA_CACHE[k]);
    logAction('Cache Cleared', 'all', 'debug');
  }
}

function getActiveSpreadsheet() {
  if (!CACHE.ss) CACHE.ss = SpreadsheetApp.getActiveSpreadsheet();
  return CACHE.ss;
}

// --- API Configuration ---
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : null;
  
  // Dispatcher
  let result = {};
  try {
    if (action === 'getDashboardStats') {
      result = getDashboardStats();
    } else if (action === 'getStats') {
      result = getDashboardStats();
    } else if (action === 'getInventory') {
      result = getInventory();
    } else if (action === 'getAllItems') {
      result = getInventory();
    } else if (action === 'getItems') {
      result = getItems_(e);
    } else if (action === 'getItem') {
      result = getItem_(e && e.parameter ? e.parameter.id : '');
    } else if (action === 'getInventoryInit') {
      const sheetName = getInventorySheetName_();
      const columns = getInventorySheetColumns_();
      const inventory = getInventory();
      result = { 
        sheetName, 
        columns: columns.columns, 
        inventory,
        sheets: listSheetNames_()
      };
    } else if (action === 'getSheets') {
      result = { 
        sheets: listSheetNames_(),
        currentSheet: getInventorySheetName_()
      };
    } else if (action === 'getInventorySheetName') {
      result = { name: getInventorySheetName_() };
    } else if (action === 'getInventorySheetColumns') {
      result = getInventorySheetColumns_();
    } else if (action === 'getSuppliers') {
      result = getSuppliers();
    } else if (action === 'getAuditLogs') {
      result = getAuditLogs();
    } else if (action === 'checkDisposalNotifications') {
      result = checkDisposalNotifications();
    } else if (action === 'verifyAccess') {
      const hash = e && e.parameter && e.parameter.hash ? e.parameter.hash : '';
      result = verifyAccess_(hash);
    } else if (action === 'getSystemConfig') {
      result = getSystemConfig();
    } else if (action === 'updateSystemConfig') {
      const updates = e && e.parameter ? e.parameter : {};
      result = updateSystemConfig(updates);
    } else if (action === 'getDataSource') {
      const source = e && e.parameter && e.parameter.source ? e.parameter.source : '';
      result = { source: getDataSource(source) };
    } else if (action === 'clearCache') {
      const key = e && e.parameter && e.parameter.key ? e.parameter.key : '';
      clearCache(key);
      result = { success: true, message: 'Cache cleared' };
    } else if (action === 'getSystemLogs') {
      const limit = parseInt(e && e.parameter && e.parameter.limit ? e.parameter.limit : '100');
      result = getSystemLogs_(limit);
    } else if (action === 'savePAR') {
      // In case it comes via GET (though it should be POST)
      result = { error: 'savePAR requires POST' };
    } else {
      result = { error: action ? 'Invalid action' : 'Missing action' };
    }
  } catch (err) {
    result = { error: err.message };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
}

function getItems_(e) {
  const config = getSystemConfig();
  const p = e && e.parameter ? e.parameter : {};
  const limitRaw = Number(p.limit || config.batchSize);
  const limit = Math.max(1, Math.min(200, Number.isFinite(limitRaw) ? limitRaw : config.batchSize));
  const pageRaw = Number(p.page || 1);
  const page = Math.max(1, Number.isFinite(pageRaw) ? pageRaw : 1);
  const search = String(p.search || '').trim().toLowerCase();
  const order = String(p.order || 'asc').trim().toLowerCase();
  const nocache = p.nocache === 'true';

  // Check cache first
  const sheetName = getInventorySheetName_();
  const cacheKey = `items_${sheetName}_${search}_${page}_${limit}_${order}`;
  if (!nocache) {
    const cached = getCachedData(cacheKey);
    if (cached) return cached;
  }

  const sheet = getInventorySheet_();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { items: [], hasMore: false };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h ?? '').trim());
  const totalRows = lastRow - 1;

  logAction('Get Items', `Fetching page ${page}, limit ${limit}, search: "${search}"`, 'debug');

  if (search) {
    const all = getInventory();
    const filtered = all.filter(it => {
      const hay = [
        it.Project,
        it.Category,
        it.Item,
        it.BrandModel,
        it.Serial,
        it.Unit,
        it.Status,
        it.Remarks,
        it.ID
      ]
        .map(v => String(v ?? '').toLowerCase())
        .join(' | ');
      return hay.includes(search);
    });
    const startIndex = (page - 1) * limit;
    const endIndex = startIndex + limit;
    const result = { 
      items: filtered.slice(startIndex, endIndex), 
      hasMore: endIndex < filtered.length,
      pageCount: Math.max(1, Math.ceil(filtered.length / limit)),
      total: filtered.length
    };
    setCachedData(cacheKey, result);
    return result;
  }

  if (order === 'desc') {
    const endRow = lastRow - (page - 1) * limit;
    if (endRow < 2) return { items: [], hasMore: false };
    const startRow = Math.max(2, endRow - limit + 1);
    const numRows = endRow - startRow + 1;
    const values = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
    const items = values
      .map((row, i) => standardizeInventoryRow_(headers, row, startRow + i))
      .reverse();
    const result = { 
      items, 
      hasMore: startRow > 2,
      pageCount: Math.max(1, Math.ceil(totalRows / limit)),
      total: totalRows
    };
    setCachedData(cacheKey, result);
    return result;
  }

  const startRow = 2 + (page - 1) * limit;
  if (startRow > lastRow) return { items: [], hasMore: false };
  const numRows = Math.min(limit, lastRow - startRow + 1);
  const values = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  const items = values.map((row, i) => standardizeInventoryRow_(headers, row, startRow + i));
  const hasMore = startRow + numRows - 1 < lastRow;
  const result = { 
    items, 
    hasMore,
    pageCount: Math.max(1, Math.ceil(totalRows / limit)),
    total: totalRows
  };
  setCachedData(cacheKey, result);
  return result;
}

function getItem_(id) {
  const sheetName = getInventorySheetName_();
  const sheet = getInventorySheet_();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { error: 'Item not found' };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h ?? '').trim());
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rowNumber = findInventoryRowNumber_(headers, rows, id);
  if (!rowNumber) return { error: 'Item not found' };

  const row = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  return standardizeInventoryRow_(headers, row, rowNumber);
}

function standardizeInventoryRow_(headers, row, rowNumber) {
  const raw = {};
  for (let j = 0; j < headers.length; j++) raw[headers[j]] = row[j];
  const rowId = `ROW#${rowNumber}`;
  const item = {
    ID: raw.ID || raw.Id || raw.id || raw.Serial || rowId,
    Project: raw.Project || raw['Project Name'] || '',
    Category: raw.Category || '',
    Item: raw.Item || '',
    BrandModel: raw.BrandModel || raw['Brand and model'] || raw['Brand and Model'] || '',
    Serial: raw.Serial || '',
    Qty: raw.Qty ?? '',
    Unit: raw.Unit || '',
    UnitCost: raw.UnitCost ?? raw['Unit Cost'] ?? '',
    DateAcquired: raw.DateAcquired || raw['Date Acquired'] || '',
    ProcurementProject:
      raw.ProcurementProject ||
      raw['Procurement/\nProject'] ||
      raw['Procurement/Project'] ||
      raw['Procurement Project'] ||
      raw.Procurement ||
      '',
    PersonInCharge: raw.PersonInCharge || raw['Person-in-charge'] || raw['Person in Charge'] || '',
    Location: raw.Location || '',
    Status: raw.Status || '',
    Remarks: raw.Remarks || ''
  };
  return { ...raw, ...item, _raw: raw };
}

function doPost(e) {
  // Handle POST requests (Actions that modify data)
  // Use text/plain to avoid CORS preflight issues on some clients
  let data;
  if (!e || !e.postData || typeof e.postData.contents !== 'string') {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Missing request body' }))
        .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    data = JSON.parse(e.postData.contents);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid JSON body' }))
        .setMimeType(ContentService.MimeType.JSON);
  }
  
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : (data && data.action ? data.action : null);
  let result = {};
  
  try {
    if (action === 'addItem') {
      result = addItem(data);
    } else if (action === 'editItem') {
      result = editItem(data.id, data);
    } else if (action === 'deleteItem') {
      result = deleteItem(data.id);
    } else if (action === 'adjustStock') {
      result = adjustStock(data.id, data.amount, data.reason);
    } else if (action === 'addSupplier') {
      result = addSupplier(data);
    } else if (action === 'geminiChat') {
      result = geminiChat(data);
    } else if (action === 'setInventorySheetName') {
      result = setInventorySheetName_(String(data && data.name ? data.name : ''));
    } else if (action === 'createInventorySheet') {
      result = createInventorySheet_(String(data && data.name ? data.name : ''), data && Array.isArray(data.columns) ? data.columns : []);
    } else if (action === 'deleteInventorySheet') {
      result = deleteInventorySheet_(String(data && data.name ? data.name : ''));
    } else if (action === 'getInventorySheetColumns') {
      result = getInventorySheetColumns_();
    } else if (action === 'renameInventorySheetColumn') {
      result = renameInventorySheetColumn_(data.oldName, data.newName);
    } else if (action === 'addInventorySheetColumn') {
      result = addInventorySheetColumn_(data.name);
    } else if (action === 'removeInventorySheetColumn') {
      result = removeInventorySheetColumn_(data.name);
    } else if (action === 'verifyAccess') {
      const hash = data && (data.hashedPassword || data.hash) ? (data.hashedPassword || data.hash) : '';
      result = verifyAccess_(hash);
    } else if (action === 'savePAR') {
      result = savePAR_(data);
    } else {
      result = { error: action ? 'Invalid action' : 'Missing action' };
    }
  } catch (err) {
    result = { error: err.message };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
}

function verifyAccess_(hash) {
  const ACCESS_SCRIPT = '1Vt_jqc3vo0Z_YMlkSTJVFGDNjB9efBC1075DVu0qbt9p_-0rZ1qfDNYC';
  const expectedHash = sha256Hex_(ACCESS_SCRIPT);
  if (!hash || !String(hash).trim()) {
    return { success: false, message: 'Missing access hash' };
  }
  if (String(hash).trim() === expectedHash) {
    return { success: true };
  }
  return { success: false, message: 'Invalid access script' };
}

function sha256Hex_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value || ''),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function geminiChat(payload) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return { error: 'Missing GEMINI_API_KEY in Script Properties' };

  const userMessage = payload && payload.message ? String(payload.message) : '';
  if (!userMessage.trim()) return { error: 'Missing message' };

  const requestedModel = (payload && payload.model ? String(payload.model) : '').trim();
  const model = pickGeminiModel_(apiKey, requestedModel);
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`;

  const context = payload && payload.context ? String(payload.context) : '';
  const systemPrompt = payload && payload.system ? String(payload.system) : '';

  const includeBackendData = payload && typeof payload.includeBackendData === 'boolean' ? payload.includeBackendData : true;
  let backendContext = '';
  if (includeBackendData) {
    try {
      backendContext = buildBackendContext_();
    } catch (e) {
      backendContext = `Backend context unavailable: ${e && e.message ? e.message : String(e)}`;
    }
  }

  const parts = [];
  if (systemPrompt.trim()) parts.push({ text: systemPrompt });
  if (backendContext.trim()) parts.push({ text: backendContext });
  if (context.trim()) parts.push({ text: context });
  parts.push({ text: userMessage });

  const body = {
    contents: [{ role: 'user', parts }]
  };

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) {
    return { error: `Gemini API error (${code}): ${text}` };
  }

  let json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    return { error: `Invalid Gemini response: ${text}` };
  }

  const out = json && json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts
    ? json.candidates[0].content.parts.map(p => p.text).filter(Boolean).join('')
    : '';

  if (!out.trim()) return { error: 'Empty Gemini response' };
  return { success: true, text: out };
}

function pickGeminiModel_(apiKey, requestedModel) {
  const available = listGeminiModels_(apiKey);
  const normalizedAvailable = Array.isArray(available) ? available : [];
  const has = name => normalizedAvailable.includes(name);

  const candidates = [
    requestedModel,
    'gemini-3.0-flash-preview',
    'gemini-3-flash-preview',
    'gemini-3.0-flash',
    'gemini-3-flash',
    'gemini-2.0-pro',
    'gemini-2.0-flash',
    'gemini-2.0-flash-lite',
    'gemini-1.5-flash-latest',
    'gemini-1.5-flash',
    'gemini-1.5-pro-latest',
    'gemini-1.5-pro'
  ]
    .map(s => String(s || '').trim())
    .filter(Boolean);

  for (let i = 0; i < candidates.length; i++) {
    const c = candidates[i];
    if (has(c)) return c;
  }

  if (normalizedAvailable.length > 0) return normalizedAvailable[0];
  return requestedModel || 'gemini-1.5-flash-latest';
}

function listGeminiModels_(apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${encodeURIComponent(apiKey)}`;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) return [];

  let json;
  try {
    json = JSON.parse(text);
  } catch (_) {
    return [];
  }

  const models = json && Array.isArray(json.models) ? json.models : [];
  const out = [];
  for (let i = 0; i < models.length; i++) {
    const m = models[i];
    const methods = Array.isArray(m.supportedGenerationMethods) ? m.supportedGenerationMethods : [];
    if (!methods.includes('generateContent')) continue;
    const name = String(m.name || '').trim();
    if (!name) continue;
    out.push(name.startsWith('models/') ? name.slice('models/'.length) : name);
  }
  return out;
}

function buildBackendContext_() {
  const toNum = v => {
    const n = Number(String(v ?? '').replace(/[^0-9.\-]+/g, ''));
    return Number.isFinite(n) ? n : 0;
  };

  const items = getInventory();
  const total = items.length;
  const lowStockItems = items
    .map(it => ({ it, qty: toNum(it.Qty) }))
    .filter(x => x.qty > 0 && x.qty <= 10)
    .sort((a, b) => a.qty - b.qty)
    .slice(0, 20)
    .map(x => `${String(x.it.Item ?? '').trim()} | Qty: ${String(x.it.Qty ?? '').trim()} | Serial: ${String(x.it.Serial ?? '').trim()} | Location: ${String(x.it.Location ?? '').trim()}`);

  const outStockItems = items
    .map(it => ({ it, qty: toNum(it.Qty) }))
    .filter(x => x.qty <= 0)
    .slice(0, 20)
    .map(x => `${String(x.it.Item ?? '').trim()} | Qty: ${String(x.it.Qty ?? '').trim()} | Serial: ${String(x.it.Serial ?? '').trim()} | Location: ${String(x.it.Location ?? '').trim()}`);

  const categoryCounts = {};
  items.forEach(it => {
    const c = String(it.Category ?? '').trim() || 'Uncategorized';
    categoryCounts[c] = (categoryCounts[c] || 0) + 1;
  });
  const topCategories = Object.entries(categoryCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(([c, n]) => `${c}: ${n}`);

  const logs = getAuditLogs();
  const recentLogs = Array.isArray(logs)
    ? logs
        .slice(-20)
        .reverse()
        .map(l => `${String(l.Timestamp ?? l.Date ?? '').trim()} | ${String(l.User ?? '').trim()} | ${String(l.Action ?? '').trim()} | ${String(l.Details ?? '').trim()}`)
    : [];

  const sampleItems = items
    .slice(0, 50)
    .map(it => `${String(it.Item ?? '').trim()} | Qty: ${String(it.Qty ?? '').trim()} | Category: ${String(it.Category ?? '').trim()} | Location: ${String(it.Location ?? '').trim()}`)
    .filter(s => s.replace(/\s+/g, '').length > 0);

  return [
    'Backend Snapshot (Google Sheets)',
    `Total items: ${total}`,
    topCategories.length ? `Top categories:\n${topCategories.join('\n')}` : '',
    lowStockItems.length ? `Low stock (1-9) sample:\n${lowStockItems.join('\n')}` : 'Low stock: none',
    outStockItems.length ? `Out of stock sample:\n${outStockItems.join('\n')}` : 'Out of stock: none',
    recentLogs.length ? `Recent audit logs (up to 20):\n${recentLogs.join('\n')}` : 'Audit logs: none',
    sampleItems.length ? `Inventory sample (up to 50):\n${sampleItems.join('\n')}` : ''
  ].filter(Boolean).join('\n\n');
}

function findHeaderIndex_(headers, candidates) {
  const cleaned = headers.map(h => String(h ?? '').trim());
  for (let i = 0; i < candidates.length; i++) {
    const candidate = candidates[i];
    const idx = cleaned.indexOf(candidate);
    if (idx > -1) return idx;
  }
  return -1;
}

function ensureInventorySheetColumns_(sheet, customColumns) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  // Default columns
  const defaultCols = [
      'Project',
      'Category',
      'Item',
      'BrandModel',
      'Serial',
      'Qty',
      'Unit',
      'UnitCost',
      'DateAcquired',
      'ProcurementProject',
      'PersonInCharge',
      'Location',
      'Status',
      'Remarks'
  ];

  // Merge custom columns if provided
  let finalCols = [...defaultCols];
  if (Array.isArray(customColumns) && customColumns.length > 0) {
      // Filter out duplicates
      const additional = customColumns.filter(c => !defaultCols.includes(c));
      finalCols = [...defaultCols, ...additional];
  }

  if (lastRow < 1 || lastCol < 1) {
    sheet.appendRow(finalCols);
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h ?? '').trim());
  }

  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h ?? '').trim());
  return headers;
}

function getInventorySheetName_() {
  const props = PropertiesService.getScriptProperties();
  const name = props.getProperty('INVENTORY_SHEET_NAME');
  
  if (name && name.trim() !== '') {
    return name;
  }
  return 'Inventory';
}

function getInventorySheet_() {
  const sheetName = getInventorySheetName_();
  const ss = getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Inventory sheet not found: ${sheetName}`);
  }
  return sheet;
}

function setInventorySheetName_(rawName) {
  const name = String(rawName || '').trim();
  if (!name) throw new Error('Sheet name is required');
  const ss = getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  
  if (!sheet) {
    // Do not create sheet automatically when selecting. 
    // User must use createInventorySheet for that.
    return { error: `Sheet "${name}" not found.` };
  }
  
  ensureInventorySheetColumns_(sheet);
  PropertiesService.getScriptProperties().setProperty('INVENTORY_SHEET_NAME', name);
  CACHE.config = null;
  CACHE.configTimestamp = 0;
  clearCache();
  logAction('Select Inventory Sheet', 'Using sheet: ' + name);
  return { success: true, name };
}

function createInventorySheet_(rawName, customColumns) {
  const base = String(rawName || '').trim() || 'Inventory';
  const ss = getActiveSpreadsheet();
  let name = base;
  let idx = 1;
  while (ss.getSheetByName(name)) {
    idx += 1;
    name = `${base} (${idx})`;
    if (idx > 50) break;
  }
  const sheet = ss.insertSheet(name);
  ensureInventorySheetColumns_(sheet, customColumns);
  PropertiesService.getScriptProperties().setProperty('INVENTORY_SHEET_NAME', name);
  CACHE.config = null;
  CACHE.configTimestamp = 0;
  logAction('Create Inventory Sheet', 'Created and selected sheet: ' + name);
  return { success: true, name };
}

function deleteInventorySheet_(rawName) {
  const ss = getActiveSpreadsheet();
  const current = getInventorySheetName_();
  const name = String(rawName || current).trim();
  if (!name) return { error: 'Sheet name is required' };
  const sheet = ss.getSheetByName(name);
  if (!sheet) return { error: `Sheet "${name}" not found.` };
  const sheets = ss.getSheets();
  if (sheets.length <= 1) return { error: 'Cannot delete the last remaining sheet.' };
  
  ss.deleteSheet(sheet);
  
  const names = listSheetNames_();
  let newName = '';
  if (names.includes('Inventory')) {
    newName = 'Inventory';
  } else {
    newName = names[0] || '';
  }
  if (newName) {
    PropertiesService.getScriptProperties().setProperty('INVENTORY_SHEET_NAME', newName);
  } else {
    PropertiesService.getScriptProperties().deleteProperty('INVENTORY_SHEET_NAME');
  }
  CACHE.config = null;
  CACHE.configTimestamp = 0;
  clearCache();
  logAction('Delete Inventory Sheet', `Deleted sheet: ${name}. Selected: ${newName || '(none)'}`);
  return { success: true, name: newName, sheets: names };
}

function savePAR_(payload) {
  const ss = getActiveSpreadsheet();
  const name = 'PAR';
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const header = payload && payload.header ? payload.header : {};
  const items = payload && Array.isArray(payload.items) ? payload.items : [];
  const signatories = payload && payload.signatories ? payload.signatories : {};
  const totalAmount = Number(payload && payload.totalAmount ? payload.totalAmount : 0);
  const cols = ['PAR Number','Entity Name','Fund Cluster','Quantity','Unit','Description','Property Number','Date Acquired','Amount','Received By Name','Received By Position','Issued By Name','Issued By Position','Total Amount','Created At'];
  if (sheet.getLastRow() === 0) sheet.getRange(1,1,1,cols.length).setValues([cols]);
  const rows = items.map(it => [
    String(header.parNumber || ''),
    String(header.entityName || ''),
    String(header.fundCluster || ''),
    Number(it.qty || 0),
    String(it.unit || ''),
    String(it.description || ''),
    String(it.propertyNumber || ''),
    String(it.dateAcquired || ''),
    Number(it.amount || 0),
    String(signatories.receivedBy && signatories.receivedBy.name || ''),
    String(signatories.receivedBy && signatories.receivedBy.position || ''),
    String(signatories.issuedBy && signatories.issuedBy.name || ''),
    String(signatories.issuedBy && signatories.issuedBy.position || ''),
    totalAmount,
    new Date()
  ]);
  if (rows.length) {
    sheet.getRange(sheet.getLastRow()+1,1,rows.length,cols.length).setValues(rows);
    
    // Automatically deduct quantity from inventory
    items.forEach(it => {
      if (it.itemId && it.qty) {
        try {
          const qtyDeduct = -Math.abs(Number(it.qty));
          adjustStock(it.itemId, qtyDeduct, `PAR Issued: ${header.parNumber || 'N/A'}`);
        } catch (e) {
          console.error(`Failed to deduct stock for item ${it.itemId}: ${e.message}`);
          logAction('PAR Stock Error', `Failed to deduct stock for item ${it.itemId}: ${e.message}`);
        }
      }
    });
  }
  logAction('Save PAR', `Saved PAR ${String(header.parNumber || '')} items: ${rows.length}`);
  return { success: true, count: rows.length };
}

function renameInventorySheetColumn_(oldName, newName) {
  const sheetName = getInventorySheetName_();
  const sheet = getInventorySheet_();
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return { error: 'Sheet is empty' };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h ?? '').trim());
  const idx = headers.indexOf(String(oldName || '').trim());
  
  if (idx === -1) {
    return { error: 'Column not found: ' + oldName };
  }
  
  sheet.getRange(1, idx + 1).setValue(String(newName || '').trim());
  logAction('Rename Column', `Renamed column '${oldName}' to '${newName}' in sheet '${sheetName}'`);
  return { success: true };
}

function addInventorySheetColumn_(name) {
  const sheetName = getInventorySheetName_();
  const sheet = getInventorySheet_();
  const lastCol = sheet.getLastColumn();
  
  // Add new header at the end
  sheet.getRange(1, lastCol + 1).setValue(name);
  
  logAction('Add Column', `Added new column "${name}" to sheet: ${sheetName}`);
  return { success: true };
}

function removeInventorySheetColumn_(name) {
  const sheetName = getInventorySheetName_();
  const sheet = getInventorySheet_();
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return { error: 'Sheet is empty' };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h ?? '').trim());
  const idx = headers.indexOf(String(name || '').trim());
  
  if (idx === -1) {
    return { error: 'Column not found: ' + name };
  }
  
  // deleteColumn is 1-based index
  sheet.deleteColumn(idx + 1);
  logAction('Delete Column', `Deleted column "${name}" from sheet: ${sheetName}`);
  return { success: true };
}

function getInventorySheetColumns_() {
  const sheetName = getInventorySheetName_();
  const sheet = getInventorySheet_();
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return { columns: [] };
  
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h ?? '').trim());
  return { columns: headers };
}

function listSheetNames_() {
  const ss = getActiveSpreadsheet();
  return ss.getSheets().map(s => s.getName());
}

function getSheet(name) {
  const ss = getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function getDataFromSheet(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const headers = data[0].map(h => String(h ?? '').trim());
  const rows = data.slice(1);
  
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i];
    });
    return obj;
  });
}

function getTransactions_() {
  const sheet = getSheet('Transactions');
  const values = sheet.getDataRange().getValues();
  if (!values || values.length === 0) return [];

  const firstRow = values[0].map(v => String(v ?? '').trim());
  const normalizedHeaders = firstRow.map(h => h.toLowerCase());
  const expected = ['id', 'timestamp', 'date', 'type', 'itemid', 'item', 'itemname', 'quantity', 'qty', 'user', 'reason'];
  const looksLikeHeader = normalizedHeaders.some(h => expected.includes(h));

  const rows = looksLikeHeader ? values.slice(1) : values;
  const tz = Session.getScriptTimeZone();

  if (looksLikeHeader) {
    const idx = name => normalizedHeaders.indexOf(name);
    const idIdx = idx('id');
    const tsIdx = idx('timestamp') > -1 ? idx('timestamp') : idx('date');
    const typeIdx = idx('type');
    const itemIdIdx = idx('itemid');
    const itemNameIdx = idx('itemname') > -1 ? idx('itemname') : idx('item');
    const qtyIdx = idx('quantity') > -1 ? idx('quantity') : idx('qty');
    const userIdx = idx('user');
    const reasonIdx = idx('reason');

    return rows
      .filter(r => Array.isArray(r) && r.some(v => String(v ?? '').trim() !== ''))
      .map(r => ({
        ID: idIdx > -1 ? r[idIdx] : '',
        Timestamp: tsIdx > -1 ? r[tsIdx] : '',
        Type: typeIdx > -1 ? r[typeIdx] : '',
        ItemId: itemIdIdx > -1 ? r[itemIdIdx] : '',
        ItemName: itemNameIdx > -1 ? r[itemNameIdx] : '',
        Quantity: qtyIdx > -1 ? r[qtyIdx] : '',
        User: userIdx > -1 ? r[userIdx] : '',
        Reason: reasonIdx > -1 ? r[reasonIdx] : '',
        _tz: tz
      }));
  }

  return rows
    .filter(r => Array.isArray(r) && r.some(v => String(v ?? '').trim() !== ''))
    .map(r => ({
      ID: r[0],
      Timestamp: r[1],
      Type: r[2],
      ItemId: r[3],
      ItemName: r[4],
      Quantity: r[5],
      User: r[6],
      Reason: r[7],
      _tz: tz
    }));
}

function findInventoryRowNumber_(headers, data, idValue) {
  // Check Serial first (most unique)
  const serialIndex = findHeaderIndex_(headers, ['Serial']);
  if (serialIndex > -1) {
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][serialIndex]) === String(idValue)) return i + 2;
    }
  }
  
  // Fallback: Check ID/Id if it exists (legacy)
  const idIndex = findHeaderIndex_(headers, ['ID', 'Id', 'id']);
  if (idIndex > -1) {
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][idIndex]) === String(idValue)) return i + 2;
    }
  }

  // Fallback: Check for ROW# format (e.g. "ROW#5")
  if (typeof idValue === 'string' && idValue.startsWith('ROW#')) {
    const rowNum = parseInt(idValue.replace('ROW#', ''), 10);
    if (!isNaN(rowNum) && rowNum > 1 && rowNum <= data.length + 1) { // +1 because data is 0-indexed relative to sheet start? No, data is array of values.
       // data is rows excluding header.
       // rowNum is 1-based sheet row index. Header is row 1.
       // So data[0] is row 2.
       // If rowNum is 2, index is 0.
       const index = rowNum - 2;
       if (index >= 0 && index < data.length) return rowNum;
    }
  }

  return null;
}

function getSystemLogs_(limit = 100) {
  const config = getSystemConfig();
  try {
    const sheet = getSheet(config.auditSheetName);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const startRow = Math.max(2, lastRow - limit + 1);
    const numRows = lastRow - startRow + 1;
    const values = sheet.getRange(startRow, 1, numRows, 5).getValues();
    
    return values.reverse().map(row => ({
      timestamp: row[0],
      action: row[1],
      details: row[2],
      level: row[3],
      user: row[4]
    }));
  } catch (error) {
    logAction('Get System Logs Error', error.message, 'error');
    return [];
  }
}

// --- Dashboard Stats ---
function getDashboardStats() {
  const config = getSystemConfig();
  const sheetName = getInventorySheetName_();
  const cacheKey = `dashboard_stats_${sheetName}`;
  const cached = getCachedData(cacheKey);
  if (cached) return cached;
  
  const dataSource = getDataSource('inventory');
  logAction('Get Dashboard Stats', `Data source: ${dataSource}`, 'debug');
  
  let inventory = [];
  if (dataSource === 'internal') {
    inventory = getInventory();
  } else if (dataSource === 'external' && config.externalApis.inventory) {
    try {
      const externalData = fetchExternalData(config.externalApis.inventory);
      inventory = externalData.items || [];
    } catch (error) {
      logAction('External Inventory Error', error.message, 'error');
      inventory = getInventory(); // Fallback to internal
    }
  } else {
    inventory = getInventory();
  }
  
  const totalItems = inventory.length;
  
  const lowStock = inventory.filter(i => {
    const qty = parseFloat(i.Qty);
    return !isNaN(qty) && qty > 0 && qty < 10;
  }).length;
  
  const outOfStock = inventory.filter(i => {
    const qty = parseFloat(i.Qty);
    return !isNaN(qty) && qty <= 0;
  }).length;
  
  const auditLogs = getSystemLogs_(50);
  const recentActivities = auditLogs.slice(-5).reverse();
  
  const transactions = getTransactions_();
  logAction('Dashboard Stats', `Found ${transactions.length} transactions`, 'debug');

  const tz = Session.getScriptTimeZone();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const start = new Date(today);
  start.setDate(start.getDate() - 6);

  const labels = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    labels.push(Utilities.formatDate(d, tz, 'EEE'));
  }

  const sales = new Array(7).fill(0);
  const restocks = new Array(7).fill(0);

  const toNum = v => {
    const n = Number(String(v ?? '').replace(/[^0-9.\-]+/g, ''));
    return Number.isFinite(n) ? n : 0;
  };

  const inRangeIndex = dt => {
    if (!dt) return -1;
    let d = dt;
    if (!(d instanceof Date)) {
      d = new Date(dt);
    }
    if (isNaN(d.getTime())) return -1;
    
    const copy = new Date(d);
    copy.setHours(0, 0, 0, 0);
    
    if (copy < start || copy > today) return -1;
    
    const diffTime = Math.abs(copy.getTime() - start.getTime());
    const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24)); 
    return diffDays >= 0 && diffDays < 7 ? diffDays : -1;
  };

  transactions.forEach(t => {
    const idx = inRangeIndex(t.Timestamp);
    if (idx < 0) return;

    const qty = Math.abs(toNum(t.Quantity));
    if (!qty) return;

    const type = String(t.Type ?? '').toLowerCase();
    if (type.includes('out')) {
      sales[idx] += qty;
    } else if (type.includes('in')) {
      restocks[idx] += qty;
    }
  });

  const trend = sales.map((s, i) => (s + restocks[i]) / 2);
  
  // Calculate Total Value
  const totalValue = inventory.reduce((sum, item) => {
    return sum + ((parseFloat(item.Qty) || 0) * (parseFloat(item.UnitCost) || 0));
  }, 0);

  const result = {
    totalItems,
    lowStock,
    outOfStock,
    totalValue,
    recentActivities,
    weeklyActivity: { labels, sales, restocks, trend }
  };
  
  setCachedData(cacheKey, result);
  return result;
}

// --- Products / Items Masterlist ---
function getInventory() {
  const sheetName = getInventorySheetName_();
  const sheet = getInventorySheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h ?? '').trim());
  const rows = data.slice(1);

  return rows.map((row, i) => {
    const raw = {};
    for (let j = 0; j < headers.length; j++) raw[headers[j]] = row[j];
    
    // Row number is i + 2 (1-based index, +1 for header)
    const rowId = `ROW#${i + 2}`;
    
    // Standardize known keys with robust alias matching
    const item = {
      ID: raw.ID || raw.Id || raw.id || raw['Property No.'] || raw['Property Number'] || raw.Serial || rowId,
      Project: raw.Project || raw['Project Name'] || raw['Project/Program'] || '',
      Category: raw.Category || raw['Classification'] || raw['Type'] || '',
      Item: raw.Item || raw['Item Name'] || raw['Item Description'] || raw.Description || raw.Article || '',
      BrandModel: raw.BrandModel || raw['Brand and model'] || raw['Brand and Model'] || raw.Brand || raw.Model || '',
      Serial: raw.Serial || raw['Serial No.'] || raw['Serial Number'] || '',
      Qty: raw.Qty || raw.Quantity || raw.Stock || raw.Count || 0,
      Unit: raw.Unit || raw.UOM || raw['Unit of Measure'] || '',
      UnitCost: raw.UnitCost || raw['Unit Cost'] || raw.Cost || raw.Price || raw['Unit Value'] || 0,
      DateAcquired: raw.DateAcquired || raw['Date Acquired'] || raw['Date Added'] || raw.Date || '',
      ProcurementProject:
        raw.ProcurementProject ||
        raw['Procurement/\nProject'] ||
        raw['Procurement/Project'] ||
        raw['Procurement Project'] ||
        raw.Procurement ||
        '',
      PersonInCharge: raw.PersonInCharge || raw['Person-in-charge'] || raw['Person in Charge'] || raw.Custodian || raw['End User'] || '',
      Location: raw.Location || raw.Office || raw['Whereabout'] || '',
      Status: raw.Status || raw.Condition || raw.State || raw.Remarks || '', // Fallback to Remarks if status is missing
      Remarks: raw.Remarks || raw.Comment || raw.Notes || ''
    };
    
    // Merge raw data so custom columns are available
    return { ...raw, ...item, _raw: raw };
  });
}

function addItem(item) {
  item = item && typeof item === 'object' ? item : {};
  if (item && item.item && typeof item.item === 'object') item = item.item;

  const sheetName = getInventorySheetName_();
  const sheet = getInventorySheet_();
  const headers = ensureInventorySheetColumns_(sheet);

  // Helper to prevent formula injection
  const sanitizeForSheet = (value) => {
    if (typeof value === 'string' && ['=', '+', '-', '@'].includes(value.charAt(0))) {
      return "'" + value;
    }
    return value;
  };

  const row = new Array(headers.length).fill('');
  
  // Dynamic mapping: Loop through all headers and find values in the item object
  headers.forEach((header, idx) => {
    // Try exact match first
    if (item && item[header] !== undefined) {
      row[idx] = sanitizeForSheet(item[header]);
    } else {
      // Try alias matching if not found
      for (const [key, aliases] of Object.entries(INVENTORY_COLUMN_VALIDATION)) {
        if (aliases.some(a => a.toLowerCase() === header.toLowerCase())) {
          if (item && item[key] !== undefined) {
            row[idx] = sanitizeForSheet(item[key]);
          }
          break;
        }
      }
    }
  });

  const hasAnyValue = row.some(v => v !== '' && v !== null && v !== undefined);
  if (!hasAnyValue) return { error: 'Missing item payload' };

  // Handle Status default if empty
  const statusIdx = findHeaderIndex_(headers, ['Status']);
  if (statusIdx > -1 && !row[statusIdx]) {
    const qtyIdx = findHeaderIndex_(headers, ['Qty']);
    const qty = qtyIdx > -1 ? Number(row[qtyIdx]) : 0;
    row[statusIdx] = qty > 0 ? 'Good' : 'Out of Stock';
  }

  sheet.appendRow(row);
  
  logAction('Add Item', `Added item to sheet: ${sheetName}`);
  return { success: true };
}

function editItem(id, updatedItem) {
  const sheetName = getInventorySheetName_();
  const sheet = getInventorySheet_();
  const headers = ensureInventorySheetColumns_(sheet);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Item not found');

  const updates = Object.assign({}, updatedItem);
  delete updates.id;
  delete updates.ID;

  // Helper to prevent formula injection
  const sanitizeForSheet = (value) => {
    if (typeof value === 'string' && ['=', '+', '-', '@'].includes(value.charAt(0))) {
      return "'" + value;
    }
    return value;
  };

  // Find row
  const rows = data.slice(1);
  const foundRowNumber = findInventoryRowNumber_(headers, rows, id);
  
  if (foundRowNumber) {
    // Iterate through all updates and find the corresponding column index
    for (const [key, value] of Object.entries(updates)) {
      // Find the index of this header or its alias
      let colIdx = headers.indexOf(key);
      if (colIdx === -1) {
        // Try alias matching
        for (const [stdKey, aliases] of Object.entries(INVENTORY_COLUMN_VALIDATION)) {
          if (stdKey === key) {
            // Find which header in the sheet matches one of these aliases
            colIdx = headers.findIndex(h => aliases.some(a => a.toLowerCase() === h.toLowerCase()));
            break;
          }
        }
      }

      if (colIdx > -1) {
        sheet.getRange(foundRowNumber, colIdx + 1).setValue(sanitizeForSheet(value));
      }
    }

    logAction('Edit Item', `Edited item ID: ${id} in sheet: ${sheetName}`);
    return { success: true };
  }
  
  throw new Error('Item not found');
}

function deleteItem(id) {
  const sheetName = getInventorySheetName_();
  const sheet = getInventorySheet_();
  const headers = ensureInventorySheetColumns_(sheet);
  const data = sheet.getDataRange().getValues();
  
  const rows = data.slice(1);
  const rowNumber = findInventoryRowNumber_(headers, rows, id);
  
  if (rowNumber) {
    sheet.deleteRow(rowNumber);
    logAction('Delete Item', `Deleted item ID: ${id}`);
    return { success: true };
  }
  
  throw new Error('Item not found');
}

function adjustStock(itemId, quantityChange, reason) {
  const sheetName = getInventorySheetName_();
  const sheet = getInventorySheet_();
  const headers = ensureInventorySheetColumns_(sheet);
  const data = sheet.getDataRange().getValues();
  const qtyIndex = findHeaderIndex_(headers, ['Qty']);
  const statusIndex = findHeaderIndex_(headers, ['Status']);
  const nameIndex = findHeaderIndex_(headers, ['Item']);
  if (qtyIndex === -1 || statusIndex === -1 || nameIndex === -1) {
    throw new Error('Missing required Inventory columns');
  }

  // Validate numeric input
  const qtyChange = parseInt(quantityChange, 10);
  if (isNaN(qtyChange)) {
    throw new Error('Invalid quantity change value');
  }

  const rowNumber = findInventoryRowNumber_(headers, data.slice(1), itemId);
  if (rowNumber) {
      const i = rowNumber - 2; // Correct index for data slice
      const rowData = data[rowNumber - 1];
      const currentQtyVal = parseInt(rowData[qtyIndex] || 0);
      const newQty = currentQtyVal + qtyChange;
      
      sheet.getRange(rowNumber, qtyIndex + 1).setValue(newQty);

      if (newQty <= 0) {
         sheet.getRange(rowNumber, statusIndex + 1).setValue('Out of Stock');
      } else if (rowData[statusIndex] === 'Out of Stock') {
         sheet.getRange(rowNumber, statusIndex + 1).setValue('Good');
      }

      // Record Transaction
      const transSheet = getSheet('Transactions');
      transSheet.appendRow([
        Utilities.getUuid(), 
        new Date(), 
        qtyChange > 0 ? 'Stock In' : 'Stock Out', 
        itemId, 
        rowData[nameIndex],
        Math.abs(qtyChange), 
        Session.getActiveUser().getEmail(), 
        reason
      ]);
      
      logAction('Adjust Stock', `Adjusted stock for ${rowData[nameIndex]} by ${qtyChange}. Reason: ${reason}`);
      return { success: true, newQty: newQty };
  }
  throw new Error('Item not found');
}

// --- Purchasing ---
function getSuppliers() {
  return getDataFromSheet('Suppliers');
}

function addSupplier(supplier) {
  const sheet = getSheet('Suppliers');
  sheet.appendRow([Utilities.getUuid(), supplier.name, supplier.contact, supplier.email, supplier.address]);
  logAction('Add Supplier', `Added supplier: ${supplier.name}`);
  return { success: true };
}

// --- Reports & Logs ---
function logAction(action, details) {
  const sheet = getSheet('AuditLogs');
  // Ensure headers if empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'User', 'Action', 'Details']);
  }
  sheet.appendRow([new Date(), Session.getActiveUser().getEmail(), action, details]);
}

function getAuditLogs() {
  const sheet = getSheet('AuditLogs');
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return [];

  // Check if first row looks like a header
  const firstRow = data[0];
  const hasHeader = typeof firstRow[0] === 'string' && (firstRow[0] === 'Timestamp' || firstRow[0] === 'Date');
  
  const rows = hasHeader ? data.slice(1) : data;
  
  return rows.map(row => ({
    Timestamp: row[0],
    User: row[1],
    Action: row[2],
    Details: row[3]
  }));
}

// --- Notifications ---

function setupDisposalTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const triggerName = 'checkDisposalNotifications';
  
  // Delete existing triggers for this function to avoid duplicates/conflicts
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === triggerName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create Daily Trigger (Every Day at 8 AM)
  ScriptApp.newTrigger(triggerName)
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
      
  return `Notification trigger set to Daily (Every Day at 8 AM).`;
}

function checkDisposalNotifications() {
  const inventory = getInventory();
  const now = new Date();
  now.setHours(0, 0, 0, 0);

  // Get already notified items
  const props = PropertiesService.getScriptProperties();
  let notifiedItems = [];
  try {
    const stored = props.getProperty('NOTIFIED_DISPOSAL_ITEMS');
    if (stored) notifiedItems = JSON.parse(stored);
  } catch (e) {
    console.error('Error parsing notified items:', e);
  }

  const expiringItems = [];
  const expiredItems = [];
  const newNotifiedSerials = [];

  inventory.forEach(item => {
    let dateStr = String(item.DateAcquired || '').trim();
    if (!dateStr) return;

    let acquiredDate;
    // Handle Year-only input (e.g. "2025")
    if (/^\d{4}$/.test(dateStr)) {
      acquiredDate = new Date(`${dateStr}-01-01`);
    } else {
      acquiredDate = new Date(dateStr);
    }

    if (isNaN(acquiredDate.getTime())) return;

    // 5 Years Disposal Logic
    const disposalDate = new Date(acquiredDate);
    disposalDate.setFullYear(disposalDate.getFullYear() + 5);
    disposalDate.setHours(0, 0, 0, 0);

    const diffTime = disposalDate - now;
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    // Unique Identifier (Serial or ID)
    const itemId = item.Serial || item.ID;

    // Skip if already notified
    if (notifiedItems.includes(itemId)) {
      // Keep tracking it so we don't lose it if we re-save
      newNotifiedSerials.push(itemId);
      return;
    }

    if (diffDays < 0) {
      expiredItems.push({ ...item, disposalDate, daysOverdue: Math.abs(diffDays) });
      newNotifiedSerials.push(itemId);
    } else if (diffDays <= 90) {
      expiringItems.push({ ...item, disposalDate, daysLeft: diffDays });
      // Note: We might NOT want to suppress "Expiring Soon" updates if days change?
      // But user said "Once detected... dun lang".
      // Let's suppress repeat "Expiring Soon" too.
      newNotifiedSerials.push(itemId);
    }
  });

  if (expiringItems.length === 0 && expiredItems.length === 0) {
    // Cleanup: Update notified list to only include items that are still relevant (optional, but good for size limits)
    // Actually, simple cleanup: just save the new list if it changed?
    // Risk: If we only save newNotifiedSerials, we might drop items that were skipped above?
    // Logic above: `if (notifiedItems.includes(itemId)) { newNotifiedSerials.push(itemId); return; }`
    // So newNotifiedSerials contains ALL relevant notified items found in current inventory.
    // We should update properties to keep it clean (remove deleted items).
    if (newNotifiedSerials.length !== notifiedItems.length) {
       props.setProperty('NOTIFIED_DISPOSAL_ITEMS', JSON.stringify(newNotifiedSerials));
    }
    return { message: 'No new items require disposal notification.' };
  }

  const sentTo = sendDisposalEmail(expiredItems, expiringItems);
  
  if (sentTo) {
    // Save updated notified list
    props.setProperty('NOTIFIED_DISPOSAL_ITEMS', JSON.stringify(newNotifiedSerials));
    
    return { 
      success: true, 
      message: `Email sent to ${sentTo}. New Expired: ${expiredItems.length}, New Expiring Soon: ${expiringItems.length}` 
    };
  } else {
    return {
      success: false,
      message: `Failed to send email. Could not determine recipient email address.`
    };
  }
}

function sendDisposalEmail(expired, expiring) {
  let recipient = Session.getActiveUser().getEmail(); 
  if (!recipient) {
    recipient = Session.getEffectiveUser().getEmail();
  }

  if (!recipient) {
    Logger.log('No active user email found. Please configure a recipient.');
    return null;
  }

  const subject = `Inventory Disposal Alert - ${new Date().toLocaleDateString()}`;
  
  let htmlBody = '<h2>Inventory Disposal Notification</h2>';
  
  if (expired.length > 0) {
    htmlBody += '<h3 style="color: #dc2626;">Expired Items (Needs Disposal)</h3>';
    htmlBody += createEmailTable(expired, 'expired');
  }

  if (expiring.length > 0) {
    htmlBody += '<h3 style="color: #d97706;">Expiring Soon (Within 90 Days)</h3>';
    htmlBody += createEmailTable(expiring, 'expiring');
  }

  htmlBody += '<p>Please check the inventory system for more details.</p>';

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody
  });
  
  return recipient;
}

function createEmailTable(items, type) {
  let html = '<table style="border-collapse: collapse; width: 100%; border: 1px solid #ddd;">';
  html += '<tr style="background-color: #f3f4f6;">';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">Item</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">Serial</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">Location</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">Disposal Date</th>';
  html += `<th style="border: 1px solid #ddd; padding: 8px;">${type === 'expired' ? 'Overdue By' : 'Days Left'}</th>`;
  html += '</tr>';

  items.forEach(item => {
    const dateStr = item.disposalDate.toLocaleDateString();
    const daysInfo = type === 'expired' ? `${item.daysOverdue} days` : `${item.daysLeft} days`;
    
    html += '<tr>';
    html += `<td style="border: 1px solid #ddd; padding: 8px;">${item.Item}</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px;">${item.Serial}</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px;">${item.Location}</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px;">${dateStr}</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px;">${daysInfo}</td>`;
    html += '</tr>';
  });

  html += '</table>';
  return html;
}
