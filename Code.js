// --- API Configuration ---
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : null;
  
  // Dispatcher
  let result = {};
  try {
    if (action === 'getDashboardStats') {
      result = getDashboardStats();
    } else if (action === 'getInventory') {
      result = getInventory();
    } else if (action === 'getSuppliers') {
      result = getSuppliers();
    } else if (action === 'getAuditLogs') {
      result = getAuditLogs();
    } else if (action === 'checkDisposalNotifications') {
      result = checkDisposalNotifications();
    } else {
      result = { error: action ? 'Invalid action' : 'Missing action' };
    }
  } catch (err) {
    result = { error: err.message };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
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
  
  const action = e && e.parameter ? e.parameter.action : null;
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
    } else {
      result = { error: action ? 'Invalid action' : 'Missing action' };
    }
  } catch (err) {
    result = { error: err.message };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
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
    lowStockItems.length ? `Low stock (1-10) sample:\n${lowStockItems.join('\n')}` : 'Low stock: none',
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

function ensureInventorySheetColumns_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    sheet.appendRow([
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
    ]);
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h ?? '').trim());
  }

  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h ?? '').trim());
  return headers;
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

// --- Dashboard Stats ---
function getDashboardStats() {
  const inventory = getInventory();
  
  const totalItems = inventory.length;
  
  const lowStock = inventory.filter(i => {
    const qty = parseFloat(i.Qty);
    return !isNaN(qty) && qty > 0 && qty < 10;
  }).length;
  
  const outOfStock = inventory.filter(i => {
    const qty = parseFloat(i.Qty);
    return !isNaN(qty) && qty <= 0;
  }).length;
  
  const auditLogs = getAuditLogs();
  const recentActivities = auditLogs.slice(-5).reverse();
  
  // Calculate Total Value
  const totalValue = inventory.reduce((sum, item) => {
    return sum + ((parseFloat(item.Qty) || 0) * (parseFloat(item.UnitCost) || 0));
  }, 0);

  return {
    totalItems,
    lowStock,
    outOfStock,
    totalValue,
    recentActivities
  };
}

// --- Products / Items Masterlist ---
function getInventory() {
  const sheet = getSheet('Inventory');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h ?? '').trim());
  const rows = data.slice(1);

  return rows.map((row, i) => {
    const raw = {};
    for (let j = 0; j < headers.length; j++) raw[headers[j]] = row[j];
    
    // Row number is i + 2 (1-based index, +1 for header)
    const rowId = `ROW#${i + 2}`;
    
    return {
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
  });
}

function addItem(item) {
  const sheet = getSheet('Inventory');
  const headers = ensureInventorySheetColumns_(sheet);

  let status = item.Status;
  if (!status) status = item.Qty > 0 ? 'Good' : 'Out of Stock';

  const row = new Array(headers.length).fill('');
  const set = (candidates, value) => {
    const idx = findHeaderIndex_(headers, candidates);
    if (idx > -1) row[idx] = value;
  };

  set(['Project', 'Project Name'], item.Project);
  set(['Category'], item.Category);
  set(['Item'], item.Item);
  set(['BrandModel', 'Brand and model', 'Brand and Model'], item.BrandModel);
  set(['Serial'], item.Serial);
  set(['Qty'], item.Qty);
  set(['Unit'], item.Unit);
  set(['UnitCost', 'Unit Cost'], item.UnitCost);
  set(['DateAcquired', 'Date Acquired'], item.DateAcquired);
  set(['ProcurementProject', 'Procurement/\nProject', 'Procurement/Project', 'Procurement Project', 'Procurement'], item.ProcurementProject);
  set(['PersonInCharge', 'Person-in-charge', 'Person in Charge'], item.PersonInCharge);
  set(['Location'], item.Location);
  set(['Status'], status);
  set(['Remarks'], item.Remarks);

  sheet.appendRow(row);
  
  logAction('Add Item', `Added item: ${item.Item} (Serial: ${item.Serial})`);
  return { success: true };
}

function editItem(id, updatedItem) {
  const sheet = getSheet('Inventory');
  const headers = ensureInventorySheetColumns_(sheet);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Item not found');

  const updates = Object.assign({}, updatedItem);
  delete updates.id;
  delete updates.ID;

  // Find row
  const rowNumber = findInventoryRowNumber_(headers, data.slice(1), id); // Pass data excluding header for helper? 
  // Wait, findInventoryRowNumber_ expects data as rows array?
  // Let's check findInventoryRowNumber_ implementation.
  // It expects `data` to be the array of rows (without header? or with?).
  // Implementation:
  // if (String(data[i][serialIndex]) === String(idValue)) return i + 2;
  // This implies `data` passed to it should be the rows excluding header.
  // And it returns i + 2, meaning row index 1-based.
  
  const rows = data.slice(1);
  const foundRowNumber = findInventoryRowNumber_(headers, rows, id);
  
  if (foundRowNumber) {
    const row = rows[foundRowNumber - 2]; // Get the actual row array
    
    // Helper to update cell if key exists in updates
    const updateCell = (candidates, value) => {
      const idx = findHeaderIndex_(headers, candidates);
      if (idx > -1) {
         sheet.getRange(foundRowNumber, idx + 1).setValue(value);
      }
    };

    if (updates.Project !== undefined) updateCell(['Project', 'Project Name'], updates.Project);
    if (updates.Category !== undefined) updateCell(['Category'], updates.Category);
    if (updates.Item !== undefined) updateCell(['Item'], updates.Item);
    if (updates.BrandModel !== undefined) updateCell(['BrandModel', 'Brand and model', 'Brand and Model'], updates.BrandModel);
    if (updates.Serial !== undefined) updateCell(['Serial'], updates.Serial);
    if (updates.Qty !== undefined) updateCell(['Qty'], updates.Qty);
    if (updates.Unit !== undefined) updateCell(['Unit'], updates.Unit);
    if (updates.UnitCost !== undefined) updateCell(['UnitCost', 'Unit Cost'], updates.UnitCost);
    if (updates.DateAcquired !== undefined) updateCell(['DateAcquired', 'Date Acquired'], updates.DateAcquired);
    if (updates.ProcurementProject !== undefined) updateCell(['ProcurementProject', 'Procurement/\nProject', 'Procurement/Project', 'Procurement Project', 'Procurement'], updates.ProcurementProject);
    if (updates.PersonInCharge !== undefined) updateCell(['PersonInCharge', 'Person-in-charge', 'Person in Charge'], updates.PersonInCharge);
    if (updates.Location !== undefined) updateCell(['Location'], updates.Location);
    if (updates.Status !== undefined) updateCell(['Status'], updates.Status);
    if (updates.Remarks !== undefined) updateCell(['Remarks'], updates.Remarks);

    logAction('Edit Item', `Edited item: ${updatedItem.Item || id}`);
    return { success: true };
  }
  
  throw new Error('Item not found');
}

function deleteItem(id) {
  const sheet = getSheet('Inventory');
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
  const sheet = getSheet('Inventory');
  const headers = ensureInventorySheetColumns_(sheet);
  const data = sheet.getDataRange().getValues();
  const qtyIndex = findHeaderIndex_(headers, ['Qty']);
  const statusIndex = findHeaderIndex_(headers, ['Status']);
  const nameIndex = findHeaderIndex_(headers, ['Item']);
  if (qtyIndex === -1 || statusIndex === -1 || nameIndex === -1) {
    throw new Error('Missing required Inventory columns');
  }

  const rowNumber = findInventoryRowNumber_(headers, data.slice(1), itemId);
  if (rowNumber) {
      const i = rowNumber - 2; // Correct index for data slice
      const currentQty = parseInt(data[i + 1][qtyIndex] || 0); // +1 because data includes header? No data slice? 
      // Wait, data is from getDataRange().getValues(), so data[0] is header.
      // rowNumber is 1-based.
      // so rowNumber 2 is index 1 in data.
      
      const rowData = data[rowNumber - 1];
      const currentQtyVal = parseInt(rowData[qtyIndex] || 0);
      const newQty = currentQtyVal + parseInt(quantityChange);
      
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
        quantityChange > 0 ? 'Stock In' : 'Stock Out', 
        itemId, 
        rowData[nameIndex],
        Math.abs(quantityChange), 
        Session.getActiveUser().getEmail(), 
        reason
      ]);
      
      logAction('Adjust Stock', `Adjusted stock for ${rowData[nameIndex]} by ${quantityChange}. Reason: ${reason}`);
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
function getAuditLogs() {
  return getDataFromSheet('AuditLogs');
}

function logAction(action, details) {
  const sheet = getSheet('AuditLogs');
  sheet.appendRow([new Date(), Session.getActiveUser().getEmail(), action, details]);
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
