// Global Config
const API_URL = "https://script.google.com/macros/s/AKfycby8fW5VWaXu5tgDUe76LwaWJb8L-zTcTyqwtiSADM_f_-JRbQwnKnYmEYACV6Oj3SGm/exec";

// --- API Helper ---
async function callApi(action, data = null) {
  try {
    let response;
    if (data) {
      // POST request
      // Use text/plain to avoid CORS preflight, handled manually in GAS doPost
      response = await fetch(`${API_URL}?action=${action}`, {
        method: 'POST',
        body: JSON.stringify(data),
        headers: {
            'Content-Type': 'text/plain;charset=utf-8' 
        }
      });
    } else {
      // GET request
      response = await fetch(`${API_URL}?action=${action}`);
    }
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    return await response.json();
  } catch (error) {
    console.error('API Error:', error);
    alert('API Error: ' + error.message);
    throw error;
  }
}

// --- Navigation & Loading ---
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
      'products': 'Products Masterlist',
      'inventory': 'Inventory Management',
      'reports': 'Reports',
      'audit': 'Audit Logs'
    };
    document.getElementById('page-title').innerText = titles[sectionId];
    
    // Load Content dynamically if not already loaded (simple cache)
    // Note: Since we are fetching local HTML files, we assume they are in the same directory.
    // If we want to reload fresh data every time, we call the data loader.
    // But the HTML structure only needs to be loaded once.
    if (!wrapper.hasAttribute('data-loaded')) {
        try {
            const resp = await fetch(`${sectionId}.html`);
            if(!resp.ok) throw new Error('Failed to load template');
            const html = await resp.text();
            wrapper.innerHTML = html;
            wrapper.setAttribute('data-loaded', 'true');
        } catch(e) {
            wrapper.innerHTML = `<p class="text-red-500">Error loading module: ${e.message}</p>`;
            return;
        }
    }
    
    // Close mobile menu on navigation
    if (window.innerWidth < 1024) {
      const sidebar = document.getElementById('sidebar');
      if (sidebar) {
        sidebar.classList.add('-translate-x-full');
      }
    }
    
    // Trigger Data Load
    if(sectionId === 'dashboard') loadDashboard();
    if(sectionId === 'products') loadProducts();
    if(sectionId === 'inventory') loadInventory();
    if(sectionId === 'audit') loadAuditLogs();
}

// --- Dashboard Logic ---
async function loadDashboard() {
  const stats = await callApi('getDashboardStats');
  if(stats.error) return; // Handled in callApi or show alert
  renderDashboard(stats);
}

function renderDashboard(stats) {
  document.getElementById('dash-total-items').innerText = stats.totalItems;
  document.getElementById('dash-low-stock').innerText = stats.lowStock;
  document.getElementById('dash-out-stock').innerText = stats.outOfStock;
  document.getElementById('dash-total-value').innerText = '$' + stats.totalValue.toLocaleString();
  document.getElementById('dash-notification-count').innerText = stats.lowStock;
  
  const activityList = document.getElementById('recent-activities-list');
  if(activityList) {
      activityList.innerHTML = '';
      if(stats.recentActivities.length === 0) {
         activityList.innerHTML = '<li class="py-3 text-sm text-gray-500">No recent activity.</li>';
      } else {
         stats.recentActivities.forEach(act => {
           const li = document.createElement('li');
           li.className = 'py-3';
           li.innerHTML = `
             <div class="flex space-x-3">
               <div class="flex-1 space-y-1">
                 <div class="flex items-center justify-between">
                   <h3 class="text-sm font-medium text-gray-900">${act.Type} - ${act.ItemName}</h3>
                   <p class="text-sm text-gray-500">${new Date(act.Date).toLocaleDateString()}</p>
                 </div>
                 <p class="text-sm text-gray-500">${act.Notes} (Qty: ${act.Quantity}) by ${act.User}</p>
               </div>
             </div>
           `;
           activityList.appendChild(li);
         });
      }
  }
  
  // Render Chart
  renderChart();
}

function renderChart() {
    const ctx = document.getElementById('mainChart');
    if(!ctx) return;
    
    // Destroy existing if needed (simple check)
    // In a real app, track the chart instance
    if(window.myChart instanceof Chart) {
        window.myChart.destroy();
    }
    
    window.myChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
            datasets: [{
                label: 'Sales',
                data: [12, 19, 3, 5, 2, 3, 10], // Mock data
                backgroundColor: 'rgba(54, 162, 235, 0.2)',
                borderColor: 'rgba(54, 162, 235, 1)',
                borderWidth: 1
            }, {
                label: 'Restocks',
                data: [2, 3, 20, 5, 1, 4, 2], // Mock data
                backgroundColor: 'rgba(75, 192, 192, 0.2)',
                borderColor: 'rgba(75, 192, 192, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false
        }
    });
}

// --- Products Logic ---
async function loadProducts() {
  const items = await callApi('getInventory');
  if(items.error) return;
  renderProducts(items);
}

function renderProducts(items) {
  const tbody = document.getElementById('products-table-body');
  if(!tbody) return;
  tbody.innerHTML = '';
  items.forEach(item => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${item.Item}</td>
      <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${item.BrandModel || '-'}</td>
      <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${item.Serial || '-'}</td>
      <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${item.Category || '-'}</td>
      <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${item.Location || '-'}</td>
      <td class="px-6 py-4 whitespace-nowrap">
        <span class="px-2 inline-flex text-xs leading-5 font-semibold rounded-full ${getStatusColor(item.Status)}">
          ${item.Status}
        </span>
      </td>
      <td class="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
        <button onclick="deleteItem('${item.ID}')" class="text-red-600 hover:text-red-900">Delete</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function getStatusColor(status) {
    switch(status) {
        case 'Good': return 'bg-green-100 text-green-800';
        case 'Damaged': return 'bg-red-100 text-red-800';
        case 'For Repair': return 'bg-yellow-100 text-yellow-800';
        case 'Lost': return 'bg-gray-100 text-gray-800';
        default: return 'bg-blue-100 text-blue-800';
    }
}

async function submitNewItem() {
   const item = {
     Project: document.getElementById('new-project').value,
     Category: document.getElementById('new-category').value,
     Item: document.getElementById('new-item').value,
     BrandModel: document.getElementById('new-brandmodel').value,
     Serial: document.getElementById('new-serial').value,
     Qty: parseInt(document.getElementById('new-qty').value) || 0,
     Unit: document.getElementById('new-unit').value,
     UnitCost: parseFloat(document.getElementById('new-unitcost').value) || 0,
     DateAcquired: document.getElementById('new-dateacquired').value,
     ProcurementProject: document.getElementById('new-procurementproject').value,
     PersonInCharge: document.getElementById('new-personincharge').value,
     Location: document.getElementById('new-location').value,
     Status: document.getElementById('new-status').value,
     Remarks: document.getElementById('new-remarks').value
   };
   
   const res = await callApi('addItem', item);
   if(res.success) {
     closeModal('addProductModal');
     loadProducts();
     alert('Item added successfully');
   }
}

async function deleteItem(id) {
  if(confirm('Are you sure you want to delete this item?')) {
    const res = await callApi('deleteItem', { id: id });
    if(res.success) loadProducts();
  }
}

// --- Inventory Logic ---
async function loadInventory() {
  const items = await callApi('getInventory');
  renderInventory(items);
}

function renderInventory(items) {
  const tbody = document.getElementById('inventory-table-body');
  if(!tbody) return;
  tbody.innerHTML = '';
  items.forEach(item => {
     const tr = document.createElement('tr');
     tr.innerHTML = `
       <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${item.Item}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${item.BrandModel || ''}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900 font-bold">${item.Qty}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${item.Unit || ''}</td>
       <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${item.Status}</td>
       <td class="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
         <button onclick="openAdjustStockModal('${item.ID}', '${item.Item}')" class="text-indigo-600 hover:text-indigo-900">Adjust</button>
       </td>
     `;
     tbody.appendChild(tr);
  });
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
  }
}

// --- Audit Logic ---
async function loadAuditLogs() {
  const logs = await callApi('getAuditLogs');
  renderAuditLogs(logs);
}

function renderAuditLogs(logs) {
  const tbody = document.getElementById('audit-table-body');
  if(!tbody) return;
  tbody.innerHTML = '';
  logs.reverse().forEach(log => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${new Date(log.Timestamp).toLocaleString()}</td>
      <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">${log.User}</td>
      <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900 font-medium">${log.Action}</td>
      <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${log.Details}</td>
    `;
    tbody.appendChild(tr);
  });
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

// Close mobile menu when navigating
// Mobile menu functionality added to the existing loadSection function

// Initialize mobile menu on page load
document.addEventListener('DOMContentLoaded', function() {
  initMobileMenu();
});
