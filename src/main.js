import { createClient } from '@supabase/supabase-js';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';
import initialData from './initialData.json';

// ==========================================================================
// CONFIGURATION & DATABASE SETUP
// ==========================================================================
const SUPABASE_URL = 'https://eqdccbnlimuvchkuhjpe.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVxZGNjYm5saW11dmNoa3VoanBlIiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTA3ODMxMDcsImV4cCI6MjEwNjM1OTEwN30.sX7Au7chhG1Sgu9RvOP2G2PVh4aiOYfgPdaRqRPEJZI';

// Clean legacy localStorage keys so they do not override built-in config
localStorage.removeItem('v2_supabase_url');
localStorage.removeItem('v2_supabase_key');

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// ==========================================================================
// APPLICATION STATE
// ==========================================================================
let allOrders = [];
let filteredOrders = [];
let pendingUploadRows = [];
let currentStore = 'WISEROUTLET';
let currentTimeFilter = 'all';
let currentStatusFilter = 'all';
let searchQuery = '';
let customDateFrom = null;
let customDateTo = null;
let sortColumn = 'purchase_date';
let sortDirection = 'desc';
let currentPage = 1;
let rowsPerPage = 15;
let timelineViewMode = 'daily'; // 'daily' or 'cumulative'

let timelineChart = null;
let breakdownChart = null;

// ==========================================================================
// DOM ELEMENT REFERENCES
// ==========================================================================
const storeSelector = document.getElementById('store-selector');
const brandLogoIcon = document.getElementById('brand-logo-icon');
const dbStatusBtn = document.getElementById('db-status-btn');
const dbStatusDot = document.getElementById('db-status-dot');
const dbStatusText = document.getElementById('db-status-text');

// Sync Strip
const lastUpdatedIst = document.getElementById('last-updated-ist');
const lastUpdatedEst = document.getElementById('last-updated-est');
const totalDbCount = document.getElementById('total-db-count');
const revertBtn = document.getElementById('revert-btn');
const refreshDbBtn = document.getElementById('refresh-db-btn');

// Search & Filters
const searchInput = document.getElementById('search-input');
const searchClearBtn = document.getElementById('search-clear-btn');
const searchSuggestions = document.getElementById('search-suggestions');
const dateFromInput = document.getElementById('date-from');
const dateToInput = document.getElementById('date-to');
const applyDatesBtn = document.getElementById('apply-dates-btn');
const resetDatesBtn = document.getElementById('reset-dates-btn');
const timeFilterPills = document.getElementById('time-filter-pills');
const statusFilterPills = document.getElementById('status-filter-pills');

// KPI Cards
const kpiTotalProfit = document.getElementById('kpi-total-profit');
const kpiProfitMargin = document.getElementById('kpi-profit-margin');
const kpiProfitCard = document.getElementById('kpi-profit-card');
const kpiTotalRevenue = document.getElementById('kpi-total-revenue');
const kpiAov = document.getElementById('kpi-aov');
const kpiTotalOrders = document.getElementById('kpi-total-orders');
const kpiShippedRatio = document.getElementById('kpi-shipped-ratio');
const kpiOrderBreakdown = document.getElementById('kpi-order-breakdown');
const kpiTotalShipping = document.getElementById('kpi-total-shipping');
const kpiAvgShipping = document.getElementById('kpi-avg-shipping');
const kpiTotalTax = document.getElementById('kpi-total-tax');
const kpiTotalCogs = document.getElementById('kpi-total-cogs');

// Charts
const timelineCanvas = document.getElementById('timelineChart');
const breakdownCanvas = document.getElementById('breakdownChart');
const chartViewDaily = document.getElementById('chart-view-daily');
const chartViewCumulative = document.getElementById('chart-view-cumulative');

// Table
const ordersTbody = document.getElementById('orders-tbody');
const tableShowingText = document.getElementById('table-showing-text');
const rowsPerPageSelect = document.getElementById('rows-per-page');
const paginationInfo = document.getElementById('pagination-info');
const paginationButtons = document.getElementById('pagination-buttons');
const exportExcelBtn = document.getElementById('export-excel-btn');

// Upload Modal
const openUploadModalBtn = document.getElementById('open-upload-modal-btn');
const uploadModal = document.getElementById('upload-modal');
const closeUploadModalBtn = document.getElementById('close-upload-modal-btn');
const cancelUploadBtn = document.getElementById('cancel-upload-btn');
const confirmUploadBtn = document.getElementById('confirm-upload-btn');
const dropzone = document.getElementById('dropzone');
const masterFileInput = document.getElementById('master-file-input');
const uploadStoreSelect = document.getElementById('upload-store-select');
const uploadPreviewSection = document.getElementById('upload-preview-section');
const previewFilename = document.getElementById('preview-filename');
const previewTotalRows = document.getElementById('preview-total-rows');
const previewShippedRows = document.getElementById('preview-shipped-rows');
const previewTotalProfit = document.getElementById('preview-total-profit');
const uploadProgressWrap = document.getElementById('upload-progress-wrap');
const uploadProgressBar = document.getElementById('upload-progress-bar');
const uploadProgressText = document.getElementById('upload-progress-text');
const uploadProgressPct = document.getElementById('upload-progress-pct');

// Settings Modal
const openSettingsBtn = document.getElementById('open-settings-btn');
const settingsModal = document.getElementById('settings-modal');
const closeSettingsModalBtn = document.getElementById('close-settings-modal-btn');
const cancelSettingsBtn = document.getElementById('cancel-settings-btn');
const clearCacheBtn = document.getElementById('clear-cache-btn');
const toastContainer = document.getElementById('toast-container');

// ==========================================================================
// TOAST NOTIFICATIONS
// ==========================================================================
function showToast(title, desc, type = 'info', duration = 4000) {
  const toast = document.createElement('div');
  toast.className = `status-toast toast-${type}`;
  
  const icon = type === 'success' ? '✓' : type === 'error' ? '✕' : 'ℹ';
  toast.innerHTML = `
    <span class="toast-icon">${icon}</span>
    <div class="toast-content">
      <div class="toast-title">${title}</div>
      <div class="toast-desc">${desc}</div>
    </div>
  `;
  
  toastContainer.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

// ==========================================================================
// LOCAL STORAGE CACHE HELPERS
// ==========================================================================
const CACHE_KEY = 'v2_orders_cache_v3';
const SYNC_TIME_KEY = 'v2_last_sync_timestamp';
const UPLOAD_KEYS_KEY = 'v2_last_upload_keys';

// Invalidate legacy cache versions
try {
  localStorage.removeItem('v2_orders_cache_v2');
  localStorage.removeItem('v2_orders_cache');
} catch (e) {}

function saveOrdersToCache(orders) {
  try {
    localStorage.setItem(CACHE_KEY, JSON.stringify(orders));
  } catch (e) {
    console.warn('LocalStorage quota exceeded or unavailable:', e);
  }
}

function loadOrdersFromCache() {
  try {
    const cached = localStorage.getItem(CACHE_KEY);
    if (cached) {
      const parsed = JSON.parse(cached);
      if (Array.isArray(parsed) && parsed.length > 0) return normalizeOrders(parsed, { isRawUpload: false });
    }
    return Array.isArray(initialData) ? normalizeOrders(initialData, { isRawUpload: false }) : [];
  } catch (e) {
    console.error('Error loading cache:', e);
    return Array.isArray(initialData) ? normalizeOrders(initialData, { isRawUpload: false }) : [];
  }
}

// ==========================================================================
// SUPABASE DATA FETCH & SYNC
// ==========================================================================
async function fetchOrdersFromDatabase() {
  updateDbStatus('connecting', 'Connecting…');

  try {
    const { data, error } = await supabase
      .from('orders')
      .select('*')
      .order('purchase_date', { ascending: false });

    if (error) throw error;

    if (data && Array.isArray(data)) {
      allOrders = normalizeOrders(data, { isRawUpload: false });
      saveOrdersToCache(allOrders);
      updateDbStatus('online', 'Cloud Connected');
      renderLastSyncTime(new Date().toISOString());
      localStorage.setItem(SYNC_TIME_KEY, new Date().toISOString());
      applyFilters();
      return true;
    }
  } catch (err) {
    console.warn('Could not fetch from database (using local cache):', err.message);
    updateDbStatus('offline', 'Local Cache Active');
    
    // Load cached fallback
    const cached = loadOrdersFromCache();
    if (cached.length > 0) {
      allOrders = cached;
      applyFilters();
      showToast('Offline / Cache Mode', `Loaded ${cached.length} records from local cache.`, 'info');
    } else {
      showToast('Database Notice', 'No orders in database or local cache. Upload Master_Orders.xlsx to get started.', 'info');
    }
    return false;
  }
}

function updateDbStatus(status, label) {
  dbStatusDot.className = 'status-dot ' + (status === 'online' ? '' : status);
  dbStatusText.textContent = label;
}

function renderLastSyncTime(timestampStr) {
  if (!timestampStr) return;
  const d = new Date(timestampStr);
  if (isNaN(d)) return;

  const istStr = d.toLocaleString('en-IN', {
    timeZone: 'Asia/Kolkata',
    day: '2-digit', month: 'short', year: 'numeric',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true
  });
  const estStr = d.toLocaleString('en-US', {
    timeZone: 'America/New_York',
    day: '2-digit', month: 'short', year: 'numeric',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true
  });

  lastUpdatedIst.textContent = istStr;
  lastUpdatedEst.textContent = estStr;
  totalDbCount.textContent = allOrders.length;

  // Check revert eligibility
  const savedKeys = localStorage.getItem(UPLOAD_KEYS_KEY);
  if (savedKeys) {
    revertBtn.classList.remove('hidden');
  } else {
    revertBtn.classList.add('hidden');
  }
}

// ==========================================================================
// DATA NORMALIZATION (Ensures V2 Strict Calculation)
// ==========================================================================
function parseNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  const cleaned = String(val).replace(/[^0-9.\-]+/g, '');
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function getVal(row, ...keys) {
  for (const k of keys) {
    if (row[k] !== undefined && row[k] !== null && row[k] !== '') {
      return row[k];
    }
    const lowerKey = k.toLowerCase();
    for (const rk of Object.keys(row)) {
      if (rk.toLowerCase() === lowerKey && row[rk] !== undefined && row[rk] !== null && row[rk] !== '') {
        return row[rk];
      }
    }
  }
  return 0;
}

function normalizeOrders(rows, options = {}) {
  const { isRawUpload = false } = options;
  return rows
    .filter(r => (r.order_id || r['amazon-order-id'] || r['order-id']) && (r.sku || r['sku']))
    .filter(r => {
      const pl = String(r['profit/loss'] || r.profit_loss || '');
      return !pl.includes('Net Profit') && !pl.includes('Grand Total Net');
    })
    .map(r => {
      const orderId = String(r.order_id || r['amazon-order-id'] || r['order-id']).trim();
      const sku = String(r.sku || r['sku']).trim();
      const title = String(r.title || r['product-name'] || r['name'] || '').trim();
      
      let purchaseDate = r.purchase_date || r['purchase-date'] || r['order-date'] || null;
      if (purchaseDate) {
        const d = new Date(purchaseDate);
        if (!isNaN(d)) purchaseDate = d.toISOString();
      }

      const qty = parseInt(r.quantity || r['quantity'] || 1, 10) || 1;
      const itemPrice = parseNum(r.item_price || r['item-price'] || r.price);
      
      // Cost price: unit cost from sheet / DB
      const rawCostPrice = parseNum(getVal(r, 'cost_price', 'cost price', 'cost', 'Unit Cost'));
      const rawTotalCost = parseNum(getVal(r, 'total_cost', 'total cost', 'Total Cost'));

      let costPrice = rawCostPrice;
      let totalCost = rawTotalCost;

      if (isRawUpload) {
        // Cost price calculation on uploaded sheet values
        costPrice = parseFloat((rawCostPrice * 1.05).toFixed(2));
        totalCost = rawTotalCost > 0 
          ? parseFloat((rawTotalCost * 1.05).toFixed(2)) 
          : parseFloat((costPrice * qty).toFixed(2));
      } else {
        // From database or cache
        if (totalCost === 0 && costPrice > 0) {
          totalCost = parseFloat((costPrice * qty).toFixed(2));
        }
      }

      // Grand total: if present use it, else default to itemPrice * quantity
      const grandTotalRaw = parseNum(getVal(r, 'grand_total', 'grand total', 'Grand Total'));
      const grandTotal = grandTotalRaw > 0 ? grandTotalRaw : parseFloat((itemPrice * qty).toFixed(2));

      // Carrier shipping fetched
      const shippingCost = parseNum(getVal(r, 'shipping_cost', 'fetched shipping', 'shipping-price', 'Carrier Shipping'));

      // Amazon tax
      const amazonTax = parseNum(getVal(r, 'amazon_tax', 'amazon tax', 'tax', 'Amazon Tax'));

      // Indicator / Status
      let rawStatus = String(r.order_status || r['order-status'] || r.indicator || r['indicator'] || '').trim();
      const isCancelled = rawStatus.toLowerCase().includes('cancelled') || String(r.indicator || '').toLowerCase().includes('cancelled');

      // Strict V2 Profit/Loss formula:
      // Profit = Grand Total - Total Cost - Shipping - Amazon Tax
      let profitLoss = 0;
      if (!isCancelled) {
        profitLoss = parseFloat((grandTotal - totalCost - shippingCost - amazonTax).toFixed(2));
      }

      let indicator = '';
      if (isCancelled) {
        indicator = '⚪ Cancelled';
      } else {
        indicator = profitLoss >= 0 ? '🟢 Profit' : '🔴 Loss';
      }

      const storeName = String(r.store_name || currentStore || 'WISEROUTLET').trim().toUpperCase();

      return {
        order_id: orderId,
        sku: sku,
        title: title,
        quantity: qty,
        item_price: itemPrice,
        grand_total: grandTotal,
        cost_price: costPrice,
        total_cost: totalCost,
        shipping_cost: shippingCost,
        amazon_tax: amazonTax,
        profit_loss: profitLoss,
        indicator: indicator,
        order_status: rawStatus,
        store_name: storeName,
        purchase_date: purchaseDate
      };
    });
}

// ==========================================================================
// FILTERING & SEARCH ENGINE
// ==========================================================================
function applyFilters() {
  let list = [...allOrders];

  // 1. Store Filter
  if (currentStore !== 'ALL') {
    list = list.filter(o => (o.store_name || '').toUpperCase() === currentStore.toUpperCase());
  }

  // 2. Time Window Filter
  const now = new Date();
  if (currentTimeFilter === 'today') {
    const todayStr = now.toISOString().split('T')[0];
    list = list.filter(o => o.purchase_date && o.purchase_date.startsWith(todayStr));
  } else if (currentTimeFilter === '7days') {
    const cutoff = new Date();
    cutoff.setDate(now.getDate() - 7);
    list = list.filter(o => o.purchase_date && new Date(o.purchase_date) >= cutoff);
  } else if (currentTimeFilter === '30days') {
    const cutoff = new Date();
    cutoff.setDate(now.getDate() - 30);
    list = list.filter(o => o.purchase_date && new Date(o.purchase_date) >= cutoff);
  } else if (currentTimeFilter === 'thisMonth') {
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    list = list.filter(o => o.purchase_date && new Date(o.purchase_date) >= monthStart);
  } else if (currentTimeFilter === 'custom') {
    if (customDateFrom) {
      list = list.filter(o => o.purchase_date && new Date(o.purchase_date) >= customDateFrom);
    }
    if (customDateTo) {
      list = list.filter(o => o.purchase_date && new Date(o.purchase_date) <= customDateTo);
    }
  }

  // 3. Status Filter
  if (currentStatusFilter === 'shipped') {
    list = list.filter(o => !o.indicator.includes('Cancelled') && !o.order_status.toLowerCase().includes('cancelled'));
  } else if (currentStatusFilter === 'profit') {
    list = list.filter(o => (o.profit_loss || 0) >= 0 && !o.indicator.includes('Cancelled'));
  } else if (currentStatusFilter === 'loss') {
    list = list.filter(o => (o.profit_loss || 0) < 0 && !o.indicator.includes('Cancelled'));
  } else if (currentStatusFilter === 'cancelled') {
    list = list.filter(o => o.indicator.includes('Cancelled') || o.order_status.toLowerCase().includes('cancelled'));
  }

  // 4. Search Filter
  if (searchQuery.trim()) {
    const q = searchQuery.toLowerCase();
    list = list.filter(o => 
      (o.order_id && o.order_id.toLowerCase().includes(q)) ||
      (o.sku && o.sku.toLowerCase().includes(q)) ||
      (o.title && o.title.toLowerCase().includes(q))
    );
  }

  // 5. Sorting
  list.sort((a, b) => {
    let valA = a[sortColumn];
    let valB = b[sortColumn];

    if (valA === null || valA === undefined) valA = '';
    if (valB === null || valB === undefined) valB = '';

    if (typeof valA === 'number' && typeof valB === 'number') {
      return sortDirection === 'asc' ? valA - valB : valB - valA;
    }
    const strA = String(valA).toLowerCase();
    const strB = String(valB).toLowerCase();
    return sortDirection === 'asc' ? strA.localeCompare(strB) : strB.localeCompare(strA);
  });

  filteredOrders = list;
  currentPage = 1;

  renderDashboard();
}

// ==========================================================================
// DASHBOARD RENDERING
// ==========================================================================
function renderDashboard() {
  renderKPIs(filteredOrders);
  renderCharts(filteredOrders);
  renderTable(filteredOrders);
}

function renderKPIs(orders) {
  // We only count shipped orders for financial KPIs
  const shippedOrders = orders.filter(o => !o.indicator.includes('Cancelled') && !o.order_status.toLowerCase().includes('cancelled'));
  
  const totalOrdersCount = orders.length;
  const shippedCount = shippedOrders.length;
  const profitableCount = shippedOrders.filter(o => (o.profit_loss || 0) >= 0).length;
  const lossCount = shippedOrders.filter(o => (o.profit_loss || 0) < 0).length;
  const cancelledCount = totalOrdersCount - shippedCount;

  // Financial sums
  const totalGrandRevenue = shippedOrders.reduce((sum, o) => sum + (o.grand_total || 0), 0);
  const totalNetProfit = shippedOrders.reduce((sum, o) => sum + (o.profit_loss || 0), 0);
  const totalShipping = shippedOrders.reduce((sum, o) => sum + (o.shipping_cost || 0), 0);
  const totalTax = shippedOrders.reduce((sum, o) => sum + (o.amazon_tax || 0), 0);
  const totalCogs = shippedOrders.reduce((sum, o) => sum + (o.total_cost || 0), 0);

  const profitMargin = totalGrandRevenue > 0 ? ((totalNetProfit / totalGrandRevenue) * 100) : 0;
  const aov = shippedCount > 0 ? (totalGrandRevenue / shippedCount) : 0;
  const avgShipping = shippedCount > 0 ? (totalShipping / shippedCount) : 0;

  // 1. Net Profit Card
  kpiTotalProfit.textContent = `$${totalNetProfit.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  kpiTotalProfit.className = 'kpi-value ' + (totalNetProfit >= 0 ? 'profit-color' : 'loss-color');
  kpiProfitMargin.textContent = `${profitMargin >= 0 ? '+' : ''}${profitMargin.toFixed(1)}% Margin`;
  kpiProfitMargin.className = 'kpi-badge ' + (profitMargin >= 0 ? 'badge-profit' : 'badge-loss');
  kpiProfitCard.className = 'kpi-card glass-panel ' + (totalNetProfit >= 0 ? 'highlight-profit' : 'highlight-loss');

  // 2. Revenue Card
  kpiTotalRevenue.textContent = `$${totalGrandRevenue.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  kpiAov.textContent = `$${aov.toFixed(2)}`;

  // 3. Orders Count Card
  kpiTotalOrders.textContent = totalOrdersCount.toLocaleString();
  kpiShippedRatio.textContent = `${shippedCount} Shipped`;
  kpiOrderBreakdown.textContent = `${profitableCount} Profitable • ${lossCount} Loss • ${cancelledCount} Cancelled`;

  // 4. Shipping Card
  kpiTotalShipping.textContent = `$${totalShipping.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  kpiAvgShipping.textContent = `$${avgShipping.toFixed(2)}`;

  // 5. Amazon Tax Card
  kpiTotalTax.textContent = `$${totalTax.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

  // 6. COGS Card
  kpiTotalCogs.textContent = `$${totalCogs.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

// ==========================================================================
// CHARTS (Chart.js)
// ==========================================================================
function renderCharts(orders) {
  const shippedOrders = orders.filter(o => !o.indicator.includes('Cancelled') && !o.order_status.toLowerCase().includes('cancelled'));

  // 1. Timeline Chart Data Grouping
  const dayMap = {};
  shippedOrders.forEach(o => {
    if (!o.purchase_date) return;
    const day = o.purchase_date.split('T')[0];
    if (!dayMap[day]) {
      dayMap[day] = { revenue: 0, profit: 0, count: 0 };
    }
    dayMap[day].revenue += (o.grand_total || 0);
    dayMap[day].profit += (o.profit_loss || 0);
    dayMap[day].count += 1;
  });

  const sortedDays = Object.keys(dayMap).sort();
  const labels = sortedDays.map(d => {
    const parts = d.split('-');
    return `${parts[1]}/${parts[2]}`;
  });

  let revData = [];
  let profData = [];

  if (timelineViewMode === 'cumulative') {
    let runRev = 0;
    let runProf = 0;
    sortedDays.forEach(d => {
      runRev += dayMap[d].revenue;
      runProf += dayMap[d].profit;
      revData.push(parseFloat(runRev.toFixed(2)));
      profData.push(parseFloat(runProf.toFixed(2)));
    });
  } else {
    sortedDays.forEach(d => {
      revData.push(parseFloat(dayMap[d].revenue.toFixed(2)));
      profData.push(parseFloat(dayMap[d].profit.toFixed(2)));
    });
  }

  // Draw Timeline Chart
  if (timelineChart) timelineChart.destroy();
  timelineChart = new Chart(timelineCanvas, {
    type: 'bar',
    data: {
      labels: labels.length ? labels : ['No Data'],
      datasets: [
        {
          label: 'Net Profit ($)',
          data: profData.length ? profData : [0],
          type: 'line',
          borderColor: '#10b981',
          backgroundColor: 'rgba(16, 185, 129, 0.15)',
          fill: true,
          tension: 0.35,
          borderWidth: 2.5,
          pointRadius: 4,
          pointBackgroundColor: '#10b981',
          yAxisID: 'y'
        },
        {
          label: 'Grand Revenue ($)',
          data: revData.length ? revData : [0],
          type: 'bar',
          backgroundColor: 'rgba(99, 102, 241, 0.35)',
          borderColor: '#6366f1',
          borderWidth: 1.5,
          borderRadius: 6,
          yAxisID: 'y'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: 'index',
        intersect: false,
      },
      scales: {
        y: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: { color: '#94a3b8', font: { family: 'Inter', size: 11 } }
        },
        x: {
          grid: { display: false },
          ticks: { color: '#94a3b8', font: { family: 'Inter', size: 11 } }
        }
      },
      plugins: {
        legend: {
          position: 'top',
          labels: { color: '#f8fafc', font: { family: 'Inter', size: 12, weight: 600 } }
        },
        tooltip: {
          backgroundColor: '#0f172a',
          titleColor: '#fff',
          bodyColor: '#94a3b8',
          borderColor: 'rgba(255,255,255,0.1)',
          borderWidth: 1,
          padding: 10,
          cornerRadius: 8
        }
      }
    }
  });

  // 2. Breakdown Donut Chart
  const totalCogs = shippedOrders.reduce((s, o) => s + (o.total_cost || 0), 0);
  const totalShipping = shippedOrders.reduce((s, o) => s + (o.shipping_cost || 0), 0);
  const totalTax = shippedOrders.reduce((s, o) => s + (o.amazon_tax || 0), 0);
  const totalNetProfit = Math.max(0, shippedOrders.reduce((s, o) => s + (o.profit_loss || 0), 0));

  if (breakdownChart) breakdownChart.destroy();
  breakdownChart = new Chart(breakdownCanvas, {
    type: 'doughnut',
    data: {
      labels: ['Total Cost', 'Carrier Shipping', 'Amazon Tax', 'Net Profit'],
      datasets: [{
        data: [
          parseFloat(totalCogs.toFixed(2)),
          parseFloat(totalShipping.toFixed(2)),
          parseFloat(totalTax.toFixed(2)),
          parseFloat(totalNetProfit.toFixed(2))
        ],
        backgroundColor: [
          '#64748b', // COGS
          '#f59e0b', // Shipping
          '#06b6d4', // Tax
          '#10b981'  // Profit
        ],
        borderColor: '#0f172a',
        borderWidth: 2,
        hoverOffset: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { color: '#94a3b8', font: { family: 'Inter', size: 11 }, boxWidth: 12 }
        },
        tooltip: {
          backgroundColor: '#0f172a',
          titleColor: '#fff',
          bodyColor: '#94a3b8',
          borderColor: 'rgba(255,255,255,0.1)',
          borderWidth: 1,
          padding: 10,
          callbacks: {
            label: function(ctx) {
              const val = ctx.raw || 0;
              const sum = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct = sum > 0 ? ((val / sum) * 100).toFixed(1) : 0;
              return ` ${ctx.label}: $${val.toLocaleString()} (${pct}%)`;
            }
          }
        }
      },
      cutout: '68%'
    }
  });
}

// ==========================================================================
// TABLE RENDERING & PAGINATION
// ==========================================================================
function renderTable(orders) {
  tableShowingText.textContent = `Showing ${orders.length.toLocaleString()} matching orders`;

  if (orders.length === 0) {
    ordersTbody.innerHTML = `
      <tr>
        <td colspan="11" style="text-align: center; padding: 3rem 1rem; color: var(--text-dim);">
          <div style="font-size: 2rem; margin-bottom: 0.5rem;">🔍</div>
          <div style="font-size: 0.95rem; font-weight: 600; color: var(--text-muted);">No matching orders found</div>
          <div style="font-size: 0.8rem; margin-top: 0.25rem;">Try adjusting your filters, search terms, or store selection.</div>
        </td>
      </tr>
    `;
    paginationInfo.textContent = 'Showing 0 to 0 of 0 records';
    paginationButtons.innerHTML = '';
    return;
  }

  const totalPages = Math.ceil(orders.length / rowsPerPage);
  if (currentPage > totalPages) currentPage = totalPages;
  const startIndex = (currentPage - 1) * rowsPerPage;
  const endIndex = Math.min(startIndex + rowsPerPage, orders.length);
  const pageOrders = orders.slice(startIndex, endIndex);

  ordersTbody.innerHTML = pageOrders.map(o => {
    const dateFormatted = o.purchase_date 
      ? new Date(o.purchase_date).toLocaleDateString('en-US', { month: 'short', day: '2-digit', year: 'numeric' })
      : '—';

    const isLoss = (o.profit_loss || 0) < 0;
    const isCancelled = o.indicator.includes('Cancelled') || o.order_status.toLowerCase().includes('cancelled');
    
    let indicatorHtml = '';
    if (isCancelled) {
      indicatorHtml = `<span class="badge-indicator cancelled">⚪ Cancelled</span>`;
    } else if (isLoss) {
      indicatorHtml = `<span class="badge-indicator shipped-loss">🔴 Loss</span>`;
    } else {
      indicatorHtml = `<span class="badge-indicator shipped-profit">🟢 Profit</span>`;
    }

    const profitClass = isCancelled ? '' : (isLoss ? 'cell-profit-neg' : 'cell-profit-pos');
    const profitSign = (o.profit_loss || 0) >= 0 ? '+' : '';

    return `
      <tr>
        <td>${dateFormatted}</td>
        <td>
          <span class="cell-order-id">
            ${o.order_id}
            <button class="copy-mini-btn" data-copy="${o.order_id}" title="Copy Order ID">📋</button>
          </span>
        </td>
        <td><span style="font-weight: 600; font-size: 0.75rem; color: var(--text-muted);">${o.store_name}</span></td>
        <td>
          <div class="cell-product-name" title="${o.title || ''}">${o.title || '—'}</div>
          <span class="cell-sku">${o.sku}</span>
        </td>
        <td style="font-weight: 600;">${o.quantity}</td>
        <td style="font-weight: 700; color: #fff;">$${(o.grand_total || 0).toFixed(2)}</td>
        <td style="color: var(--text-muted);" title="Total Cost: $${(o.total_cost || 0).toFixed(2)} (Unit Cost: $${(o.cost_price || 0).toFixed(2)})">
          <div>$${(o.total_cost || 0).toFixed(2)}</div>
          ${o.quantity > 1 ? `<div style="font-size: 0.72rem; color: var(--text-dim);">$${(o.cost_price || 0).toFixed(2)} ea</div>` : ''}
        </td>
        <td style="color: var(--warning);">$${(o.shipping_cost || 0).toFixed(2)}</td>
        <td style="color: var(--accent-cyan);">$${(o.amazon_tax || 0).toFixed(2)}</td>
        <td class="${profitClass}">${profitSign}$${(o.profit_loss || 0).toFixed(2)}</td>
        <td>${indicatorHtml}</td>
      </tr>
    `;
  }).join('');

  // Wire up copy buttons
  ordersTbody.querySelectorAll('.copy-mini-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const text = btn.getAttribute('data-copy');
      navigator.clipboard.writeText(text);
      showToast('Copied', `Order ID ${text} copied to clipboard!`, 'info', 2000);
    });
  });

  // Render pagination
  paginationInfo.textContent = `Showing ${startIndex + 1} to ${endIndex} of ${orders.length.toLocaleString()} records`;

  let pageBtnsHtml = '';
  if (totalPages > 1) {
    pageBtnsHtml += `<button class="page-num-btn" id="prev-page-btn" ${currentPage === 1 ? 'disabled' : ''}>‹</button>`;
    
    // Show max 5 page buttons
    let startPage = Math.max(1, currentPage - 2);
    let endPage = Math.min(totalPages, startPage + 4);
    if (endPage - startPage < 4) startPage = Math.max(1, endPage - 4);

    for (let p = startPage; p <= endPage; p++) {
      pageBtnsHtml += `<button class="page-num-btn ${p === currentPage ? 'active' : ''}" data-page="${p}">${p}</button>`;
    }

    pageBtnsHtml += `<button class="page-num-btn" id="next-page-btn" ${currentPage === totalPages ? 'disabled' : ''}>›</button>`;
  }
  paginationButtons.innerHTML = pageBtnsHtml;

  const prevBtn = document.getElementById('prev-page-btn');
  if (prevBtn) prevBtn.addEventListener('click', () => { if (currentPage > 1) { currentPage--; renderTable(filteredOrders); } });

  const nextBtn = document.getElementById('next-page-btn');
  if (nextBtn) nextBtn.addEventListener('click', () => { if (currentPage < totalPages) { currentPage++; renderTable(filteredOrders); } });

  paginationButtons.querySelectorAll('[data-page]').forEach(b => {
    b.addEventListener('click', (e) => {
      currentPage = parseInt(e.target.dataset.page, 10);
      renderTable(filteredOrders);
    });
  });
}

// Table column sorting headers
document.querySelectorAll('#orders-table th[data-sort]').forEach(th => {
  th.addEventListener('click', () => {
    const col = th.dataset.sort;
    if (sortColumn === col) {
      sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
      sortColumn = col;
      sortDirection = 'desc';
    }
    applyFilters();
  });
});

// Rows per page
rowsPerPageSelect.addEventListener('change', (e) => {
  rowsPerPage = parseInt(e.target.value, 10);
  currentPage = 1;
  renderTable(filteredOrders);
});

// ==========================================================================
// SEARCH & AUTOCOMPLETE
// ==========================================================================
let searchDebounceTimer = null;
searchInput.addEventListener('input', () => {
  clearTimeout(searchDebounceTimer);
  const q = searchInput.value.trim();
  searchQuery = q;

  if (q.length > 0) {
    searchClearBtn.classList.remove('hidden');
  } else {
    searchClearBtn.classList.add('hidden');
    searchSuggestions.classList.add('hidden');
  }

  searchDebounceTimer = setTimeout(() => {
    if (q.length >= 2) {
      buildSuggestions(q);
    } else {
      searchSuggestions.classList.add('hidden');
    }
    applyFilters();
  }, 200);
});

searchClearBtn.addEventListener('click', () => {
  searchInput.value = '';
  searchQuery = '';
  searchClearBtn.classList.add('hidden');
  searchSuggestions.classList.add('hidden');
  applyFilters();
});

function buildSuggestions(q) {
  const query = q.toLowerCase();
  const suggestions = [];
  const seen = new Set();

  for (const o of allOrders) {
    if (suggestions.length >= 6) break;

    if (o.order_id && o.order_id.toLowerCase().includes(query) && !seen.has(o.order_id)) {
      seen.add(o.order_id);
      suggestions.push({ type: 'ORDER', val: o.order_id });
    }
    if (o.sku && String(o.sku).toLowerCase().includes(query) && !seen.has(o.sku)) {
      seen.add(o.sku);
      suggestions.push({ type: 'SKU', val: String(o.sku) });
    }
    if (o.title && o.title.toLowerCase().includes(query) && !seen.has(o.title)) {
      seen.add(o.title);
      suggestions.push({ type: 'PRODUCT', val: o.title.length > 55 ? o.title.slice(0, 55) + '…' : o.title });
    }
  }

  if (suggestions.length > 0) {
    searchSuggestions.innerHTML = suggestions.map(s => `
      <div class="suggestion-item" data-value="${s.val}">
        <span class="suggestion-type-tag">${s.type}</span>
        <span>${s.val}</span>
      </div>
    `).join('');
    searchSuggestions.classList.remove('hidden');

    searchSuggestions.querySelectorAll('.suggestion-item').forEach(item => {
      item.addEventListener('click', () => {
        const val = item.getAttribute('data-value');
        searchInput.value = val;
        searchQuery = val;
        searchSuggestions.classList.add('hidden');
        applyFilters();
      });
    });
  } else {
    searchSuggestions.classList.add('hidden');
  }
}

document.addEventListener('click', (e) => {
  if (!e.target.closest('.search-box-wrapper')) {
    searchSuggestions.classList.add('hidden');
  }
});

// ==========================================================================
// STORE & FILTER CONTROLS
// ==========================================================================
storeSelector.addEventListener('change', (e) => {
  currentStore = e.target.value;
  brandLogoIcon.textContent = currentStore === 'ALL' ? 'A' : currentStore.charAt(0).toUpperCase();
  applyFilters();
});

timeFilterPills.querySelectorAll('.pill-btn').forEach(btn => {
  btn.addEventListener('click', (e) => {
    timeFilterPills.querySelectorAll('.pill-btn').forEach(b => b.classList.remove('active'));
    e.target.classList.add('active');
    currentTimeFilter = e.target.dataset.time;
    applyFilters();
  });
});

statusFilterPills.querySelectorAll('.pill-btn').forEach(btn => {
  btn.addEventListener('click', (e) => {
    statusFilterPills.querySelectorAll('.pill-btn').forEach(b => b.classList.remove('active'));
    e.target.classList.add('active');
    currentStatusFilter = e.target.dataset.status;
    applyFilters();
  });
});

applyDatesBtn.addEventListener('click', () => {
  if (dateFromInput.value) customDateFrom = new Date(dateFromInput.value);
  if (dateToInput.value) customDateTo = new Date(dateToInput.value + 'T23:59:59');
  
  timeFilterPills.querySelectorAll('.pill-btn').forEach(b => b.classList.remove('active'));
  currentTimeFilter = 'custom';
  applyFilters();
});

resetDatesBtn.addEventListener('click', () => {
  dateFromInput.value = '';
  dateToInput.value = '';
  customDateFrom = null;
  customDateTo = null;
  const allTimeBtn = timeFilterPills.querySelector('[data-time="all"]');
  if (allTimeBtn) allTimeBtn.click();
});

chartViewDaily.addEventListener('click', () => {
  timelineViewMode = 'daily';
  chartViewDaily.classList.add('active');
  chartViewCumulative.classList.remove('active');
  renderCharts(filteredOrders);
});

chartViewCumulative.addEventListener('click', () => {
  timelineViewMode = 'cumulative';
  chartViewCumulative.classList.add('active');
  chartViewDaily.classList.remove('active');
  renderCharts(filteredOrders);
});

// Refresh button in sync strip
refreshDbBtn.addEventListener('click', () => {
  fetchOrdersFromDatabase();
});

// ==========================================================================
// EXCEL EXPORT
// ==========================================================================
exportExcelBtn.addEventListener('click', () => {
  if (!filteredOrders.length) {
    showToast('Export Notice', 'No orders available to export.', 'info');
    return;
  }

  const exportData = filteredOrders.map(o => ({
    'Date': o.purchase_date ? o.purchase_date.split('T')[0] : '',
    'Order ID': o.order_id,
    'Store': o.store_name,
    'SKU': o.sku,
    'Product Title': o.title,
    'Quantity': o.quantity,
    'Unit Price': o.item_price,
    'Grand Total': o.grand_total,
    'Unit Cost': o.cost_price,
    'Total Cost': o.total_cost,
    'Carrier Shipping': o.shipping_cost,
    'Amazon Tax': o.amazon_tax,
    'Net Profit': o.profit_loss,
    'Indicator': o.indicator,
    'Status': o.order_status
  }));

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Master Orders');
  
  const dateStr = new Date().toISOString().split('T')[0];
  const filename = `Amazon_Orders_Export_${currentStore}_${dateStr}.xlsx`;
  XLSX.writeFile(workbook, filename);
  showToast('Export Complete', `Saved file ${filename}`, 'success');
});

// ==========================================================================
// MASTERDATA UPLOAD & DATABASE UPSERT (V2)
// ==========================================================================
openUploadModalBtn.addEventListener('click', () => {
  uploadModal.classList.remove('hidden');
  resetUploadModal();
});

closeUploadModalBtn.addEventListener('click', () => uploadModal.classList.add('hidden'));
cancelUploadBtn.addEventListener('click', () => uploadModal.classList.add('hidden'));

function resetUploadModal() {
  masterFileInput.value = '';
  pendingUploadRows = [];
  uploadPreviewSection.classList.add('hidden');
  uploadProgressWrap.classList.add('hidden');
  confirmUploadBtn.disabled = true;
  confirmUploadBtn.textContent = 'Save & Sync to Database';
}

// Drag & drop support
dropzone.addEventListener('click', () => masterFileInput.click());
dropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropzone.classList.add('drag-over');
});
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag-over'));
dropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropzone.classList.remove('drag-over');
  if (e.dataTransfer.files.length) {
    handleSelectedFile(e.dataTransfer.files[0]);
  }
});

masterFileInput.addEventListener('change', (e) => {
  if (e.target.files.length) {
    handleSelectedFile(e.target.files[0]);
  }
});

function handleSelectedFile(file) {
  previewFilename.textContent = `📄 ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

      const targetStore = uploadStoreSelect.value;
      const normalized = normalizeOrders(rawRows.map(r => ({ ...r, store_name: targetStore })), { isRawUpload: true });

      if (!normalized.length) {
        showToast('Parse Error', 'No valid order records found with Order ID and SKU.', 'error');
        return;
      }

      pendingUploadRows = normalized;

      const shipped = normalized.filter(o => !o.indicator.includes('Cancelled'));
      const netProfit = shipped.reduce((s, o) => s + (o.profit_loss || 0), 0);

      previewTotalRows.textContent = normalized.length.toLocaleString();
      previewShippedRows.textContent = shipped.length.toLocaleString();
      previewTotalProfit.textContent = `$${netProfit.toLocaleString('en-US', { minimumFractionDigits: 2 })}`;
      previewTotalProfit.style.color = netProfit >= 0 ? 'var(--profit)' : 'var(--loss)';

      uploadPreviewSection.classList.remove('hidden');
      confirmUploadBtn.disabled = false;
    } catch (err) {
      console.error(err);
      showToast('File Error', `Failed to parse Excel file: ${err.message}`, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

confirmUploadBtn.addEventListener('click', async () => {
  if (!pendingUploadRows.length) return;

  confirmUploadBtn.disabled = true;
  uploadProgressWrap.classList.remove('hidden');

  const total = pendingUploadRows.length;
  const BATCH_SIZE = 50;
  const totalBatches = Math.ceil(total / BATCH_SIZE);
  let processed = 0;

  try {
    for (let i = 0; i < total; i += BATCH_SIZE) {
      const batch = pendingUploadRows.slice(i, i + BATCH_SIZE);
      const batchNum = Math.floor(i / BATCH_SIZE) + 1;
      
      const pct = Math.round((batchNum / totalBatches) * 100);
      uploadProgressBar.style.width = `${pct}%`;
      uploadProgressPct.textContent = `${pct}%`;
      uploadProgressText.textContent = `Syncing batch ${batchNum} of ${totalBatches} (${batch.length} records)…`;

      // Try full V2 schema upsert
      let { error } = await supabase
        .from('orders')
        .upsert(batch, { onConflict: 'order_id,sku,store_name' });

      // Fallback if custom columns not yet migrated in Supabase
      if (error && error.message && error.message.includes('column')) {
        console.warn('V2 extended columns not yet created in Supabase. Falling back to base columns:', error.message);
        const fallbackBatch = batch.map(b => ({
          order_id: b.order_id,
          sku: b.sku,
          title: b.title,
          item_price: b.item_price,
          cost_price: b.cost_price,
          shipping_cost: b.shipping_cost,
          profit_loss: b.profit_loss,
          indicator: b.indicator,
          store_name: b.store_name,
          purchase_date: b.purchase_date
        }));

        const retry = await supabase
          .from('orders')
          .upsert(fallbackBatch, { onConflict: 'order_id,sku,store_name' });
        
        if (retry.error) throw retry.error;
      } else if (error) {
        throw error;
      }

      processed += batch.length;
    }

    // Save upload keys for Revert functionality
    const uploadKeys = pendingUploadRows.map(r => ({ order_id: r.order_id, sku: r.sku, store_name: r.store_name }));
    localStorage.setItem(UPLOAD_KEYS_KEY, JSON.stringify(uploadKeys));
    localStorage.setItem(SYNC_TIME_KEY, new Date().toISOString());

    // Merge into local cache
    const existingMap = new Map();
    allOrders.forEach(o => existingMap.set(`${o.order_id}_${o.sku}_${o.store_name}`, o));
    pendingUploadRows.forEach(o => existingMap.set(`${o.order_id}_${o.sku}_${o.store_name}`, o));
    allOrders = Array.from(existingMap.values());
    saveOrdersToCache(allOrders);

    showToast('Sync Successful', `Successfully saved & synced ${total} orders to cloud database!`, 'success');
    renderLastSyncTime(new Date().toISOString());
    uploadModal.classList.add('hidden');
    applyFilters();
  } catch (err) {
    console.error('Database Sync Error:', err);
    // If Supabase was unreachable, still save to local cache so user can work uninterrupted
    const existingMap = new Map();
    allOrders.forEach(o => existingMap.set(`${o.order_id}_${o.sku}_${o.store_name}`, o));
    pendingUploadRows.forEach(o => existingMap.set(`${o.order_id}_${o.sku}_${o.store_name}`, o));
    allOrders = Array.from(existingMap.values());
    saveOrdersToCache(allOrders);
    
    showToast('Saved to Local Cache', `Database offline (${err.message}). Data saved safely in Local Cache.`, 'warning', 6000);
    renderLastSyncTime(new Date().toISOString());
    uploadModal.classList.add('hidden');
    applyFilters();
  }
});

// ==========================================================================
// REVERT LAST UPLOAD
// ==========================================================================
revertBtn.addEventListener('click', async () => {
  const savedKeys = localStorage.getItem(UPLOAD_KEYS_KEY);
  if (!savedKeys) return;

  const keys = JSON.parse(savedKeys);
  if (!confirm(`Are you sure you want to revert the last upload? This will remove ${keys.length} records.`)) {
    return;
  }

  revertBtn.textContent = 'Reverting…';
  revertBtn.disabled = true;

  try {
    for (const k of keys) {
      await supabase
        .from('orders')
        .delete()
        .eq('order_id', k.order_id)
        .eq('sku', k.sku)
        .eq('store_name', k.store_name);
    }

    // Remove from in-memory and local cache
    const deleteSet = new Set(keys.map(k => `${k.order_id}_${k.sku}_${k.store_name}`));
    allOrders = allOrders.filter(o => !deleteSet.has(`${o.order_id}_${o.sku}_${o.store_name}`));
    saveOrdersToCache(allOrders);

    localStorage.removeItem(UPLOAD_KEYS_KEY);
    revertBtn.classList.add('hidden');
    showToast('Reverted', `Removed ${keys.length} records from database.`, 'info');
    applyFilters();
  } catch (err) {
    showToast('Revert Notice', `Could not delete from database (${err.message}). Removed locally.`, 'warning');
    const deleteSet = new Set(keys.map(k => `${k.order_id}_${k.sku}_${k.store_name}`));
    allOrders = allOrders.filter(o => !deleteSet.has(`${o.order_id}_${o.sku}_${o.store_name}`));
    saveOrdersToCache(allOrders);
    localStorage.removeItem(UPLOAD_KEYS_KEY);
    revertBtn.classList.add('hidden');
    applyFilters();
  } finally {
    revertBtn.textContent = '↩ Revert Last Upload';
    revertBtn.disabled = false;
  }
});

// ==========================================================================
// SETTINGS & STORAGE MANAGEMENT
// ==========================================================================
openSettingsBtn.addEventListener('click', () => {
  settingsModal.classList.remove('hidden');
});

dbStatusBtn.addEventListener('click', () => openSettingsBtn.click());
closeSettingsModalBtn.addEventListener('click', () => settingsModal.classList.add('hidden'));
cancelSettingsBtn.addEventListener('click', () => settingsModal.classList.add('hidden'));

clearCacheBtn.addEventListener('click', () => {
  if (confirm('Clear local cache? This will reset local data and re-fetch from the database.')) {
    localStorage.removeItem(CACHE_KEY);
    localStorage.removeItem('v2_orders_cache_v2');
    localStorage.removeItem(UPLOAD_KEYS_KEY);
    localStorage.removeItem(SYNC_TIME_KEY);
    allOrders = [];
    applyFilters();
    showToast('Cache Cleared', 'Local cache cleared. Fetching fresh data…', 'info');
    settingsModal.classList.add('hidden');
    fetchOrdersFromDatabase();
  }
});

// ==========================================================================
// APPLICATION INITIALIZATION
// ==========================================================================
(function init() {
  // 1. Initial cached state load for instantaneous visual response
  const cached = loadOrdersFromCache();
  if (cached.length > 0) {
    allOrders = cached;
    applyFilters();
  }

  const lastSync = localStorage.getItem(SYNC_TIME_KEY);
  if (lastSync) renderLastSyncTime(lastSync);

  // 2. Fetch fresh from database
  fetchOrdersFromDatabase();
})();
