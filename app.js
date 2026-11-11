/**
 * Flujo de Caja Personal - PWA Application Logic
 * Supports: Mobile-first responsive UI, Offline storage, GitHub REST API Sync, Excel Export/Import
 */

// 1. CONFIGURATION & CATALOGS
const CALENDAR_MONTHS = [
  { name: 'Octubre 2026', start: '2026-10-01', end: '2026-10-31', short: 'Oct-26', year: 2026 },
  { name: 'Noviembre 2026', start: '2026-11-01', end: '2026-11-30', short: 'Nov-26', year: 2026 },
  { name: 'Diciembre 2026', start: '2026-12-01', end: '2026-12-31', short: 'Dic-26', year: 2026 },
  { name: 'Enero 2027', start: '2027-01-01', end: '2027-01-31', short: 'Ene-27', year: 2027 },
  { name: 'Febrero 2027', start: '2027-02-01', end: '2027-02-28', short: 'Feb-27', year: 2027 },
  { name: 'Marzo 2027', start: '2027-03-01', end: '2027-03-31', short: 'Mar-27', year: 2027 },
  { name: 'Abril 2027', start: '2027-04-01', end: '2027-04-30', short: 'Abr-27', year: 2027 },
  { name: 'Mayo 2027', start: '2027-05-01', end: '2027-05-31', short: 'May-27', year: 2027 },
  { name: 'Junio 2027', start: '2027-06-01', end: '2027-06-30', short: 'Jun-27', year: 2027 },
  { name: 'Julio 2027', start: '2027-07-01', end: '2027-07-31', short: 'Jul-27', year: 2027 },
  { name: 'Agosto 2027', start: '2027-08-01', end: '2027-08-31', short: 'Ago-27', year: 2027 },
  { name: 'Septiembre 2027', start: '2027-09-01', end: '2027-09-30', short: 'Sep-27', year: 2027 },
  { name: 'Octubre 2027', start: '2027-10-01', end: '2027-10-31', short: 'Oct-27', year: 2027 },
  { name: 'Noviembre 2027', start: '2027-11-01', end: '2027-11-30', short: 'Nov-27', year: 2027 },
  { name: 'Diciembre 2027', start: '2027-12-01', end: '2027-12-31', short: 'Dic-27', year: 2027 }
];

const CATEGORIES = [
  // INGRESOS
  { name: 'Sueldo / Salario', group: 'Ingreso', budget: 600000 },
  { name: 'Honorarios / Freelance', group: 'Ingreso', budget: 0 },
  { name: 'Rendimientos / Inversiones', group: 'Ingreso', budget: 0 },
  { name: 'Otros Ingresos', group: 'Ingreso', budget: 0 },
  // GASTOS FIJOS
  { name: 'Vivienda (Alquiler / Hipoteca)', group: 'Gasto Fijo', budget: 180000 },
  { name: 'Expensas y Mantenimiento Edificio', group: 'Gasto Fijo', budget: 35000 },
  { name: 'Servicios Básicos (Luz, Gas, Agua)', group: 'Gasto Fijo', budget: 25000 },
  { name: 'Internet, Telefonía y Cable', group: 'Gasto Fijo', budget: 22000 },
  { name: 'Seguros (Salud / Auto / Hogar)', group: 'Gasto Fijo', budget: 30000 },
  { name: 'Educación / Cursos / Cuotas', group: 'Gasto Fijo', budget: 20000 },
  { name: 'Suscripciones Digitales (Streaming, Apps)', group: 'Gasto Fijo', budget: 21000 },
  { name: 'Pago de Deudas y Préstamos', group: 'Gasto Fijo', budget: 0 },
  // GASTOS VARIABLES
  { name: 'Supermercado y Alimentos del Hogar', group: 'Gasto Variable', budget: 120000 },
  { name: 'Danza, Ballet y Actividades Deportivas', group: 'Gasto Variable', budget: 240000 },
  { name: 'Transporte y Combustible', group: 'Gasto Variable', budget: 40000 },
  { name: 'Farmacia y Salud', group: 'Gasto Variable', budget: 15000 },
  { name: 'Cuidado Personal y Limpieza', group: 'Gasto Variable', budget: 15000 },
  { name: 'Ropa y Calzado', group: 'Gasto Variable', budget: 25000 },
  { name: 'Salidas, Bares y Restaurantes', group: 'Gasto Variable', budget: 35000 },
  { name: 'Delivery y Comida Rápida', group: 'Gasto Variable', budget: 20000 },
  // GASTOS HORMIGA
  { name: 'Café al paso y Desayunos fuera', group: 'Gasto Hormiga', budget: 30000 },
  { name: 'Kiosco, Snacks y Golosinas', group: 'Gasto Hormiga', budget: 8000 },
  { name: 'Propinas y Pequeños Gastos en Efectivo', group: 'Gasto Hormiga', budget: 6000 },
  { name: 'Donaciones, Rifas y Colaboraciones', group: 'Gasto Hormiga', budget: 5000 },
  { name: 'Compras Impulsivas / Antojos', group: 'Gasto Hormiga', budget: 10000 },
  // AHORRO / INVERSION
  { name: 'Fondo de Emergencia', group: 'Ahorro / Inversión', budget: 40000 },
  { name: 'Inversiones y Metas de Ahorro', group: 'Ahorro / Inversión', budget: 50000 }
];

// 2. STATE
let state = {
  transactions: [],
  selectedMonth: 'Noviembre 2026',
  saldoInicialBase: 0,
  githubToken: localStorage.getItem('gh_token') || '',
  githubRepo: 'PiteIbarbia/FlujoDeCajaPersonal',
  currentTab: 'tablero'
};

let chartWeeklyInstance = null;
let chartDistInstance = null;

// Currency Formatter
const formatMoney = (num) => {
  return new Intl.NumberFormat('es-AR', {
    style: 'currency',
    currency: 'ARS',
    minimumFractionDigits: 2
  }).format(num || 0);
};

// Toast notifications
function showToast(message, type = 'info') {
  let toast = document.getElementById('appToast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'appToast';
    document.body.appendChild(toast);
  }

  let bgClass = 'bg-slate-800 text-white';
  if (type === 'success') bgClass = 'bg-emerald-600 text-white';
  else if (type === 'warning') bgClass = 'bg-amber-600 text-white';
  else if (type === 'error') bgClass = 'bg-rose-600 text-white';

  toast.className = `fixed top-14 left-1/2 -translate-x-1/2 z-50 px-4 py-2.5 rounded-2xl shadow-2xl text-xs font-bold transition-all duration-300 pointer-events-none max-w-[90vw] text-center ${bgClass} opacity-100 transform translate-y-0`;
  toast.textContent = message;

  setTimeout(() => {
    toast.classList.remove('opacity-100', 'translate-y-0');
    toast.classList.add('opacity-0', '-translate-y-2');
  }, 4000);
}

// Calculate Week Number of the Month
const getWeekOfMonth = (dateStr) => {
  const day = parseInt(dateStr.split('-')[2], 10);
  const weekNum = Math.min(5, Math.floor((day - 1) / 7) + 1);
  return `Semana ${weekNum}`;
};

// 3. INITIALIZATION
document.addEventListener('DOMContentLoaded', async () => {
  initMonthSelector();
  initCategoryDropdown();
  initNavigation();
  initModalEvents();
  initQuickButtons();
  initBackupEvents();

  // Load transactions from localStorage or fallback to data/transacciones.json
  await loadTransactions();

  // Render current view
  renderDashboard();
  renderTransactionsList();
  renderAnnualConsolidated();

  // Auto-sync with GitHub on startup if token is configured
  if (state.githubToken) {
    syncPullGitHub(true);
  }

  // Register Service Worker for offline PWA
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js')
      .then(() => console.log('PWA Service Worker registered'))
      .catch((err) => console.warn('SW registration failed:', err));
  }
});

// Populate Month Selectors
function initMonthSelector() {
  const monthSelect = document.getElementById('monthSelect');
  monthSelect.innerHTML = '';
  CALENDAR_MONTHS.forEach(m => {
    const opt = document.createElement('option');
    opt.value = m.name;
    opt.textContent = m.name;
    if (m.name === state.selectedMonth) opt.selected = true;
    monthSelect.appendChild(opt);
  });

  monthSelect.addEventListener('change', (e) => {
    state.selectedMonth = e.target.value;
    renderDashboard();
    renderTransactionsList();
  });
}

// Populate Category Dropdown based on Type
function initCategoryDropdown() {
  const tipoSelect = document.getElementById('txTipo');
  const catSelect = document.getElementById('txCategoria');

  const updateCategories = () => {
    const selectedTipo = tipoSelect.value;
    catSelect.innerHTML = '';
    const filtered = CATEGORIES.filter(c => c.group === selectedTipo);
    filtered.forEach(c => {
      const opt = document.createElement('option');
      opt.value = c.name;
      opt.textContent = c.name;
      catSelect.appendChild(opt);
    });
  };

  tipoSelect.addEventListener('change', updateCategories);
  updateCategories();
}

// Navigation Tabs
function initNavigation() {
  const tabs = document.querySelectorAll('.nav-tab');
  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const target = tab.dataset.tab;
      state.currentTab = target;

      tabs.forEach(t => {
        t.classList.remove('border-emerald-400', 'text-white', 'font-semibold');
        t.classList.add('border-transparent', 'text-slate-300');
      });
      tab.classList.add('border-emerald-400', 'text-white', 'font-semibold');
      tab.classList.remove('border-transparent', 'text-slate-300');

      document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.add('hidden');
      });
      document.getElementById(`tab-${target}`).classList.remove('hidden');

      if (target === 'tablero') renderDashboard();
      if (target === 'movimientos') renderTransactionsList();
      if (target === 'anual') renderAnnualConsolidated();
    });
  });
}

// Quick Buttons Handler
function initQuickButtons() {
  document.querySelectorAll('.quick-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      openModal();
      document.getElementById('txTipo').value = btn.dataset.tipo;
      initCategoryDropdown();
      document.getElementById('txCategoria').value = btn.dataset.cat;
      document.getElementById('txConcepto').value = btn.dataset.desc;
      document.getElementById('txMedioPago').value = btn.dataset.mp;
      document.getElementById('txMonto').focus();
    });
  });
}

// 4. DATA PERSISTENCE & LOADING
async function loadTransactions() {
  const localData = localStorage.getItem('cashflow_transactions');
  if (localData) {
    try {
      state.transactions = JSON.parse(localData);
      updateSyncStatus('Local');
      return;
    } catch (e) {
      console.warn('Error parsing localStorage:', e);
    }
  }

  // Fetch initial data/transacciones.json
  try {
    const res = await fetch('./data/transacciones.json');
    if (res.ok) {
      const data = await res.json();
      state.transactions = data.transacciones || [];
      state.saldoInicialBase = data.saldoInicialBase || 0;
      saveLocal();
      updateSyncStatus('Local');
    }
  } catch (err) {
    console.warn('Fallback loading data/transacciones.json:', err);
  }
}

function saveLocal() {
  localStorage.setItem('cashflow_transactions', JSON.stringify(state.transactions));
}

// 5. MODAL & TRANSACTION CRUD
function initModalEvents() {
  const modal = document.getElementById('modalTransaction');
  const btnOpen = document.getElementById('btnOpenModal');
  const btnClose = document.getElementById('btnCloseModal');
  const form = document.getElementById('formTransaction');

  btnOpen.addEventListener('click', () => openModal());
  btnClose.addEventListener('click', () => closeModal());
  modal.addEventListener('click', (e) => {
    if (e.target === modal) closeModal();
  });

  form.addEventListener('submit', (e) => {
    e.preventDefault();
    saveTransactionFromForm();
  });
}

function openModal(editingTx = null) {
  const modal = document.getElementById('modalTransaction');
  const title = document.getElementById('modalTitle');
  const idInput = document.getElementById('txId');
  const fechaInput = document.getElementById('txFecha');
  const tipoInput = document.getElementById('txTipo');
  const catInput = document.getElementById('txCategoria');
  const descInput = document.getElementById('txConcepto');
  const mpInput = document.getElementById('txMedioPago');
  const montoInput = document.getElementById('txMonto');
  const notasInput = document.getElementById('txNotas');

  if (editingTx) {
    title.textContent = 'Editar Movimiento';
    idInput.value = editingTx.id;
    fechaInput.value = editingTx.fecha;
    tipoInput.value = editingTx.tipo;
    initCategoryDropdown();
    catInput.value = editingTx.categoria;
    descInput.value = editingTx.concepto;
    mpInput.value = editingTx.medioPago;
    montoInput.value = editingTx.monto;
    notasInput.value = editingTx.notas || '';
  } else {
    title.textContent = 'Registrar Nuevo Movimiento';
    idInput.value = '';
    // Default today date in local format (YYYY-MM-DD)
    const today = new Date().toISOString().split('T')[0];
    fechaInput.value = today;
    tipoInput.value = 'Gasto Variable';
    initCategoryDropdown();
    descInput.value = '';
    montoInput.value = '';
    notasInput.value = '';
  }

  modal.classList.remove('hidden');
}

function closeModal() {
  document.getElementById('modalTransaction').classList.add('hidden');
}

function saveTransactionFromForm() {
  const idVal = document.getElementById('txId').value;
  const fechaVal = document.getElementById('txFecha').value;
  const tipoVal = document.getElementById('txTipo').value;
  const catVal = document.getElementById('txCategoria').value;
  const descVal = document.getElementById('txConcepto').value.trim();
  const mpVal = document.getElementById('txMedioPago').value;
  const montoVal = parseFloat(document.getElementById('txMonto').value) || 0;
  const notasVal = document.getElementById('txNotas').value.trim();

  const semanaVal = getWeekOfMonth(fechaVal);

  if (idVal) {
    // Edit existing
    const idx = state.transactions.findIndex(t => t.id == idVal);
    if (idx !== -1) {
      state.transactions[idx] = {
        id: parseInt(idVal, 10),
        fecha: fechaVal,
        semana: semanaVal,
        tipo: tipoVal,
        categoria: catVal,
        concepto: descVal,
        medioPago: mpVal,
        monto: montoVal,
        notas: notasVal
      };
    }
  } else {
    // Create new
    const nextId = state.transactions.length > 0 
      ? Math.max(...state.transactions.map(t => t.id || 0)) + 1 
      : 1;

    state.transactions.unshift({
      id: nextId,
      fecha: fechaVal,
      semana: semanaVal,
      tipo: tipoVal,
      categoria: catVal,
      concepto: descVal,
      medioPago: mpVal,
      monto: montoVal,
      notas: notasVal
    });
  }

  saveLocal();
  closeModal();

  // Re-render
  renderDashboard();
  renderTransactionsList();
  renderAnnualConsolidated();

  // Trigger background GitHub auto-sync if token is set
  if (state.githubToken) {
    showToast('⏳ Subiendo movimiento a GitHub...');
    syncPushGitHub(true);
  } else {
    showToast('💾 Guardado en este dispositivo. Para sincronizar con la PC, ingresa tu Token en Respaldos.', 'warning');
  }
}

function deleteTransaction(id) {
  if (confirm('¿Estás seguro de que deseas eliminar este movimiento?')) {
    state.transactions = state.transactions.filter(t => t.id !== id);
    saveLocal();
    renderDashboard();
    renderTransactionsList();
    renderAnnualConsolidated();

    if (state.githubToken) {
      syncPushGitHub(true);
    }
  }
}

// 6. DASHBOARD CALCULATIONS & RENDERING
function renderDashboard() {
  const curMonthObj = CALENDAR_MONTHS.find(m => m.name === state.selectedMonth) || CALENDAR_MONTHS[1];
  document.getElementById('lblSelectedMonthName').textContent = curMonthObj.name;

  const startD = curMonthObj.start;
  const endD = curMonthObj.end;

  // Filter transactions
  const prevTransactions = state.transactions.filter(t => t.fecha < startD);
  const curTransactions = state.transactions.filter(t => t.fecha >= startD && t.fecha <= endD);

  // 1. Dynamic Saldo Inicial = Base + previous net flows
  let saldoInicial = state.saldoInicialBase;
  prevTransactions.forEach(t => {
    if (t.tipo === 'Ingreso') saldoInicial += t.monto;
    else saldoInicial -= t.monto;
  });

  // 2. Current Month Totals
  let totalIngresos = 0;
  let totalFijos = 0;
  let totalVariables = 0;
  let totalHormiga = 0;
  let totalAhorro = 0;

  curTransactions.forEach(t => {
    if (t.tipo === 'Ingreso') totalIngresos += t.monto;
    else if (t.tipo === 'Gasto Fijo') totalFijos += t.monto;
    else if (t.tipo === 'Gasto Variable') totalVariables += t.monto;
    else if (t.tipo === 'Gasto Hormiga') totalHormiga += t.monto;
    else if (t.tipo === 'Ahorro / Inversión') totalAhorro += t.monto;
  });

  const totalGastosOperativos = totalFijos + totalVariables + totalHormiga;
  const flujoNeto = totalIngresos - totalGastosOperativos - totalAhorro;
  const saldoFinal = saldoInicial + flujoNeto;

  // Update KPI Cards DOM
  document.getElementById('kpiSaldoInicial').textContent = formatMoney(saldoInicial);
  document.getElementById('kpiIngresos').textContent = formatMoney(totalIngresos);
  document.getElementById('kpiGastos').textContent = formatMoney(totalGastosOperativos);
  document.getElementById('kpiHormiga').textContent = formatMoney(totalHormiga);
  document.getElementById('kpiFlujoNeto').textContent = formatMoney(flujoNeto);
  document.getElementById('kpiSaldoFinal').textContent = formatMoney(saldoFinal);

  // 3. Weekly Table
  const weeklyBody = document.getElementById('tableWeeklyBody');
  weeklyBody.innerHTML = '';

  const semanas = ['Semana 1', 'Semana 2', 'Semana 3', 'Semana 4', 'Semana 5'];
  let runningBalance = saldoInicial;
  const weeklyChartLabels = [];
  const weeklyChartIngresos = [];
  const weeklyChartGastos = [];

  semanas.forEach(sem => {
    const semTxs = curTransactions.filter(t => t.semana === sem);
    let sIng = 0;
    let sGas = 0;
    let sHor = 0;
    let sAhr = 0;

    semTxs.forEach(t => {
      if (t.tipo === 'Ingreso') sIng += t.monto;
      else if (t.tipo === 'Gasto Hormiga') { sGas += t.monto; sHor += t.monto; }
      else if (t.tipo === 'Ahorro / Inversión') sAhr += t.monto;
      else sGas += t.monto;
    });

    const sNet = sIng - sGas - sAhr;
    runningBalance += sNet;

    weeklyChartLabels.push(sem);
    weeklyChartIngresos.push(sIng);
    weeklyChartGastos.push(sGas);

    const row = document.createElement('tr');
    row.className = 'hover:bg-slate-50 font-medium';
    row.innerHTML = `
      <td class="py-2 px-2.5 font-bold text-slate-700">${sem}</td>
      <td class="py-2 px-2.5 text-right text-emerald-600 font-semibold">${sIng > 0 ? formatMoney(sIng) : '-'}</td>
      <td class="py-2 px-2.5 text-right text-rose-600">${sGas > 0 ? formatMoney(sGas) : '-'}</td>
      <td class="py-2 px-2.5 text-right text-amber-700">${sHor > 0 ? formatMoney(sHor) : '-'}</td>
      <td class="py-2 px-2.5 text-right ${sNet >= 0 ? 'text-blue-700 font-bold' : 'text-rose-700 font-bold'}">${formatMoney(sNet)}</td>
      <td class="py-2 px-2.5 text-right font-bold text-slate-800">${formatMoney(runningBalance)}</td>
    `;
    weeklyBody.appendChild(row);
  });

  // Render Charts
  renderWeeklyChart(weeklyChartLabels, weeklyChartIngresos, weeklyChartGastos);
  renderDistributionChart(totalFijos, totalVariables, totalHormiga, totalAhorro);

  // 4. Category Budget vs Actual
  renderCategoryBudget(curTransactions);
}

// Chart.js Weekly Bar Chart
function renderWeeklyChart(labels, ingresos, gastos) {
  const ctx = document.getElementById('chartWeekly').getContext('2d');
  if (chartWeeklyInstance) chartWeeklyInstance.destroy();

  chartWeeklyInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        { label: 'Ingresos', data: ingresos, backgroundColor: '#27AE60', borderRadius: 6 },
        { label: 'Gastos', data: gastos, backgroundColor: '#C0392B', borderRadius: 6 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'top', labels: { boxWidth: 12, font: { size: 10 } } } },
      scales: {
        y: { ticks: { callback: v => '$' + v.toLocaleString('es-AR') }, grid: { color: '#F1F5F9' } },
        x: { grid: { display: false } }
      }
    }
  });
}

// Chart.js Distribution Donut Chart
function renderDistributionChart(fijos, variables, hormiga, ahorro) {
  const ctx = document.getElementById('chartDistribution').getContext('2d');
  if (chartDistInstance) chartDistInstance.destroy();

  const total = fijos + variables + hormiga + ahorro;
  const dataVals = total > 0 ? [fijos, variables, hormiga, ahorro] : [1, 0, 0, 0];
  const bgColors = total > 0 
    ? ['#2C4C7C', '#2980B9', '#D35400', '#27AE60'] 
    : ['#E2E8F0', '#E2E8F0', '#E2E8F0', '#E2E8F0'];

  chartDistInstance = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Fijos', 'Variables', 'Hormiga', 'Ahorro'],
      datasets: [{ data: dataVals, backgroundColor: bgColors, borderWidth: 2, hoverOffset: 4 }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '68%',
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } }
      }
    }
  });
}

// Category Budget Progress Bars
function renderCategoryBudget(curTransactions) {
  const container = document.getElementById('categoryBudgetList');
  container.innerHTML = '';

  const expenseCats = CATEGORIES.filter(c => c.group !== 'Ingreso');

  expenseCats.forEach(cat => {
    const spent = curTransactions
      .filter(t => t.categoria === cat.name)
      .reduce((acc, t) => acc + t.monto, 0);

    const budget = cat.budget || 0;
    const diff = budget - spent;
    const pct = budget > 0 ? Math.min(100, Math.round((spent / budget) * 100)) : (spent > 0 ? 100 : 0);

    let badgeClass = 'bg-emerald-100 text-emerald-800';
    let badgeText = '✅ Control';
    let barColor = 'bg-emerald-500';

    if (spent > budget && budget > 0) {
      badgeClass = 'bg-rose-100 text-rose-800';
      badgeText = '⚠️ Excedido';
      barColor = 'bg-rose-500';
    } else if (pct >= 85) {
      badgeClass = 'bg-amber-100 text-amber-800';
      badgeText = '🟡 Alerta +85%';
      barColor = 'bg-amber-500';
    }

    const item = document.createElement('div');
    item.className = 'p-3 rounded-xl bg-slate-50 border border-slate-200/60 text-xs space-y-1.5';
    item.innerHTML = `
      <div class="flex items-center justify-between">
        <span class="font-bold text-slate-800 truncate max-w-[200px]">${cat.name}</span>
        <span class="text-[10px] font-bold px-2 py-0.5 rounded-full ${badgeClass}">${badgeText}</span>
      </div>
      <div class="flex items-center justify-between text-slate-500 text-[11px]">
        <span>Gastado: <b>${formatMoney(spent)}</b></span>
        <span>Presupuesto: ${formatMoney(budget)}</span>
      </div>
      <div class="w-full bg-slate-200 h-2 rounded-full overflow-hidden">
        <div class="${barColor} h-full transition-all duration-500" style="width: ${pct}%"></div>
      </div>
    `;
    container.appendChild(item);
  });
}

// 7. TRANSACTIONS LIST RENDERING
function renderTransactionsList() {
  const container = document.getElementById('transactionsContainer');
  container.innerHTML = '';

  const searchQ = document.getElementById('searchInput').value.toLowerCase();
  const filterType = document.getElementById('filterType').value;

  let filtered = state.transactions;
  if (filterType !== 'all') {
    filtered = filtered.filter(t => t.tipo === filterType);
  }
  if (searchQ) {
    filtered = filtered.filter(t => 
      t.concepto.toLowerCase().includes(searchQ) || 
      t.categoria.toLowerCase().includes(searchQ) ||
      (t.notas && t.notas.toLowerCase().includes(searchQ))
    );
  }

  if (filtered.length === 0) {
    container.innerHTML = `
      <div class="p-8 text-center text-slate-400 text-xs">
        <p class="text-3xl mb-2">🔍</p>
        <p>No se encontraron movimientos registrados con ese filtro.</p>
      </div>
    `;
    return;
  }

  filtered.forEach(tx => {
    let icon = '📦';
    let typeBadge = 'bg-blue-100 text-blue-700';

    if (tx.tipo === 'Ingreso') { icon = '💵'; typeBadge = 'bg-emerald-100 text-emerald-700'; }
    else if (tx.tipo === 'Gasto Hormiga') { icon = '🐜'; typeBadge = 'bg-amber-100 text-amber-800'; }
    else if (tx.tipo === 'Gasto Fijo') { icon = '🔒'; typeBadge = 'bg-rose-100 text-rose-700'; }
    else if (tx.tipo === 'Ahorro / Inversión') { icon = '🏦'; typeBadge = 'bg-indigo-100 text-indigo-700'; }

    const card = document.createElement('div');
    card.className = 'bg-white p-3.5 rounded-2xl border border-slate-200/80 shadow-sm flex items-center justify-between text-xs transition hover:border-blue-300';
    card.innerHTML = `
      <div class="flex items-center space-x-3 truncate">
        <div class="w-10 h-10 rounded-xl bg-slate-100 flex items-center justify-center text-lg flex-shrink-0">
          ${icon}
        </div>
        <div class="truncate">
          <div class="flex items-center space-x-1.5 truncate">
            <span class="font-bold text-slate-800 truncate">${tx.concepto}</span>
            <span class="text-[9px] font-bold px-1.5 py-0.5 rounded-md ${typeBadge}">${tx.tipo}</span>
          </div>
          <p class="text-[11px] text-slate-500 truncate">${tx.categoria} · <span class="text-slate-400">${tx.fecha}</span></p>
          ${tx.notas ? `<p class="text-[10px] text-slate-400 italic truncate">${tx.notas}</p>` : ''}
        </div>
      </div>

      <div class="flex items-center space-x-3 flex-shrink-0 pl-2">
        <div class="text-right">
          <span class="font-bold text-sm ${tx.tipo === 'Ingreso' ? 'text-emerald-600' : 'text-slate-800'}">
            ${tx.tipo === 'Ingreso' ? '+' : '-'}${formatMoney(tx.monto)}
          </span>
          <p class="text-[9px] text-slate-400">${tx.semana}</p>
        </div>
        <button class="btn-del text-slate-300 hover:text-rose-500 p-1 transition" data-id="${tx.id}">
          🗑️
        </button>
      </div>
    `;

    card.querySelector('.btn-del').addEventListener('click', (e) => {
      e.stopPropagation();
      deleteTransaction(tx.id);
    });

    card.addEventListener('click', () => {
      openModal(tx);
    });

    container.appendChild(card);
  });
}

document.getElementById('searchInput').addEventListener('input', renderTransactionsList);
document.getElementById('filterType').addEventListener('change', renderTransactionsList);

// 8. ANNUAL CONSOLIDATED MATRIX RENDERING
function renderAnnualConsolidated() {
  const table = document.getElementById('tableAnnual');
  if (!table) return;

  // Header Row
  let headHtml = `
    <thead>
      <tr class="bg-slate-800 text-white font-semibold">
        <th class="py-2.5 px-3 sticky left-0 bg-slate-800 z-10">Concepto / Categoría</th>
        <th class="py-2.5 px-2">Tipo</th>
  `;
  CALENDAR_MONTHS.forEach(m => {
    headHtml += `<th class="py-2.5 px-3 text-right">${m.short}</th>`;
    if (m.name === 'Diciembre 2026') headHtml += `<th class="py-2.5 px-3 text-right bg-slate-900 font-bold">TOTAL Q4 2026</th>`;
    if (m.name === 'Diciembre 2027') headHtml += `<th class="py-2.5 px-3 text-right bg-slate-900 font-bold">TOTAL AÑO 2027</th>`;
  });
  headHtml += `</tr></thead>`;

  // Summary Rows
  const summaryDefs = [
    { label: 'Saldo Inicial', type: 'balance' },
    { label: '(+) Ingresos', type: 'Ingreso' },
    { label: '(-) Gastos Fijos', type: 'Gasto Fijo' },
    { label: '(-) Gastos Variables', type: 'Gasto Variable' },
    { label: '(-) Gastos Hormiga', type: 'Gasto Hormiga' },
    { label: '(-) Ahorro e Inversión', type: 'Ahorro / Inversión' },
    { label: '(=) Flujo Neto', type: 'neto' },
    { label: '(=) Saldo Final', type: 'final' }
  ];

  let bodyHtml = `<tbody class="divide-y divide-slate-200">`;

  // Pre-calculate all monthly numbers
  const monthlyData = {};
  let rollingBalance = state.saldoInicialBase;

  CALENDAR_MONTHS.forEach(m => {
    const mStart = m.start;
    const mEnd = m.end;
    const mTx = state.transactions.filter(t => t.fecha >= mStart && t.fecha <= mEnd);

    let ing = 0, fix = 0, var_ = 0, hor = 0, ahr = 0;
    mTx.forEach(t => {
      if (t.tipo === 'Ingreso') ing += t.monto;
      else if (t.tipo === 'Gasto Fijo') fix += t.monto;
      else if (t.tipo === 'Gasto Variable') var_ += t.monto;
      else if (t.tipo === 'Gasto Hormiga') hor += t.monto;
      else if (t.tipo === 'Ahorro / Inversión') ahr += t.monto;
    });

    const net = ing - (fix + var_ + hor) - ahr;
    const sIni = rollingBalance;
    rollingBalance += net;

    monthlyData[m.name] = {
      saldoIni: sIni,
      ingresos: ing,
      fijos: fix,
      variables: var_,
      hormiga: hor,
      ahorro: ahr,
      flujoNeto: net,
      saldoFin: rollingBalance
    };
  });

  // Render Cash Flow Summary Rows
  summaryDefs.forEach(def => {
    const isSpecial = ['neto', 'final', 'balance'].includes(def.type);
    bodyHtml += `<tr class="${isSpecial ? 'bg-slate-100 font-bold' : 'bg-white'}">`;
    bodyHtml += `<td class="py-2 px-3 sticky left-0 ${isSpecial ? 'bg-slate-100' : 'bg-white'} font-bold z-10">${def.label}</td>`;
    bodyHtml += `<td class="py-2 px-2 text-slate-400 text-[10px] uppercase">${def.type}</td>`;

    let sumQ4 = 0;
    let sum27 = 0;

    CALENDAR_MONTHS.forEach(m => {
      const d = monthlyData[m.name];
      let val = 0;
      if (def.type === 'balance') val = d.saldoIni;
      else if (def.type === 'Ingreso') val = d.ingresos;
      else if (def.type === 'Gasto Fijo') val = d.fijos;
      else if (def.type === 'Gasto Variable') val = d.variables;
      else if (def.type === 'Gasto Hormiga') val = d.hormiga;
      else if (def.type === 'Ahorro / Inversión') val = d.ahorro;
      else if (def.type === 'neto') val = d.flujoNeto;
      else if (def.type === 'final') val = d.saldoFin;

      bodyHtml += `<td class="py-2 px-3 text-right ${val < 0 ? 'text-rose-600' : ''}">${val !== 0 ? formatMoney(val) : '-'}</td>`;

      if (m.year === 2026) sumQ4 += val;
      if (m.year === 2027) sum27 += val;

      if (m.name === 'Diciembre 2026') {
        const q4Val = def.type === 'balance' ? monthlyData['Octubre 2026'].saldoIni : (def.type === 'final' ? monthlyData['Diciembre 2026'].saldoFin : sumQ4);
        bodyHtml += `<td class="py-2 px-3 text-right bg-slate-200/80 font-bold">${formatMoney(q4Val)}</td>`;
      }
      if (m.name === 'Diciembre 2027') {
        const y27Val = def.type === 'balance' ? monthlyData['Enero 2027'].saldoIni : (def.type === 'final' ? monthlyData['Diciembre 2027'].saldoFin : sum27);
        bodyHtml += `<td class="py-2 px-3 text-right bg-slate-200/80 font-bold">${formatMoney(y27Val)}</td>`;
      }
    });

    bodyHtml += `</tr>`;
  });

  bodyHtml += `</tbody>`;
  table.innerHTML = headHtml + bodyHtml;
}

// 9. GITHUB REST API SYNCHRONIZATION
function initBackupEvents() {
  const ghTokenInput = document.getElementById('ghToken');
  const btnSaveToken = document.getElementById('btnSaveToken');
  const btnPush = document.getElementById('btnSyncPush');
  const btnPull = document.getElementById('btnSyncPull');
  const btnQuickSync = document.getElementById('btnQuickSync');
  const btnExportExcel = document.getElementById('btnExportExcel');
  const btnExportAnnualExcel = document.getElementById('btnExportAnnualExcel');
  const fileImportExcel = document.getElementById('fileImportExcel');
  const btnExportJSON = document.getElementById('btnExportJSON');
  const fileImportJSON = document.getElementById('fileImportJSON');

  if (state.githubToken) {
    ghTokenInput.value = state.githubToken;
    updateSyncStatus('GitHub listo', true);
  }

  btnSaveToken.addEventListener('click', async () => {
    const tok = ghTokenInput.value.trim();
    await verifyAndSaveToken(tok);
  });

  btnPush.addEventListener('click', () => syncPushGitHub(false));
  btnPull.addEventListener('click', () => syncPullGitHub());
  btnQuickSync.addEventListener('click', () => {
    if (!state.githubToken) {
      alert('Configura primero tu GitHub Token en la pestaña "Respaldos".');
      document.querySelector('[data-tab="respaldos"]').click();
    } else {
      syncPushGitHub(false);
    }
  });

  btnExportExcel.addEventListener('click', () => exportToExcelWorkbook());
  if (btnExportAnnualExcel) btnExportAnnualExcel.addEventListener('click', () => exportToExcelWorkbook());
  fileImportExcel.addEventListener('change', handleExcelImport);

  btnExportJSON.addEventListener('click', exportToJSON);
  fileImportJSON.addEventListener('change', handleJSONImport);
}

function updateSyncStatus(msg, isSuccess = true) {
  const label = document.getElementById('syncLabel');
  const dot = document.getElementById('syncIndicator');
  const statusMsg = document.getElementById('syncStatusMsg');

  if (label) label.textContent = msg;
  if (dot) {
    dot.className = isSuccess 
      ? 'w-2 h-2 rounded-full bg-emerald-400' 
      : 'w-2 h-2 rounded-full bg-rose-400';
  }
  if (statusMsg) statusMsg.textContent = `Estado: ${msg} (${new Date().toLocaleTimeString()})`;
}

// Verify token by checking access to the repository
async function verifyAndSaveToken(token) {
  token = (token || '').trim();
  if (!token) {
    localStorage.removeItem('gh_token');
    state.githubToken = '';
    updateSyncStatus('Sin token', false);
    alert('Token eliminado.');
    return;
  }

  updateSyncStatus('Verificando...', true);

  try {
    const url = `https://api.github.com/repos/${state.githubRepo}?t=${Date.now()}`;
    const res = await fetch(url, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/vnd.github.v3+json'
      },
      cache: 'no-store'
    });

    if (res.ok) {
      const repoData = await res.json();
      state.githubToken = token;
      localStorage.setItem('gh_token', token);
      updateSyncStatus('GitHub conectado', true);
      alert(`✅ ¡Token verificado con éxito!\n\nConectado a ${repoData.full_name} (${repoData.private ? 'Repositorio Privado' : 'Público'}).\nYa puedes subir y bajar tus datos.`);
    } else {
      const err = await res.json().catch(() => ({}));
      updateSyncStatus('Error Token', false);
      if (res.status === 401) {
        alert('❌ Error 401 (Credenciales inválidas):\nEl token ingresado no es válido o está vencido. Asegúrate de copiarlo completo sin espacios.');
      } else if (res.status === 404) {
        alert(`❌ Error 404 (No encontrado):\nGitHub no permite acceder al repositorio "${state.githubRepo}".\n\nComo tu repositorio es Privado, esto significa que el token no tiene el permiso necesario:\n• Si creaste un Token Classic: debes marcar la casilla "repo".\n• Si creaste un Token Fine-Grained: debes seleccionar el repositorio "FlujoDeCajaPersonal" y dar permiso de Contents (Read & write).`);
      } else {
        alert(`❌ Error ${res.status}: ${err.message || 'No se pudo conectar a GitHub'}`);
      }
    }
  } catch (e) {
    console.error('Error verifying token:', e);
    updateSyncStatus('Error Red', false);
    alert('❌ Error de conexión al verificar el token con GitHub.');
  }
}

// Push local data to GitHub repository
async function syncPushGitHub(isSilent = false) {
  const token = (state.githubToken || '').trim();
  if (!token) {
    if (!isSilent) alert('Configura y guarda tu token de GitHub en la pestaña "Respaldos" primero.');
    return;
  }

  updateSyncStatus('Subiendo...', true);

  try {
    const url = `https://api.github.com/repos/${state.githubRepo}/contents/data/transacciones.json`;
    
    // 1. Get current SHA if file exists
    let sha = '';
    const getRes = await fetch(`${url}?t=${Date.now()}`, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/vnd.github.v3+json'
      },
      cache: 'no-store'
    });

    if (getRes.ok) {
      const fileData = await getRes.json();
      sha = fileData.sha;
    } else if (getRes.status === 401) {
      updateSyncStatus('Error 401', false);
      if (!isSilent) alert('❌ Error 401 (Token inválido o expirado). Revisa el token en la pestaña Respaldos.');
      return;
    }

    // 2. Prepare payload
    const payloadData = {
      version: '1.0',
      lastUpdated: new Date().toISOString(),
      saldoInicialBase: state.saldoInicialBase,
      transacciones: state.transactions
    };
    const jsonStr = JSON.stringify(payloadData, null, 2);

    // Safe UTF-8 to base64 encoding
    const utf8Bytes = new TextEncoder().encode(jsonStr);
    let binaryStr = '';
    for (let i = 0; i < utf8Bytes.length; i++) {
      binaryStr += String.fromCharCode(utf8Bytes[i]);
    }
    const contentBase64 = btoa(binaryStr);

    // 3. Put commit to GitHub
    const putRes = await fetch(url, {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'Accept': 'application/vnd.github.v3+json'
      },
      body: JSON.stringify({
        message: `Backup: Sincronización de movimientos [skip ci]`,
        content: contentBase64,
        sha: sha || undefined
      })
    });

    if (putRes.ok) {
      updateSyncStatus('GitHub al día', true);
      showToast('✅ Sincronizado con GitHub con éxito', 'success');
      if (!isSilent) alert('✅ Respaldo exitoso en GitHub.');
    } else {
      const err = await putRes.json().catch(() => ({}));
      updateSyncStatus('Error GitHub', false);
      showToast(`⚠️ No se pudo subir a GitHub (${putRes.status})`, 'error');
      if (!isSilent) alert(`❌ Error ${putRes.status} al respaldar en GitHub: ${err.message || 'Error desconocido'}`);
    }
  } catch (err) {
    console.error('GitHub Push error:', err);
    updateSyncStatus('Error Red', false);
    showToast('❌ Error de conexión al sincronizar con GitHub', 'error');
    if (!isSilent) alert('❌ Error de conexión al sincronizar con GitHub.');
  }
}

// Pull remote data from GitHub repository
async function syncPullGitHub(isSilent = false) {
  const token = (state.githubToken || '').trim();
  if (!token) {
    if (!isSilent) alert('Ingresa y guarda tu token de GitHub en la pestaña "Respaldos" primero.');
    return;
  }

  updateSyncStatus('Descargando...', true);

  try {
    const url = `https://api.github.com/repos/${state.githubRepo}/contents/data/transacciones.json?t=${Date.now()}`;
    const res = await fetch(url, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/vnd.github.v3+json'
      },
      cache: 'no-store'
    });

    if (res.ok) {
      const fileData = await res.json();
      // Handle base64 decoding with utf-8 and clean newlines
      const cleanBase64 = fileData.content.replace(/\s/g, '');
      const binaryStr = atob(cleanBase64);
      const bytes = new Uint8Array(binaryStr.length);
      for (let i = 0; i < binaryStr.length; i++) {
        bytes[i] = binaryStr.charCodeAt(i);
      }
      const decodedJson = new TextDecoder('utf-8').decode(bytes);
      const parsed = JSON.parse(decodedJson);

      if (parsed.transacciones && Array.isArray(parsed.transacciones)) {
        state.transactions = parsed.transacciones;
        state.saldoInicialBase = parsed.saldoInicialBase || 0;
        saveLocal();
        renderDashboard();
        renderTransactionsList();
        renderAnnualConsolidated();
        updateSyncStatus('Descargado', true);
        if (!isSilent) {
          alert(`✅ Sincronización completa: ${state.transactions.length} movimientos cargados desde GitHub.`);
        } else {
          showToast(`✅ ${state.transactions.length} movimientos sincronizados desde GitHub`, 'success');
        }
      } else {
        if (!isSilent) alert('El archivo en GitHub no tiene el formato esperado.');
      }
    } else {
      const err = await res.json().catch(() => ({}));
      updateSyncStatus('Error GitHub', false);
      if (res.status === 401) {
        if (!isSilent) alert('❌ Error 401 (Credenciales inválidas):\nEl token ingresado no es válido o está vencido. Genera uno nuevo en github.com/settings/tokens.');
      } else if (res.status === 404) {
        if (!isSilent) alert(`❌ Error 404 (No encontrado):\nGitHub no permite acceder al archivo en el repositorio "${state.githubRepo}".\n\nCausas más frecuentes:\n1. Tu repositorio es Privado y el token no tiene el permiso "repo" marcado.\n2. Si creaste un token Fine-Grained, asegúrate de haberle dado acceso a este repositorio con permiso "Contents: Read & write".`);
      } else {
        if (!isSilent) alert(`❌ Error ${res.status}: ${err.message || 'No se pudo descargar desde GitHub'}`);
      }
    }
  } catch (err) {
    console.error('GitHub Pull error:', err);
    updateSyncStatus('Error Red', false);
    if (!isSilent) alert('❌ Error de conexión al descargar datos desde GitHub.');
  }
}

// 10. EXCEL EXPORT & IMPORT (SHEETJS)
function exportToExcelWorkbook() {
  if (typeof XLSX === 'undefined') {
    alert('La librería de Excel no se ha cargado. Verifica tu conexión a internet.');
    return;
  }

  const wb = XLSX.utils.book_new();

  // Sheet 1: Registro Diario
  const dailyData = [
    ['#', 'Fecha', 'Semana', 'Tipo de Movimiento', 'Categoría', 'Concepto / Detalle Específico', 'Medio de Pago', 'Monto ($)', 'Notas / Justificación']
  ];
  state.transactions.forEach((t, i) => {
    dailyData.push([
      i + 1,
      t.fecha,
      t.semana,
      t.tipo,
      t.categoria,
      t.concepto,
      t.medioPago,
      t.monto,
      t.notas || ''
    ]);
  });
  const wsDaily = XLSX.utils.aoa_to_sheet(dailyData);
  XLSX.utils.book_append_sheet(wb, wsDaily, 'Registro Diario');

  // Sheet 2: Consolidado Anual
  const annualData = [
    ['Concepto / Categoría', 'Grupo', ...CALENDAR_MONTHS.map(m => m.short)]
  ];
  const summaryDefs = ['Ingreso', 'Gasto Fijo', 'Gasto Variable', 'Gasto Hormiga', 'Ahorro / Inversión'];
  summaryDefs.forEach(type => {
    const row = [type, type];
    CALENDAR_MONTHS.forEach(m => {
      const sum = state.transactions
        .filter(t => t.tipo === type && t.fecha >= m.start && t.fecha <= m.end)
        .reduce((acc, t) => acc + t.monto, 0);
      row.push(sum);
    });
    annualData.push(row);
  });
  const wsAnnual = XLSX.utils.aoa_to_sheet(annualData);
  XLSX.utils.book_append_sheet(wb, wsAnnual, 'Consolidado Anual');

  // Sheet 3: Categorías y Presupuestos
  const catData = [
    ['Categoría', 'Grupo', 'Presupuesto Mensual']
  ];
  CATEGORIES.forEach(c => {
    catData.push([c.name, c.group, c.budget]);
  });
  const wsCat = XLSX.utils.aoa_to_sheet(catData);
  XLSX.utils.book_append_sheet(wb, wsCat, 'Configuracion');

  // Download
  XLSX.writeFile(wb, 'Flujo_de_Caja_Personal.xlsx');
}

function handleExcelImport(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      
      const sheetName = workbook.SheetNames.includes('Registro Diario') ? 'Registro Diario' : workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // Parse rows from index 1
      const newTransactions = [];
      for (let r = 1; r < json.length; r++) {
        const row = json[r];
        if (row && row[1]) { // If date exists
          const dateStr = typeof row[1] === 'string' ? row[1].slice(0, 10) : new Date().toISOString().slice(0, 10);
          newTransactions.push({
            id: r,
            fecha: dateStr,
            semana: row[2] || getWeekOfMonth(dateStr),
            tipo: row[3] || 'Gasto Variable',
            categoria: row[4] || 'Varios',
            concepto: row[5] || 'Sin detalle',
            medioPago: row[6] || 'Efectivo',
            monto: parseFloat(row[7]) || 0,
            notas: row[8] || ''
          });
        }
      }

      if (newTransactions.length > 0) {
        if (confirm(`Se encontraron ${newTransactions.length} registros en el Excel. ¿Deseas importarlos reemplazando los actuales?`)) {
          state.transactions = newTransactions;
          saveLocal();
          renderDashboard();
          renderTransactionsList();
          renderAnnualConsolidated();
          alert('✅ Importación exitosa desde Excel.');
          if (state.githubToken) syncPushGitHub(true);
        }
      } else {
        alert('No se encontraron transacciones legibles en el archivo.');
      }
    } catch (err) {
      console.error('Error importing Excel:', err);
      alert('Error al leer el archivo de Excel.');
    }
  };
  reader.readAsArrayBuffer(file);
}

// 11. JSON BACKUP EXPORT & IMPORT
function exportToJSON() {
  const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(state.transactions, null, 2));
  const downloadAnchor = document.createElement('a');
  downloadAnchor.setAttribute("href", dataStr);
  downloadAnchor.setAttribute("download", `flujo_caja_${new Date().toISOString().slice(0, 10)}.json`);
  document.body.appendChild(downloadAnchor);
  downloadAnchor.click();
  downloadAnchor.remove();
}

function handleJSONImport(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const parsed = JSON.parse(evt.target.result);
      if (Array.isArray(parsed)) {
        state.transactions = parsed;
        saveLocal();
        renderDashboard();
        renderTransactionsList();
        renderAnnualConsolidated();
        alert(`✅ Restaurado exitosamente: ${parsed.length} movimientos.`);
        if (state.githubToken) syncPushGitHub(true);
      }
    } catch (err) {
      alert('El archivo no tiene un formato JSON válido.');
    }
  };
  reader.readAsText(file);
}
