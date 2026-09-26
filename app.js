// SheetPulse Core State
const state = {
  workbook: null,
  currentSheetName: '',
  rawRows: [],
  filteredRows: [],
  columns: [],
  chartInstance: null,
  currentPage: 1,
  rowsPerPage: 10,
};

// Initialize icons and bindings on DOM load
document.addEventListener('DOMContentLoaded', () => {
  lucide.createIcons();
  bindEvents();
});

function bindEvents() {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');

  dropzone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleFile(e.target.files[0]);
  });

  // Drag & drop handlers
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('border-emerald-500');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('border-emerald-500');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('border-emerald-500');
    if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
  });

  // Demo datasets
  document.getElementById('loadSalesBtn').addEventListener('click', loadSampleSales);
  document.getElementById('loadSaaSBtn').addEventListener('click', loadSampleSaaS);

  // Search input
  document.getElementById('tableSearch').addEventListener('input', (e) => {
    const query = e.target.value.toLowerCase().trim();
    if (!query) {
      state.filteredRows = [...state.rawRows];
    } else {
      state.filteredRows = state.rawRows.filter(row =>
        Object.values(row).some(val => String(val).toLowerCase().includes(query))
      );
    }
    state.currentPage = 1;
    renderTable();
  });

  // Chart config listeners
  document.getElementById('chartType').addEventListener('change', updateChart);
  document.getElementById('chartDimX').addEventListener('change', updateChart);
  document.getElementById('chartDimY').addEventListener('change', updateChart);
  document.getElementById('chartAgg').addEventListener('change', updateChart);

  // Pagination buttons
  document.getElementById('prevPage').addEventListener('click', () => {
    if (state.currentPage > 1) {
      state.currentPage--;
      renderTable();
    }
  });

  document.getElementById('nextPage').addEventListener('click', () => {
    const maxPage = Math.ceil(state.filteredRows.length / state.rowsPerPage);
    if (state.currentPage < maxPage) {
      state.currentPage++;
      renderTable();
    }
  });

  // Chart export
  document.getElementById('downloadChartBtn').addEventListener('click', () => {
    if (!state.chartInstance) return;
    const a = document.createElement('a');
    a.href = state.chartInstance.toBase64Image();
    a.download = `${state.currentSheetName}-chart.png`;
    a.click();
  });
}

// Parse uploaded spreadsheet with SheetJS
function handleFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    state.workbook = XLSX.read(data, { type: 'array' });

    // Show metadata bar
    document.getElementById('fileMeta').classList.remove('hidden');
    document.getElementById('fileName').textContent = file.name;
    document.getElementById('fileSize').textContent = `${(file.size / 1024).toFixed(1)} KB`;

    renderSheetTabs();
    selectSheet(state.workbook.SheetNames[0]);
  };
  reader.readAsArrayBuffer(file);
}

// Generate tabs for multi-sheet workbooks
function renderSheetTabs() {
  const container = document.getElementById('sheetTabs');
  container.innerHTML = '';

  state.workbook.SheetNames.forEach((sheetName) => {
    const btn = document.createElement('button');
    btn.className = `px-3 py-1 rounded-lg text-xs font-medium transition ${
      sheetName === state.currentSheetName
        ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/40'
        : 'bg-slate-800 text-slate-400 hover:text-slate-200'
    }`;
    btn.textContent = sheetName;
    btn.onclick = () => selectSheet(sheetName);
    container.appendChild(btn);
  });
}

// Switch active sheet
function selectSheet(sheetName) {
  state.currentSheetName = sheetName;
  const worksheet = state.workbook.Sheets[sheetName];

  // Convert Sheet to JSON objects
  state.rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
  state.filteredRows = [...state.rawRows];

  if (state.rawRows.length > 0) {
    state.columns = Object.keys(state.rawRows[0]);
  } else {
    state.columns = [];
  }

  renderSheetTabs();
  updateKPIs();
  populateDropdowns();
  updateChart();
  renderTable();

  // Show UI sections
  document.getElementById('kpiRibbon').classList.remove('hidden');
  document.getElementById('analyticsSection').classList.remove('hidden');
  document.getElementById('tableSection').classList.remove('hidden');
}

// Calculate summary numbers
function updateKPIs() {
  document.getElementById('kpiTotalRows').textContent = state.rawRows.length.toLocaleString();
  document.getElementById('kpiTotalCols').textContent = state.columns.length;

  const numericCols = state.columns.filter(col =>
    state.rawRows.some(row => typeof row[col] === 'number')
  );

  if (numericCols.length > 0) {
    const primaryCol = numericCols[0];
    const total = state.rawRows.reduce((acc, row) => acc + (Number(row[primaryCol]) || 0), 0);
    const avg = total / (state.rawRows.length || 1);

    document.getElementById('kpiNumericLabel1').textContent = `Total ${primaryCol}`;
    document.getElementById('kpiNumericVal1').textContent = total.toLocaleString(undefined, { maximumFractionDigits: 1 });

    document.getElementById('kpiNumericLabel2').textContent = `Avg ${primaryCol}`;
    document.getElementById('kpiNumericVal2').textContent = avg.toLocaleString(undefined, { maximumFractionDigits: 1 });
  }
}

// Populate X and Y axis select fields
function populateDropdowns() {
  const dimX = document.getElementById('chartDimX');
  const dimY = document.getElementById('chartDimY');
  dimX.innerHTML = '';
  dimY.innerHTML = '';

  state.columns.forEach((col, idx) => {
    const optX = new Option(col, col);
    const optY = new Option(col, col);
    dimX.add(optX);
    dimY.add(optY);
  });

  // Pick numeric column for Y if available
  const numericCol = state.columns.find(col =>
    state.rawRows.some(row => typeof row[col] === 'number')
  );
  if (numericCol) dimY.value = numericCol;
}

// Aggregation & Chart.js rendering
function updateChart() {
  const chartType = document.getElementById('chartType').value;
  const xCol = document.getElementById('chartDimX').value;
  const yCol = document.getElementById('chartDimY').value;
  const aggType = document.getElementById('chartAgg').value;

  if (!xCol || !yCol) return;

  const grouped = {};
  state.filteredRows.forEach((row) => {
    const key = String(row[xCol] || 'Unknown');
    const val = Number(row[yCol]) || 0;

    if (!grouped[key]) grouped[key] = { sum: 0, count: 0 };
    grouped[key].sum += val;
    grouped[key].count += 1;
  });

  const labels = Object.keys(grouped).slice(0, 15); // Top 15 labels
  const data = labels.map((k) => {
    if (aggType === 'sum') return grouped[k].sum;
    if (aggType === 'avg') return grouped[k].sum / grouped[k].count;
    return grouped[k].count;
  });

  if (state.chartInstance) {
    state.chartInstance.destroy();
  }

  const ctx = document.getElementById('mainChart').getContext('2d');
  state.chartInstance = new Chart(ctx, {
    type: chartType,
    data: {
      labels: labels,
      datasets: [
        {
          label: `${aggType.toUpperCase()} of ${yCol}`,
          data: data,
          backgroundColor: [
            'rgba(16, 185, 129, 0.6)',
            'rgba(99, 102, 241, 0.6)',
            'rgba(245, 158, 11, 0.6)',
            'rgba(239, 68, 68, 0.6)',
            'rgba(14, 165, 233, 0.6)'
          ],
          borderColor: 'rgba(255, 255, 255, 0.1)',
          borderWidth: 1,
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#94a3b8' } }
      },
      scales: chartType === 'pie' || chartType === 'doughnut' ? {} : {
        x: { ticks: { color: '#64748b' }, grid: { color: '#1e293b' } },
        y: { ticks: { color: '#64748b' }, grid: { color: '#1e293b' } }
      }
    }
  });
}

// Render paginated data table
function renderTable() {
  const head = document.getElementById('tableHead');
  const body = document.getElementById('tableBody');
  head.innerHTML = '';
  body.innerHTML = '';

  if (state.columns.length === 0) return;

  // Render headers
  const headTr = document.createElement('tr');
  state.columns.forEach((col) => {
    const th = document.createElement('th');
    th.className = 'px-4 py-2.5 font-medium whitespace-nowrap';
    th.textContent = col;
    headTr.appendChild(th);
  });
  head.appendChild(headTr);

  // Pagination bounds
  const start = (state.currentPage - 1) * state.rowsPerPage;
  const end = start + state.rowsPerPage;
  const pageRows = state.filteredRows.slice(start, end);

  // Render rows
  pageRows.forEach((row) => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-800/40 transition-colors';
    state.columns.forEach((col) => {
      const td = document.createElement('td');
      td.className = 'px-4 py-2 whitespace-nowrap truncate max-w-xs';
      td.textContent = row[col] !== undefined ? row[col] : '';
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });

  // Update page indicators
  document.getElementById('pageInfo').textContent = 
    `Showing ${Math.min(start + 1, state.filteredRows.length)} - ${Math.min(end, state.filteredRows.length)} of ${state.filteredRows.length} records`;
  
  document.getElementById('prevPage').disabled = state.currentPage === 1;
  document.getElementById('nextPage').disabled = end >= state.filteredRows.length;
}

// Built-in Demo Data Loaders
function loadSampleSales() {
  const sales = [
    { OrderID: 'ORD-101', Region: 'North America', Category: 'Electronics', Sales: 1250, Units: 5 },
    { OrderID: 'ORD-102', Region: 'Europe', Category: 'Furniture', Sales: 840, Units: 2 },
    { OrderID: 'ORD-103', Region: 'Asia', Category: 'Office Supplies', Sales: 310, Units: 12 },
    { OrderID: 'ORD-104', Region: 'North America', Category: 'Furniture', Sales: 2200, Units: 8 },
    { OrderID: 'ORD-105', Region: 'Europe', Category: 'Electronics', Sales: 1750, Units: 4 },
    { OrderID: 'ORD-106', Region: 'Asia', Category: 'Electronics', Sales: 3400, Units: 10 }
  ];
  loadMockData('E-Commerce_Sample.xlsx', sales);
}

function loadSampleSaaS() {
  const saas = [
    { Plan: 'Starter', Users: 120, MRR: 3480, Churn: 2.1 },
    { Plan: 'Growth', Users: 85, MRR: 8500, Churn: 1.4 },
    { Plan: 'Enterprise', Users: 24, MRR: 28800, Churn: 0.5 },
    { Plan: 'Custom', Users: 9, MRR: 19000, Churn: 0.0 }
  ];
  loadMockData('SaaS_Metrics.xlsx', saas);
}

function loadMockData(filename, rows) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Overview');

  state.workbook = wb;
  document.getElementById('fileMeta').classList.remove('hidden');
  document.getElementById('fileName').textContent = filename;
  document.getElementById('fileSize').textContent = '2.4 KB (Mock)';

  renderSheetTabs();
  selectSheet('Overview');
}