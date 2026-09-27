tailwind.config = {
  theme: {
    extend: {
      fontFamily: {
        sans: ['Inter', 'sans-serif'],
        mono: ['JetBrains Mono', 'monospace'],
      },
      colors: {
        brand: {
          50: '#f5f3ff',
          100: '#ede9fe',
          200: '#ddd6fe',
          500: '#8b5cf6',
          600: '#4f46e5',
          700: '#4338ca',
        }
      }
    }
  }
};
// Universal App State
const AppState = {
  fileName: 'Dummy_Student_and_Employee_Data.xlsx',
  activeSheet: 'Students',
  sheets: {},
  currentRows: [],
  originalRows: [],
  filteredRows: [],
  columns: [],
  colMeta: {},
  grid: {
    page: 1,
    pageSize: 50,
    sortCol: null,
    sortDir: 'asc',
    searchQuery: '',
    columnFilter: 'all'
  },
  categoryChartInstance: null,
  studioChartInstance: null
};

// Preloaded Datasets
const SAMPLE_DATASETS = {
  sales: {
    fileName: 'Sales_Performance_Q3.xlsx',
    sheets: {
      'Sales': [
        { "OrderID": "SO-1001", "Region": "North America", "Category": "Electronics", "SubCategory": "Laptops", "Sales": 10500, "Profit": 2100, "Units": 14 },
        { "OrderID": "SO-1002", "Region": "Europe", "Category": "Furniture", "SubCategory": "Chairs", "Sales": 3000, "Profit": 450, "Units": 25 },
        { "OrderID": "SO-1003", "Region": "Asia Pacific", "Category": "Office Supplies", "SubCategory": "Paper", "Sales": 1200, "Profit": 320, "Units": 80 },
        { "OrderID": "SO-1004", "Region": "North America", "Category": "Furniture", "SubCategory": "Desks", "Sales": 3600, "Profit": 720, "Units": 8 },
        { "OrderID": "SO-1005", "Region": "Europe", "Category": "Electronics", "SubCategory": "Monitors", "Sales": 6600, "Profit": 1650, "Units": 30 },
        { "OrderID": "SO-1006", "Region": "Latin America", "Category": "Office Supplies", "SubCategory": "Binders", "Sales": 810, "Profit": 190, "Units": 45 },
        { "OrderID": "SO-1007", "Region": "Asia Pacific", "Category": "Electronics", "SubCategory": "Smartphones", "Sales": 13200, "Profit": 3300, "Units": 22 },
        { "OrderID": "SO-1008", "Region": "Middle East", "Category": "Furniture", "SubCategory": "Bookcases", "Sales": 3720, "Profit": 600, "Units": 12 },
        { "OrderID": "SO-1009", "Region": "North America", "Category": "Office Supplies", "SubCategory": "Storage", "Sales": 1400, "Profit": 280, "Units": 35 },
        { "OrderID": "SO-1010", "Region": "Europe", "Category": "Electronics", "SubCategory": "Accessories", "Sales": 2250, "Profit": 700, "Units": 90 },
        { "OrderID": "SO-1011", "Region": "Asia Pacific", "Category": "Furniture", "SubCategory": "Tables", "Sales": 3300, "Profit": 420, "Units": 6 },
        { "OrderID": "SO-1012", "Region": "Latin America", "Category": "Electronics", "SubCategory": "Laptops", "Sales": 7200, "Profit": 1400, "Units": 10 },
        { "OrderID": "SO-1013", "Region": "Middle East", "Category": "Electronics", "SubCategory": "Smartphones", "Sales": 8900, "Profit": 1900, "Units": 15 },
        { "OrderID": "SO-1014", "Region": "North America", "Category": "Furniture", "SubCategory": "Chairs", "Sales": 4500, "Profit": 850, "Units": 20 },
        { "OrderID": "SO-1015", "Region": "Europe", "Category": "Office Supplies", "SubCategory": "Paper", "Sales": 950, "Profit": 210, "Units": 50 },
        { "OrderID": "SO-1016", "Region": "Asia Pacific", "Category": "Office Supplies", "SubCategory": "Binders", "Sales": 1100, "Profit": 240, "Units": 60 },
        { "OrderID": "SO-1017", "Region": "Latin America", "Category": "Furniture", "SubCategory": "Desks", "Sales": 2900, "Profit": -45, "Units": 7 },
        { "OrderID": "SO-1018", "Region": "Middle East", "Category": "Office Supplies", "SubCategory": "Storage", "Sales": 1650, "Profit": 310, "Units": 28 }
      ],
      'Targets': [
        { "Region": "North America", "Target": 25000 },
        { "Region": "Europe", "Target": 18000 },
        { "Region": "Asia Pacific", "Target": 20000 }
      ]
    }
  },
  students: {
    fileName: 'Dummy_Student_and_Employee_Data.xlsx',
    sheets: {
      'Students': [
        { "StudentID": "STU-001", "Name": "Aarav Sharma", "Gender": "Male", "Class": "12-B", "Math": 52, "Physics": 46, "Chemistry": 92, "English": 62 },
        { "StudentID": "STU-002", "Name": "Aditya Sharma", "Gender": "Female", "Class": "10-B", "Math": 53, "Physics": 92, "Chemistry": 51, "English": 88 },
        { "StudentID": "STU-003", "Name": "Arjun Sharma", "Gender": "Male", "Class": "12-A", "Math": 50, "Physics": 82, "Chemistry": 72, "English": 47 },
        { "StudentID": "STU-004", "Name": "Kabir Sharma", "Gender": "Female", "Class": "10-A", "Math": 58, "Physics": 59, "Chemistry": 65, "English": 70 },
        { "StudentID": "STU-005", "Name": "Diya Patel", "Gender": "Female", "Class": "10-A", "Math": 88, "Physics": 78, "Chemistry": 85, "English": 90 },
        { "StudentID": "STU-006", "Name": "Ishaan Verma", "Gender": "Male", "Class": "11-A", "Math": 70, "Physics": 75, "Chemistry": 68, "English": 72 },
        { "StudentID": "STU-007", "Name": "Kavya Nair", "Gender": "Female", "Class": "11-B", "Math": 65, "Physics": 70, "Chemistry": 80, "English": 70 },
        { "StudentID": "STU-008", "Name": "Rohan Gupta", "Gender": "Male", "Class": "10-A", "Math": 92, "Physics": 85, "Chemistry": 89, "English": 78 },
        { "StudentID": "STU-009", "Name": "Ananya Iyer", "Gender": "Female", "Class": "12-A", "Math": 78, "Physics": 84, "Chemistry": 80, "English": 82 },
        { "StudentID": "STU-010", "Name": "Manish Kumar", "Gender": "Male", "Class": "10-B", "Math": 60, "Physics": 58, "Chemistry": 62, "English": 64 },
        { "StudentID": "STU-011", "Name": "Pooja Reddy", "Gender": "Female", "Class": "12-B", "Math": 85, "Physics": 80, "Chemistry": 76, "English": 84 },
        { "StudentID": "STU-012", "Name": "Varun Joshi", "Gender": "Male", "Class": "11-A", "Math": 45, "Physics": 52, "Chemistry": 50, "English": 58 },
        { "StudentID": "STU-013", "Name": "Sneha Kulkarni", "Gender": "Female", "Class": "11-B", "Math": 72, "Physics": 76, "Chemistry": 74, "English": 68 },
        { "StudentID": "STU-014", "Name": "Karan Mehra", "Gender": "Male", "Class": "10-A", "Math": 58, "Physics": 64, "Chemistry": 60, "English": 62 },
        { "StudentID": "STU-015", "Name": "Meera Rao", "Gender": "Female", "Class": "12-B", "Math": 90, "Physics": 88, "Chemistry": 92, "English": 85 },
        { "StudentID": "STU-016", "Name": "Nikhil Bose", "Gender": "Male", "Class": "10-B", "Math": 74, "Physics": 70, "Chemistry": 68, "English": 72 },
        { "StudentID": "STU-017", "Name": "Priya Sen", "Gender": "Female", "Class": "12-A", "Math": 48, "Physics": 50, "Chemistry": 52, "English": 60 },
        { "StudentID": "STU-018", "Name": "Siddharth Das", "Gender": "Male", "Class": "11-A", "Math": 82, "Physics": 80, "Chemistry": 85, "English": 79 }
      ],
      'Employees': [
        { "EmpID": "EMP-101", "FullName": "Rajesh Kumar", "Department": "Engineering", "Location": "Delhi", "MonthlySalary": 85000, "ExperienceYears": 6, "Status": "Active" },
        { "EmpID": "EMP-102", "FullName": "Sunita Verma", "Department": "Finance", "Location": "Mumbai", "MonthlySalary": 72000, "ExperienceYears": 4, "Status": "Active" },
        { "EmpID": "EMP-103", "FullName": "Amit Singh", "Department": "Human Resources", "Location": "Delhi ", "MonthlySalary": 62000, "ExperienceYears": 3, "Status": "Active" },
        { "EmpID": "EMP-104", "FullName": "Priya Nair", "Department": "Engineering", "Location": "Bangalore", "MonthlySalary": 96000, "ExperienceYears": 8, "Status": "Active" },
        { "EmpID": "EMP-105", "FullName": "Vikram Patel", "Department": "Marketing", "Location": "Delhi", "MonthlySalary": 58000, "ExperienceYears": 2, "Status": "On Leave" },
        { "EmpID": "EMP-106", "FullName": "Ananya Roy", "Department": "Finance", "Location": "Bangalore", "MonthlySalary": 79000, "ExperienceYears": 5, "Status": "Active" }
      ]
    }
  },
  employees: {
    fileName: 'HR_Payroll_Directory.xlsx',
    sheets: {
      'Staff': [
        { "EmpID": "EMP-101", "FullName": "Rajesh Kumar", "Department": "Engineering", "Location": "Delhi", "MonthlySalary": 85000, "ExperienceYears": 6, "Status": "Active" },
        { "EmpID": "EMP-102", "FullName": "Sunita Verma", "Department": "Finance", "Location": "Mumbai", "MonthlySalary": 72000, "ExperienceYears": 4, "Status": "Active" },
        { "EmpID": "EMP-103", "FullName": "Amit Singh", "Department": "Human Resources", "Location": "Delhi ", "MonthlySalary": 62000, "ExperienceYears": 3, "Status": "Active" },
        { "EmpID": "EMP-104", "FullName": "Priya Nair", "Department": "Engineering", "Location": "Bangalore", "MonthlySalary": 96000, "ExperienceYears": 8, "Status": "Active" },
        { "EmpID": "EMP-105", "FullName": "Vikram Patel", "Department": "Marketing", "Location": "Delhi", "MonthlySalary": 58000, "ExperienceYears": 2, "Status": "On Leave" },
        { "EmpID": "EMP-106", "FullName": "Ananya Roy", "Department": "Finance", "Location": "Bangalore", "MonthlySalary": 79000, "ExperienceYears": 5, "Status": "Active" }
      ]
    }
  }
};
window.addEventListener('DOMContentLoaded', () => {
  lucide.createIcons();
  loadSampleDataset('sales'); // Matches screenshot default (OrderID, Region, Category, Sales, Profit)

  // Dismiss dropdowns when clicking outside
  document.addEventListener('click', (e) => {
    const sBtn = document.getElementById('sampleDropdownBtn');
    const sMenu = document.getElementById('sampleDropdownMenu');
    if (sBtn && sMenu && !sBtn.contains(e.target) && !sMenu.contains(e.target)) {
      sMenu.classList.add('hidden');
    }

    const eBtn = document.getElementById('exportDropdownBtn');
    const eMenu = document.getElementById('exportDropdownMenu');
    if (eBtn && eMenu && !eBtn.contains(e.target) && !eMenu.contains(e.target)) {
      eMenu.classList.add('hidden');
    }
  });
});

function toggleSampleMenu() {
  document.getElementById('sampleDropdownMenu').classList.toggle('hidden');
}
function toggleExportMenu() {
  document.getElementById('exportDropdownMenu').classList.toggle('hidden');
}

function loadSampleDataset(key) {
  const sample = SAMPLE_DATASETS[key];
  if (!sample) return;
  AppState.fileName = sample.fileName;
  AppState.sheets = JSON.parse(JSON.stringify(sample.sheets));
  AppState.activeSheet = Object.keys(sample.sheets)[0];
  loadActiveSheet();
  showToast(`Loaded ${sample.fileName}`);
  document.getElementById('sampleDropdownMenu').classList.add('hidden');
}

function switchSheet(sheetName) {
  if (!AppState.sheets[sheetName]) return;
  AppState.activeSheet = sheetName;
  loadActiveSheet();
}

function loadActiveSheet() {
  const rows = AppState.sheets[AppState.activeSheet] || [];
  AppState.originalRows = JSON.parse(JSON.stringify(rows));
  AppState.currentRows = JSON.parse(JSON.stringify(rows));
  AppState.filteredRows = [...AppState.currentRows];

  analyzeColumns();

  // Update banner indicators
  document.getElementById('bannerFileName').textContent = AppState.fileName;
  document.getElementById('bannerRowCount').textContent = `${AppState.currentRows.length} Rows`;
  document.getElementById('bannerColCount').textContent = `${AppState.columns.length} Columns`;

  renderSheetPills();
  populateDropdowns();

  // Reset grid pagination and filters
  AppState.grid.page = 1;
  AppState.grid.searchQuery = '';
  AppState.grid.columnFilter = 'all';
  AppState.grid.sortCol = null;
  document.getElementById('gridSearchInput').value = '';

  // Render tab panels
  renderGridData();
  renderColumnProfiles();
  runCategoryAggregation();
  renderStudioChart();

  lucide.createIcons();
}

function renderSheetPills() {
  const container = document.getElementById('sheetPillContainer');
  container.innerHTML = '';
  const sheetNames = Object.keys(AppState.sheets);

  sheetNames.forEach(name => {
    const btn = document.createElement('button');
    const isActive = name === AppState.activeSheet;
    btn.onclick = () => switchSheet(name);
    btn.className = `px-3.5 py-1 rounded-full text-xs font-semibold transition shrink-0 ${
      isActive
        ? 'bg-indigo-600 text-white shadow-xs'
        : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
    }`;
    btn.textContent = name;
    container.appendChild(btn);
  });
}

function analyzeColumns() {
  if (AppState.currentRows.length === 0) {
    AppState.columns = [];
    AppState.colMeta = {};
    return;
  }

  AppState.columns = Object.keys(AppState.currentRows[0]);
  AppState.colMeta = {};

  AppState.columns.forEach(col => {
    let nonNull = 0;
    let numCount = 0;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;
    const freqMap = {};

    AppState.currentRows.forEach(row => {
      const val = row[col];
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        nonNull++;
        const strVal = String(val).trim();
        freqMap[strVal] = (freqMap[strVal] || 0) + 1;

        const n = Number(val);
        if (!isNaN(n) && typeof val !== 'boolean') {
          numCount++;
          sum += n;
          if (n < min) min = n;
          if (n > max) max = n;
        }
      }
    });

    // Column typing
    const isNumeric = nonNull > 0 && (numCount / nonNull) > 0.6;
    const distinctKeys = Object.keys(freqMap);
    let modeKey = '-';
    let modeCount = 0;

    distinctKeys.forEach(k => {
      if (freqMap[k] > modeCount) {
        modeCount = freqMap[k];
        modeKey = k;
      }
    });

    AppState.colMeta[col] = {
      type: isNumeric ? 'numeric' : 'categorical',
      nonNull,
      nullCount: AppState.currentRows.length - nonNull,
      percentFilled: Math.round((nonNull / (AppState.currentRows.length || 1)) * 100),
      distinctCount: distinctKeys.length,
      sum: isNumeric ? sum : null,
      avg: isNumeric && numCount > 0 ? (sum / numCount) : null,
      min: isNumeric && numCount > 0 ? min : null,
      max: isNumeric && numCount > 0 ? max : null,
      mode: modeKey,
      modeCount
    };
  });
}

function populateDropdowns() {
  // 1. Grid Column Filter
  const gridColSelect = document.getElementById('gridColumnFilterSelect');
  gridColSelect.innerHTML = '<option value="all">All Columns</option>';
  AppState.columns.forEach(col => {
    gridColSelect.innerHTML += `<option value="${col}">${col}</option>`;
  });

  // 2. Category Aggregation Selectors
  const catGroup = document.getElementById('categoryGroupSelect');
  const catMetric = document.getElementById('categoryMetricSelect');
  catGroup.innerHTML = '';
  catMetric.innerHTML = '';

  AppState.columns.forEach(col => {
    catGroup.innerHTML += `<option value="${col}">${col}</option>`;
  });

  const numCols = AppState.columns.filter(c => AppState.colMeta[c]?.type === 'numeric');
  (numCols.length > 0 ? numCols : AppState.columns).forEach(col => {
    catMetric.innerHTML += `<option value="${col}">${col}</option>`;
  });

  // Defaults
  const catCol = AppState.columns.find(c => AppState.colMeta[c]?.type === 'categorical');
  if (catCol) catGroup.value = catCol;
  if (numCols.length > 0) catMetric.value = numCols[0];

  // 3. Studio Chart Selectors
  const studioX = document.getElementById('studioXCol');
  const studioY = document.getElementById('studioYCol');
  studioX.innerHTML = '';
  studioY.innerHTML = '';

  AppState.columns.forEach(c => studioX.innerHTML += `<option value="${c}">${c}</option>`);
  (numCols.length > 0 ? numCols : AppState.columns).forEach(c => studioY.innerHTML += `<option value="${c}">${c}</option>`);

  if (catCol) studioX.value = catCol;
  if (numCols.length > 0) studioY.value = numCols[0];
}

// Exact recreation of the Column Profiles & Data Health cards shown in screenshot
function renderColumnProfiles() {
  const container = document.getElementById('columnProfilesContainer');
  container.innerHTML = '';

  AppState.columns.forEach(col => {
    const meta = AppState.colMeta[col];
    const isNumeric = meta.type === 'numeric';

    // Format numbers cleanly (e.g. 3100 -> 3.1k)
    const formatVal = (val) => {
      if (val === null || val === undefined || isNaN(val)) return '-';
      if (Math.abs(val) >= 1000000) return (val / 1000000).toFixed(1) + 'M';
      if (Math.abs(val) >= 1000) return (val / 1000).toFixed(1) + 'k';
      return Number.isInteger(val) ? val : val.toFixed(1);
    };

    const card = document.createElement('div');
    card.className = 'bg-white rounded-2xl border border-slate-200/80 p-4 sm:p-5 shadow-[0_1px_3px_rgba(0,0,0,0.02)] space-y-3';

    // Pill badge colors matching screenshot
    const badgeHtml = isNumeric
      ? `<span class="text-[11px] font-semibold px-2.5 py-0.5 rounded-full bg-amber-50 text-amber-600 border border-amber-200/70">numeric</span>`
      : `<span class="text-[11px] font-semibold px-2.5 py-0.5 rounded-full bg-indigo-50 text-indigo-600 border border-indigo-200/70">categorical</span>`;

    let metricsHtml = '';
    if (isNumeric) {
      metricsHtml = `
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Filled:</span>
          <span class="font-medium text-slate-800">${meta.nonNull} (${meta.percentFilled}%)</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Unique:</span>
          <span class="font-medium text-slate-800">${meta.distinctCount}</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Min:</span>
          <span class="font-medium text-slate-800">${formatVal(meta.min)}</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Max:</span>
          <span class="font-medium text-slate-800">${formatVal(meta.max)}</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Mean:</span>
          <span class="font-medium text-indigo-600">${formatVal(meta.avg)}</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Sum:</span>
          <span class="font-medium text-slate-800">${formatVal(meta.sum)}</span>
        </div>
      `;
    } else {
      metricsHtml = `
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Filled:</span>
          <span class="font-medium text-slate-800">${meta.nonNull} (${meta.percentFilled}%)</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Unique:</span>
          <span class="font-medium text-slate-800">${meta.distinctCount}</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Top Value:</span>
          <span class="font-medium text-slate-800 truncate max-w-[150px]">${meta.mode}</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Frequency:</span>
          <span class="font-medium text-slate-800">${meta.modeCount}x</span>
        </div>
      `;
    }

    card.innerHTML = `
      <div class="flex items-center justify-between pb-1 border-b border-slate-100">
        <h4 class="font-bold text-sm text-slate-900 truncate" title="${col}">${col}</h4>
        ${badgeHtml}
      </div>

      <div class="space-y-1.5 pt-1">
        ${metricsHtml}
      </div>

      <!-- Completeness Bar -->
      <div class="w-full bg-slate-100 h-1.5 rounded-full overflow-hidden mt-2">
        <div class="h-full ${meta.percentFilled === 100 ? 'bg-emerald-500' : 'bg-amber-500'}" style="width: ${meta.percentFilled}%"></div>
      </div>
    `;

    container.appendChild(card);
  });
}

// Quick Data Transformations Handlers
function hygieneRemoveDuplicates() {
  const initial = AppState.currentRows.length;
  const seen = new Set();
  AppState.currentRows = AppState.currentRows.filter(row => {
    const s = JSON.stringify(row);
    if (seen.has(s)) return false;
    seen.add(s);
    return true;
  });

  const removed = initial - AppState.currentRows.length;
  refreshPostTransformation(`Removed ${removed} duplicate row(s)`);
}

function hygieneStripNullRows() {
  const initial = AppState.currentRows.length;
  AppState.currentRows = AppState.currentRows.filter(row => {
    return Object.values(row).some(v => v !== undefined && v !== null && String(v).trim() !== '');
  });

  const removed = initial - AppState.currentRows.length;
  refreshPostTransformation(`Stripped ${removed} empty/blank row(s)`);
}

function hygieneTrimWhitespace() {
  let count = 0;
  AppState.currentRows.forEach(row => {
    Object.keys(row).forEach(k => {
      if (typeof row[k] === 'string') {
        const trimmed = row[k].trim().replace(/\s+/g, ' ');
        if (trimmed !== row[k]) {
          row[k] = trimmed;
          count++;
        }
      }
    });
  });
  refreshPostTransformation(`Trimmed and standardized ${count} cell(s) with trailing spaces`);
}

function hygieneRestoreOriginal() {
  AppState.currentRows = JSON.parse(JSON.stringify(AppState.originalRows));
  refreshPostTransformation('Restored original dataset successfully');
}

function refreshPostTransformation(message) {
  analyzeColumns();
  document.getElementById('bannerRowCount').textContent = `${AppState.currentRows.length} Rows`;
  renderGridData();
  renderColumnProfiles();
  runCategoryAggregation();
  renderStudioChart();
  showToast(message);
}

// Grid Rendering
function renderGridData() {
  const colFilter = document.getElementById('gridColumnFilterSelect').value;
  const search = (document.getElementById('gridSearchInput').value || '').toLowerCase().trim();

  let list = [...AppState.currentRows];
  if (search) {
    list = list.filter(row => {
      if (colFilter === 'all') {
        return AppState.columns.some(col => String(row[col] ?? '').toLowerCase().includes(search));
      } else {
        return String(row[colFilter] ?? '').toLowerCase().includes(search);
      }
    });
  }

  if (AppState.grid.sortCol) {
    const col = AppState.grid.sortCol;
    const dir = AppState.grid.sortDir === 'asc' ? 1 : -1;
    const isNum = AppState.colMeta[col]?.type === 'numeric';

    list.sort((a, b) => {
      const valA = a[col];
      const valB = b[col];
      if (isNum) return (Number(valA || 0) - Number(valB || 0)) * dir;
      return String(valA || '').localeCompare(String(valB || '')) * dir;
    });
  }

  AppState.filteredRows = list;

  document.getElementById('gridShowingCountLabel').textContent =
    `Showing ${list.length} of ${AppState.currentRows.length}`;

  renderGridHeaders();
  renderGridRows();
  updateGridPagination();
}

function renderGridHeaders() {
  const tr = document.getElementById('gridHeaderRow');
  tr.innerHTML = '';

  const thIdx = document.createElement('th');
  thIdx.className = 'py-3 px-3.5 w-12 text-slate-400 font-semibold border-r border-slate-100/80';
  thIdx.textContent = '#';
  tr.appendChild(thIdx);

  AppState.columns.forEach(col => {
    const meta = AppState.colMeta[col];
    const isNumeric = meta?.type === 'numeric';
    const isSorted = AppState.grid.sortCol === col;
    const th = document.createElement('th');
    th.className = 'py-3 px-3.5 font-semibold text-slate-700 whitespace-nowrap cursor-pointer hover:bg-slate-100 transition select-none';
    th.onclick = () => sortGridBy(col);

    let sortIcon = '<i data-lucide="chevrons-up-down" class="w-3 h-3 text-slate-300 inline ml-1"></i>';
    if (isSorted) {
      sortIcon = AppState.grid.sortDir === 'asc'
        ? '<i data-lucide="chevron-up" class="w-3 h-3 text-indigo-600 inline ml-1"></i>'
        : '<i data-lucide="chevron-down" class="w-3 h-3 text-indigo-600 inline ml-1"></i>';
    }

    const typeBadge = isNumeric
      ? '<span class="text-[9.5px] font-mono font-bold uppercase tracking-wider px-1.5 py-0.5 rounded bg-amber-50 text-amber-700 border border-amber-200/60">123</span>'
      : '<span class="text-[9.5px] font-mono font-bold uppercase tracking-wider px-1.5 py-0.5 rounded bg-blue-50 text-blue-700 border border-blue-200/60">ABC</span>';

    th.innerHTML = `
      <div class="flex items-center justify-between gap-3">
        <span>${col}</span>
        <div class="flex items-center gap-1.5">
          ${typeBadge}
          ${sortIcon}
        </div>
      </div>
    `;
    tr.appendChild(th);
  });

  lucide.createIcons();
}

function renderGridRows() {
  const tbody = document.getElementById('gridTableBody');
  tbody.innerHTML = '';

  const { page, pageSize } = AppState.grid;
  const start = (page - 1) * pageSize;
  const end = start + pageSize;
  const pageRows = AppState.filteredRows.slice(start, end);

  if (pageRows.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="${AppState.columns.length + 1}" class="py-12 text-center text-slate-400">
          No matching records found.
        </td>
      </tr>
    `;
    return;
  }

  pageRows.forEach((row, i) => {
    const rowNumber = start + i + 1;
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-indigo-50/30 transition-colors border-b border-slate-100';

    let html = `<td class="py-2.5 px-3.5 text-slate-400 font-mono text-[11px] border-r border-slate-100/80">${rowNumber}</td>`;

    AppState.columns.forEach(col => {
      const val = row[col];
      const isNum = AppState.colMeta[col]?.type === 'numeric';
      const displayVal = (val === null || val === undefined || String(val).trim() === '')
        ? '<span class="text-slate-300 italic">null</span>'
        : String(val);

      html += `<td class="py-2.5 px-3.5 whitespace-nowrap text-slate-700 ${isNum ? 'font-mono text-right' : ''}">${displayVal}</td>`;
    });

    tr.innerHTML = html;
    tbody.appendChild(tr);
  });
}

function sortGridBy(col) {
  if (AppState.grid.sortCol === col) {
    AppState.grid.sortDir = AppState.grid.sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    AppState.grid.sortCol = col;
    AppState.grid.sortDir = 'asc';
  }
  renderGridData();
}

function handleGridSearch(val) {
  AppState.grid.searchQuery = val;
  AppState.grid.page = 1;
  renderGridData();
}

function resetGridSearch() {
  document.getElementById('gridSearchInput').value = '';
  document.getElementById('gridColumnFilterSelect').value = 'all';
  AppState.grid.sortCol = null;
  AppState.grid.page = 1;
  renderGridData();
}

function updateGridPagination() {
  const total = AppState.filteredRows.length;
  const totalPages = Math.ceil(total / AppState.grid.pageSize) || 1;
  document.getElementById('gridPageIndicator').textContent = `Page ${AppState.grid.page} of ${totalPages}`;
  document.getElementById('btnGridPrev').disabled = AppState.grid.page <= 1;
  document.getElementById('btnGridNext').disabled = AppState.grid.page >= totalPages;
}

function navigateGridPage(delta) {
  const totalPages = Math.ceil(AppState.filteredRows.length / AppState.grid.pageSize) || 1;
  const newPage = AppState.grid.page + delta;
  if (newPage >= 1 && newPage <= totalPages) {
    AppState.grid.page = newPage;
    renderGridRows();
    updateGridPagination();
  }
}

function changeGridPageSize(size) {
  AppState.grid.pageSize = Number(size);
  AppState.grid.page = 1;
  renderGridData();
}

// Category Summary & Averages Engine
function runCategoryAggregation() {
  const groupCol = document.getElementById('categoryGroupSelect').value;
  const metricCol = document.getElementById('categoryMetricSelect').value;

  if (!groupCol || !metricCol) return;

  const groupMap = {};
  let grandTotalSum = 0;
  let grandTotalCount = 0;

  AppState.currentRows.forEach(row => {
    const key = String(row[groupCol] || '(Blank)').trim();
    const val = Number(row[metricCol]) || 0;

    if (!groupMap[key]) {
      groupMap[key] = { count: 0, sum: 0, min: Infinity, max: -Infinity };
    }

    groupMap[key].count++;
    groupMap[key].sum += val;
    grandTotalSum += val;
    grandTotalCount++;

    if (val < groupMap[key].min) groupMap[key].min = val;
    if (val > groupMap[key].max) groupMap[key].max = val;
  });

  const categories = Object.keys(groupMap).sort();
  const tbody = document.getElementById('categoryTableBody');
  const tfoot = document.getElementById('categoryTableFoot');
  tbody.innerHTML = '';
  tfoot.innerHTML = '';

  const chartLabels = [];
  const chartAverages = [];

  categories.forEach(cat => {
    const item = groupMap[cat];
    const avg = item.sum / (item.count || 1);
    const pctOfTotal = grandTotalSum > 0 ? ((item.sum / grandTotalSum) * 100) : 0;

    chartLabels.push(cat);
    chartAverages.push(parseFloat(avg.toFixed(2)));

    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';
    tr.innerHTML = `
      <td class="py-2.5 px-3 font-sans font-semibold text-slate-900">${cat}</td>
      <td class="py-2.5 px-3">${item.count.toLocaleString()}</td>
      <td class="py-2.5 px-3 font-semibold">${item.sum.toLocaleString(undefined, { maximumFractionDigits: 1 })}</td>
      <td class="py-2.5 px-3 font-bold text-indigo-600">${avg.toLocaleString(undefined, { maximumFractionDigits: 2 })}</td>
      <td class="py-2.5 px-3 text-slate-500">${item.min === Infinity ? '-' : item.min}</td>
      <td class="py-2.5 px-3 text-slate-500">${item.max === -Infinity ? '-' : item.max}</td>
      <td class="py-2.5 px-3 text-right font-medium">${pctOfTotal.toFixed(1)}%</td>
    `;
    tbody.appendChild(tr);
  });

  const overallAvg = grandTotalSum / (grandTotalCount || 1);
  tfoot.innerHTML = `
    <tr>
      <td class="py-2.5 px-3">Grand Total (${categories.length} Categories)</td>
      <td class="py-2.5 px-3">${grandTotalCount.toLocaleString()}</td>
      <td class="py-2.5 px-3">${grandTotalSum.toLocaleString(undefined, { maximumFractionDigits: 1 })}</td>
      <td class="py-2.5 px-3 text-indigo-700">${overallAvg.toLocaleString(undefined, { maximumFractionDigits: 2 })}</td>
      <td class="py-2.5 px-3">-</td>
      <td class="py-2.5 px-3">-</td>
      <td class="py-2.5 px-3 text-right">100.0%</td>
    </tr>
  `;

  document.getElementById('categoryChartTitle').textContent = `Average ${metricCol} by ${groupCol}`;

  const ctx = document.getElementById('categoryChartCanvas').getContext('2d');
  if (AppState.categoryChartInstance) AppState.categoryChartInstance.destroy();

  AppState.categoryChartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: chartLabels,
      datasets: [{
        label: `Average ${metricCol}`,
        data: chartAverages,
        backgroundColor: '#4f46e5',
        borderRadius: 6,
        barThickness: categories.length > 8 ? undefined : 28
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => ` Average ${metricCol}: ${ctx.raw}`
          }
        }
      },
      scales: {
        y: { ticks: { font: { family: 'Inter', size: 10 } }, grid: { color: '#f1f5f9' } },
        x: { grid: { display: false }, ticks: { font: { family: 'Inter', size: 10 } } }
      }
    }
  });
}

// Chart Studio
function renderStudioChart() {
  const xCol = document.getElementById('studioXCol').value;
  const yCol = document.getElementById('studioYCol').value;
  const type = document.getElementById('studioChartType').value;
  const agg = document.getElementById('studioAgg').value;

  if (!xCol || !yCol) return;

  const grouped = {};
  AppState.currentRows.forEach(row => {
    const k = String(row[xCol] || '(Blank)').trim();
    const v = Number(row[yCol]) || 0;

    if (!grouped[k]) {
      grouped[k] = { sum: 0, count: 0, min: Infinity, max: -Infinity };
    }
    grouped[k].sum += v;
    grouped[k].count++;
    if (v < grouped[k].min) grouped[k].min = v;
    if (v > grouped[k].max) grouped[k].max = v;
  });

  const labels = Object.keys(grouped);
  const data = labels.map(k => {
    const item = grouped[k];
    if (agg === 'avg') return parseFloat((item.sum / (item.count || 1)).toFixed(2));
    if (agg === 'count') return item.count;
    if (agg === 'max') return item.max;
    if (agg === 'min') return item.min;
    return item.sum;
  });

  const ctx = document.getElementById('studioChartCanvas').getContext('2d');
  if (AppState.studioChartInstance) AppState.studioChartInstance.destroy();

  AppState.studioChartInstance = new Chart(ctx, {
    type: type,
    data: {
      labels: labels,
      datasets: [{
        label: `${agg.toUpperCase()}(${yCol})`,
        data: data,
        backgroundColor: ['#4f46e5', '#06b6d4', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6', '#64748b', '#3b82f6', '#14b8a6', '#f97316'],
        borderWidth: 1.5,
        borderColor: '#ffffff',
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: ['doughnut', 'pie'].includes(type),
          position: 'bottom',
          labels: { boxWidth: 10, font: { family: 'Inter', size: 10 } }
        }
      },
      scales: ['doughnut', 'pie'].includes(type) ? {} : {
        y: { grid: { color: '#f1f5f9' }, ticks: { font: { family: 'Inter', size: 10 } } },
        x: { grid: { display: false }, ticks: { font: { family: 'Inter', size: 10 } } }
      }
    }
  });

  document.getElementById('studioChartTitle').textContent = `${agg.toUpperCase()} of ${yCol} by ${xCol}`;
  document.getElementById('studioChartDesc').textContent = `Plotted across all ${labels.length} distinct items`;
}

function downloadStudioChart() {
  if (!AppState.studioChartInstance) return;
  const url = AppState.studioChartInstance.toBase64Image();
  const a = document.createElement('a');
  a.href = url;
  a.download = `SheetPulse_Chart_${Date.now()}.png`;
  a.click();
  showToast('Chart snapshot downloaded as PNG');
}

// File Upload Handler (SheetJS)
function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  AppState.fileName = file.name;
  const reader = new FileReader();

  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      AppState.sheets = {};
      workbook.SheetNames.forEach(sheetName => {
        AppState.sheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
      });

      AppState.activeSheet = workbook.SheetNames[0];
      loadActiveSheet();
      showToast(`Uploaded ${file.name}`);
    } catch (err) {
      console.error(err);
      showToast('Failed to parse uploaded Excel file', true);
    }
  };

  reader.readAsArrayBuffer(file);
  event.target.value = '';
}

// Export Options
function exportCurrentData(format) {
  const ext = format === 'csv' ? 'csv' : format === 'xlsx' ? 'xlsx' : 'json';
  if (ext === 'json') {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(AppState.filteredRows, null, 2));
    const a = document.createElement('a');
    a.href = dataStr;
    a.download = `SheetPulse_${AppState.activeSheet}_${Date.now()}.json`;
    a.click();
  } else {
    const ws = XLSX.utils.json_to_sheet(AppState.filteredRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, AppState.activeSheet);
    XLSX.writeFile(wb, `SheetPulse_${AppState.activeSheet}_${Date.now()}.${ext}`);
  }
  showToast(`Exported ${AppState.filteredRows.length} records to .${ext}`);
  document.getElementById('exportDropdownMenu').classList.add('hidden');
}

// Tab Navigation Switcher
function switchTab(tabId) {
  document.querySelectorAll('.tab-panel').forEach(el => el.classList.add('hidden'));
  document.querySelectorAll('.tab-nav-btn').forEach(btn => {
    btn.classList.remove('text-indigo-600', 'border-indigo-600');
    btn.classList.add('text-slate-500', 'border-transparent');
  });

  const targetPane = document.getElementById(`tabContent-${tabId}`);
  const targetBtn = document.getElementById(`tabBtn-${tabId}`);

  if (targetPane) targetPane.classList.remove('hidden');
  if (targetBtn) {
    targetBtn.classList.add('text-indigo-600', 'border-indigo-600');
    targetBtn.classList.remove('text-slate-500', 'border-transparent');
  }

  if (tabId === 'category' && AppState.categoryChartInstance) {
    AppState.categoryChartInstance.resize();
  }
  if (tabId === 'charts' && AppState.studioChartInstance) {
    AppState.studioChartInstance.resize();
  }

  lucide.createIcons();
}

function showToast(msg, isError = false) {
  const toast = document.getElementById('toastNotification');
  const msgSpan = document.getElementById('toastMessage');
  const icon = document.getElementById('toastIcon');

  msgSpan.textContent = msg;
  if (isError) {
    icon.setAttribute('data-lucide', 'alert-circle');
    icon.className = 'w-4 h-4 text-rose-400';
  } else {
    icon.setAttribute('data-lucide', 'check-circle');
    icon.className = 'w-4 h-4 text-emerald-400';
  }

  toast.classList.remove('opacity-0', 'translate-y-20', 'pointer-events-none');
  lucide.createIcons();

  setTimeout(() => {
    toast.classList.add('opacity-0', 'translate-y-20', 'pointer-events-none');
  }, 3000);
}