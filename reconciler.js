/**
 * Platform Payout Reconciler - Core Logic
 */

// Global State
let pendingWorkbook = null;
let incomeWorkbook = null;
let pendingData = []; // Array of parsed pending orders
let incomeData = [];  // Array of parsed income payouts
let reconResults = []; // Array of reconciled records
let activeFilter = 'all';
let searchQuery = '';

// Helper formatters
const formatCurrency = (val) => {
  const num = parseFloat(val) || 0;
  return '฿' + num.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};

const parseNum = (val) => {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  const cleaned = String(val).replace(/,/g, '').trim();
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
};

// UI Elements
const dropzonePending = document.getElementById('dropzone-pending');
const dropzoneIncome = document.getElementById('dropzone-income');
const filePendingInput = document.getElementById('file-pending');
const fileIncomeInput = document.getElementById('file-income');
const pendingFileInfo = document.getElementById('pending-file-info');
const incomeFileInfo = document.getElementById('income-file-info');
const pendingSheetSelect = document.getElementById('pending-sheet-select');
const incomeSheetSelect = document.getElementById('income-sheet-select');
const metricsPanel = document.getElementById('metrics-panel');
const resultsPanel = document.getElementById('results-panel');
const tableBody = document.getElementById('table-body');
const searchInput = document.getElementById('search-input');
const btnLoadDemo = document.getElementById('btn-load-demo');
const btnReset = document.getElementById('btn-reset');
const btnCopyAnomalies = document.getElementById('btn-copy-anomalies');
const btnExportExcel = document.getElementById('btn-export-excel');
const filterTabs = document.querySelectorAll('.tab-btn');

// Setup Drag & Drop Handlers
function setupDropzone(dropzone, fileInput, onFileSelected) {
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
      onFileSelected(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      onFileSelected(e.target.files[0]);
    }
  });
}

setupDropzone(dropzonePending, filePendingInput, handlePendingFile);
setupDropzone(dropzoneIncome, fileIncomeInput, handleIncomeFile);

// Read Excel / CSV File
function readFileAsync(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        resolve(workbook);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// Find header row in sheet
function parseSheetWithSmartHeaders(worksheet) {
  const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  if (!json || json.length === 0) return { headers: [], rows: [] };

  // Scan first 15 rows for header keywords
  let headerRowIndex = 0;
  for (let r = 0; r < Math.min(15, json.length); r++) {
    const rowStr = json[r].map(c => String(c).toLowerCase()).join(' ');
    if (
      rowStr.includes('หมายเลขคำสั่งซื้อ') ||
      rowStr.includes('order id') ||
      rowStr.includes('คำสั่งซื้อ') ||
      rowStr.includes('เวลาการชำระเงิน') ||
      rowStr.includes('จำนวนเงินที่ชำระ') ||
      rowStr.includes('settlement')
    ) {
      headerRowIndex = r;
      break;
    }
  }

  const rawHeaders = json[headerRowIndex] || [];
  const headers = rawHeaders.map((h, i) => String(h || '').trim() || `Column_${i + 1}`);

  const rows = [];
  for (let r = headerRowIndex + 1; r < json.length; r++) {
    const row = json[r];
    if (!row || row.length === 0 || row.every(c => c === '' || c === null || c === undefined)) continue;
    const obj = {};
    headers.forEach((h, colIdx) => {
      obj[h] = row[colIdx];
    });
    rows.push(obj);
  }

  return { headers, rows };
}

// Parse Pending File
async function handlePendingFile(file) {
  try {
    document.getElementById('badge-pending').innerText = 'กำลังอ่านไฟล์...';
    pendingWorkbook = await readFileAsync(file);
    
    // Populate sheet select
    pendingSheetSelect.innerHTML = '';
    pendingWorkbook.SheetNames.forEach(name => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.innerText = name;
      pendingSheetSelect.appendChild(opt);
    });

    // Default to the largest sheet or September sheet
    const defaultSheet = pendingWorkbook.SheetNames.find(s => s.includes('9.2569') || s.includes('9') || s.toLowerCase().includes('pending')) || pendingWorkbook.SheetNames[0];
    pendingSheetSelect.value = defaultSheet;

    document.getElementById('pending-filename').innerText = file.name;
    pendingFileInfo.classList.remove('hidden');
    document.getElementById('badge-pending').innerText = 'โหลดสำเร็จ';
    document.getElementById('badge-pending').className = 'text-xs font-semibold text-emerald-600';

    loadPendingSheet(defaultSheet);
  } catch (err) {
    alert('เกิดข้อผิดพลาดในการอ่านไฟล์รอเงินเข้า: ' + err.message);
    document.getElementById('badge-pending').innerText = 'ผิดพลาด';
  }
}

pendingSheetSelect.addEventListener('change', () => {
  loadPendingSheet(pendingSheetSelect.value);
});

function loadPendingSheet(sheetName) {
  const ws = pendingWorkbook.Sheets[sheetName];
  const { headers, rows } = parseSheetWithSmartHeaders(ws);

  // Normalize column mapping for pending
  pendingData = rows.map(r => {
    // Find matching keys
    const orderKey = Object.keys(r).find(k => k.includes('หมายเลขคำสั่งซื้อ/การปรับ') || k.includes('หมายเลขคำสั่งซื้อ') || k.toLowerCase().includes('order id') || k.toLowerCase().includes('order_id')) || '';
    const amountKey = Object.keys(r).find(k => k.includes('จำนวนเงินที่ชำระโดยประมาณ') || k.includes('ยอดรวมค่าสินค้า') || k.toLowerCase().includes('estimated') || k.includes('จำนวนเงิน')) || '';
    const estTimeKey = Object.keys(r).find(k => k.includes('เวลาการชำระเงินโดยประมาณ') || k.toLowerCase().includes('estimated payout') || k.includes('เวลาการชำระ')) || '';
    const reasonKey = Object.keys(r).find(k => k.includes('เหตุผลของการไม่ชำระเงิน') || k.toLowerCase().includes('reason')) || '';
    const typeKey = Object.keys(r).find(k => k.includes('ประเภทธุรกรรม') || k.toLowerCase().includes('type')) || '';
    const createdKey = Object.keys(r).find(k => k.includes('วันที่สร้างคำสั่งซื้อ') || k.includes('วันที่สร้างธุรกรรม')) || '';
    const deliveredKey = Object.keys(r).find(k => k.includes('วันที่ส่งมอบคำสั่งซื้อ') || k.includes('จัดส่ง')) || '';

    const orderId = String(r[orderKey] || '').trim();
    return {
      orderId,
      estimatedAmount: parseNum(r[amountKey]),
      estimatedPayoutTime: String(r[estTimeKey] || '').trim(),
      reason: String(r[reasonKey] || '').trim(),
      type: String(r[typeKey] || 'คำสั่งซื้อ').trim(),
      orderCreatedDate: String(r[createdKey] || '').trim(),
      deliveredDate: String(r[deliveredKey] || '').trim(),
      raw: r
    };
  }).filter(item => item.orderId !== '');

  document.getElementById('pending-stats').innerText = `${pendingData.length.toLocaleString()} รายการ`;
  runReconciliation();
}

// Parse Income File
async function handleIncomeFile(file) {
  try {
    document.getElementById('badge-income').innerText = 'กำลังอ่านไฟล์...';
    incomeWorkbook = await readFileAsync(file);

    incomeSheetSelect.innerHTML = '';
    incomeWorkbook.SheetNames.forEach(name => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.innerText = name;
      incomeSheetSelect.appendChild(opt);
    });

    const defaultSheet = incomeWorkbook.SheetNames[0];
    incomeSheetSelect.value = defaultSheet;

    document.getElementById('income-filename').innerText = file.name;
    incomeFileInfo.classList.remove('hidden');
    document.getElementById('badge-income').innerText = 'โหลดสำเร็จ';
    document.getElementById('badge-income').className = 'text-xs font-semibold text-emerald-600';

    loadIncomeSheet(defaultSheet);
  } catch (err) {
    alert('เกิดข้อผิดพลาดในการอ่านไฟล์ Income: ' + err.message);
    document.getElementById('badge-income').innerText = 'ผิดพลาด';
  }
}

incomeSheetSelect.addEventListener('change', () => {
  loadIncomeSheet(incomeSheetSelect.value);
});

function loadIncomeSheet(sheetName) {
  const ws = incomeWorkbook.Sheets[sheetName];
  const { headers, rows } = parseSheetWithSmartHeaders(ws);

  incomeData = rows.map(r => {
    const orderKey = Object.keys(r).find(k => k.includes('หมายเลขคำสั่งซื้อ/การปรับ') || k.includes('หมายเลขคำสั่งซื้อ') || k.toLowerCase().includes('order id') || k.toLowerCase().includes('order_id') || k.includes('รหัสคำสั่งซื้อ')) || '';
    const amountKey = Object.keys(r).find(k => k.includes('จำนวนเงินที่ชำระ') || k.includes('ยอดเงินสุทธิ') || k.includes('รายได้สุทธิ') || k.includes('จำนวนเงิน') || k.toLowerCase().includes('payout') || k.toLowerCase().includes('settled')) || '';
    const timeKey = Object.keys(r).find(k => k.includes('เวลาการชำระเงิน') || k.includes('วันที่ชำระเงิน') || k.includes('วันที่โอน') || k.toLowerCase().includes('settlement time')) || '';
    const typeKey = Object.keys(r).find(k => k.includes('ประเภทธุรกรรม') || k.toLowerCase().includes('type')) || '';

    const orderId = String(r[orderKey] || '').trim();
    return {
      orderId,
      actualAmount: parseNum(r[amountKey]),
      payoutTime: String(r[timeKey] || '').trim(),
      type: String(r[typeKey] || 'คำสั่งซื้อ').trim(),
      raw: r
    };
  }).filter(item => item.orderId !== '');

  document.getElementById('income-stats').innerText = `${incomeData.length.toLocaleString()} รายการ`;
  runReconciliation();
}

// Reconciliation Engine
// Helper to parse dates from string (supports YYYY/MM/DD, YYYY-MM-DD, DD/MM/YYYY)
function extractDateStr(str) {
  if (!str) return null;
  const match = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (match) {
    return `${match[1]}-${String(match[2]).padStart(2, '0')}-${String(match[3]).padStart(2, '0')}`;
  }
  const matchThai = str.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (matchThai) {
    return `${matchThai[3]}-${String(matchThai[2]).padStart(2, '0')}-${String(matchThai[1]).padStart(2, '0')}`;
  }
  return null;
}

// Reconciliation Engine
function runReconciliation() {
  if (pendingData.length === 0 && incomeData.length === 0) return;

  const pendingMap = new Map();
  pendingData.forEach(p => {
    pendingMap.set(p.orderId, p);
  });

  // Extract reference payout date from Income file
  let refIncomeDate = null;
  for (const inc of incomeData) {
    const d = extractDateStr(inc.payoutTime);
    if (d) {
      if (!refIncomeDate || d > refIncomeDate) refIncomeDate = d;
    }
  }

  const reconciled = [];
  const processedPendingIds = new Set();

  // 1. Process Income Items (ตรวจสอบทีละรายการที่เงินเข้าจริง เทียบกับไฟล์รอเงินเข้า)
  incomeData.forEach(inc => {
    const orderId = inc.orderId;
    const pendingItem = pendingMap.get(orderId);

    if (pendingItem) {
      processedPendingIds.add(orderId);
      const estAmt = pendingItem.estimatedAmount;
      const actAmt = inc.actualAmount;
      const diff = actAmt - estAmt;
      const hasMismatch = Math.abs(diff) >= 0.05;

      // ตรวจสอบวันที่นัดโอน
      let dateNote = '';
      const pendingDate = extractDateStr(pendingItem.estimatedPayoutTime);
      if (pendingDate && refIncomeDate && pendingDate > refIncomeDate) {
        dateNote = ` (⚡ โอนข้ามรอบ/เร่งโอน นัดเดิม: ${pendingItem.estimatedPayoutTime})`;
      }

      reconciled.push({
        status: hasMismatch ? 'MISMATCH' : 'MATCHED',
        orderId,
        type: inc.type || pendingItem.type,
        estimatedAmount: estAmt,
        actualAmount: actAmt,
        diff: diff,
        payoutTime: inc.payoutTime || pendingItem.estimatedPayoutTime,
        reason: pendingItem.reason || 'โอนสำเร็จ',
        orderCreatedDate: pendingItem.orderCreatedDate || '/',
        notes: hasMismatch 
          ? `ยอดเงินไม่ตรงกัน ส่วนต่าง ${formatCurrency(diff)}${dateNote}` 
          : `โอนสำเร็จ ยอดตรงกับประมาณการ${dateNote}`
      });
    } else {
      // 🚨 GHOST PAYOUT: เงินเข้าจริงแต่ไม่มีในไฟล์รอเงินเข้า (ตรวจจับการเอาออเดอร์อื่นมาจ่าย!)
      reconciled.push({
        status: 'GHOST',
        orderId,
        type: inc.type || 'คำสั่งซื้อ/ปรับปรุง',
        estimatedAmount: 0,
        actualAmount: inc.actualAmount,
        diff: inc.actualAmount,
        payoutTime: inc.payoutTime || 'วันที่โอน',
        reason: 'ไม่อยู่ในรายการรอเงินเข้า (Ghost Payout)',
        orderCreatedDate: '/',
        notes: '🚨 ผิดปกติ! เลขคำสั่งซื้อนี้ไม่มีอยู่ในไฟล์รอเงินเข้า (ระบบอาจนำออเดอร์อื่นมาจ่ายแทน หรือเป็นรายการแอบปล่อยยอดหลังยื่นตั๋ว)'
      });
    }
  });

  // 2. Process Remaining Pending Items (ตรวจเช็คออเดอร์ในไฟล์รอเงินเข้า ที่ถึงกำหนดแล้วแต่เงินไม่ยอมเข้า)
  pendingData.forEach(p => {
    if (!processedPendingIds.has(p.orderId)) {
      const pDate = extractDateStr(p.estimatedPayoutTime);
      const isOverdueDate = pDate && refIncomeDate && pDate <= refIncomeDate;
      const isDeliveredWaiting = p.reason.includes('จัดส่งสำเร็จแล้ว รอการชำระเงิน');

      // ถ้านัดโอนถึงกำหนดแล้ว หรือส่งสำเร็จแล้วรอเงิน แต่ไม่มีในไฟล์ Income ของวันนี้
      if (isOverdueDate || isDeliveredWaiting) {
        reconciled.push({
          status: 'OVERDUE',
          orderId: p.orderId,
          type: p.type,
          estimatedAmount: p.estimatedAmount,
          actualAmount: 0,
          diff: -p.estimatedAmount,
          payoutTime: p.estimatedPayoutTime,
          reason: p.reason,
          orderCreatedDate: p.orderCreatedDate,
          notes: `⏳ ถึงกำหนดโอน ${p.estimatedPayoutTime} แต่ไม่มีเงินเข้าในไฟล์ Income (ยอดถูกกัก/ตกหล่น ควรเปิด Ticket ถาม)`
        });
      }
    }
  });

  reconResults = reconciled;
  updateDashboardMetrics();
  renderTable();
}

// Update Dashboard Numbers
function updateDashboardMetrics() {
  metricsPanel.classList.remove('hidden');
  resultsPanel.classList.remove('hidden');

  const totalIncome = incomeData.reduce((acc, cur) => acc + cur.actualAmount, 0);
  const matched = reconResults.filter(r => r.status === 'MATCHED');
  const ghosts = reconResults.filter(r => r.status === 'GHOST');
  const overdues = reconResults.filter(r => r.status === 'OVERDUE');
  const mismatches = reconResults.filter(r => r.status === 'MISMATCH');

  document.getElementById('metric-total-income').innerText = formatCurrency(totalIncome);
  document.getElementById('metric-income-count').innerText = `${incomeData.length.toLocaleString()} รายการ`;

  document.getElementById('metric-matched-count').innerText = matched.length.toLocaleString();
  const matchedAmt = matched.reduce((acc, c) => acc + c.actualAmount, 0);
  document.getElementById('metric-matched-amount').innerText = formatCurrency(matchedAmt);

  document.getElementById('metric-ghost-count').innerText = ghosts.length.toLocaleString();
  const ghostAmt = ghosts.reduce((acc, c) => acc + c.actualAmount, 0);
  document.getElementById('metric-ghost-amount').innerText = `${formatCurrency(ghostAmt)} (ไม่อยู่ในคิว)`;

  document.getElementById('metric-overdue-count').innerText = overdues.length.toLocaleString();
  const overdueAmt = overdues.reduce((acc, c) => acc + c.estimatedAmount, 0);
  document.getElementById('metric-overdue-amount').innerText = `${formatCurrency(overdueAmt)} (ค้างโอน)`;

  document.getElementById('metric-mismatch-count').innerText = mismatches.length.toLocaleString();
  const mismatchDiff = mismatches.reduce((acc, c) => acc + Math.abs(c.diff), 0);
  document.getElementById('metric-mismatch-diff').innerText = `ส่วนต่าง ${formatCurrency(mismatchDiff)}`;

  // Tab counters
  document.getElementById('tab-count-all').innerText = reconResults.length.toLocaleString();
  document.getElementById('tab-count-ghost').innerText = ghosts.length.toLocaleString();
  document.getElementById('tab-count-overdue').innerText = overdues.length.toLocaleString();
  document.getElementById('tab-count-mismatch').innerText = mismatches.length.toLocaleString();
  document.getElementById('tab-count-matched').innerText = matched.length.toLocaleString();
}

// Render Table Rows
function renderTable() {
  tableBody.innerHTML = '';

  let filtered = reconResults;
  if (activeFilter === 'ghost') filtered = filtered.filter(r => r.status === 'GHOST');
  else if (activeFilter === 'overdue') filtered = filtered.filter(r => r.status === 'OVERDUE');
  else if (activeFilter === 'mismatch') filtered = filtered.filter(r => r.status === 'MISMATCH');
  else if (activeFilter === 'matched') filtered = filtered.filter(r => r.status === 'MATCHED');

  if (searchQuery) {
    const q = searchQuery.toLowerCase();
    filtered = filtered.filter(r => 
      r.orderId.toLowerCase().includes(q) || 
      r.reason.toLowerCase().includes(q) ||
      r.notes.toLowerCase().includes(q)
    );
  }

  document.getElementById('showing-text').innerText = `กำลังแสดง ${filtered.length.toLocaleString()} จากทั้งหมด ${reconResults.length.toLocaleString()} รายการ`;

  if (filtered.length === 0) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="9" class="py-8 text-center text-slate-400">ไม่พบรายการที่ตรงกับเงื่อนไขการค้นหา</td>`;
    tableBody.appendChild(tr);
    return;
  }

  filtered.forEach(r => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50/80 transition duration-150';

    let badgeHtml = '';
    if (r.status === 'GHOST') {
      badgeHtml = `<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[11px] font-semibold bg-rose-100 text-rose-800 border border-rose-200">🚨 เงินเข้าไม่ตรงนัด (Ghost)</span>`;
    } else if (r.status === 'OVERDUE') {
      badgeHtml = `<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[11px] font-semibold bg-amber-100 text-amber-800 border border-amber-200">⏳ ถึงกำหนดแต่ไม่โอน</span>`;
    } else if (r.status === 'MISMATCH') {
      badgeHtml = `<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[11px] font-semibold bg-purple-100 text-purple-800 border border-purple-200">⚠️ ยอดเงินไม่ตรง</span>`;
    } else {
      badgeHtml = `<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-100 text-emerald-800 border border-emerald-200">🟢 ตรงกันปกติ</span>`;
    }

    const diffClass = r.diff < 0 ? 'text-rose-600 font-semibold' : (r.diff > 0 ? 'text-emerald-600 font-semibold' : 'text-slate-400');
    const diffText = r.diff === 0 ? '-' : (r.diff > 0 ? `+${formatCurrency(r.diff)}` : formatCurrency(r.diff));

    tr.innerHTML = `
      <td class="py-2.5 px-4">${badgeHtml}</td>
      <td class="py-2.5 px-4 font-medium text-slate-800 mono select-all">${r.orderId}</td>
      <td class="py-2.5 px-4 text-slate-600">${r.type}</td>
      <td class="py-2.5 px-4 text-right mono text-slate-700">${formatCurrency(r.estimatedAmount)}</td>
      <td class="py-2.5 px-4 text-right mono font-medium text-slate-900">${formatCurrency(r.actualAmount)}</td>
      <td class="py-2.5 px-4 text-right mono ${diffClass}">${diffText}</td>
      <td class="py-2.5 px-4 text-slate-600">
        <div class="font-medium text-slate-800">${r.payoutTime || '-'}</div>
        <div class="text-[11px] text-slate-500">${r.reason || ''}</div>
      </td>
      <td class="py-2.5 px-4 text-slate-500 text-[11px]">${r.orderCreatedDate}</td>
      <td class="py-2.5 px-4">
        <button onclick="copyToClipboard('${r.orderId}')" class="px-2 py-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-600 text-[11px] font-medium transition" title="คัดลอก Order ID">
          คัดลอก
        </button>
      </td>
    `;
    tableBody.appendChild(tr);
  });
}

// Copy single order
window.copyToClipboard = (text) => {
  navigator.clipboard.writeText(text).then(() => {
    alert(`คัดลอกหมายเลขคำสั่งซื้อ ${text} แล้ว!`);
  });
};

// Filter tabs click
filterTabs.forEach(btn => {
  btn.addEventListener('click', () => {
    filterTabs.forEach(b => {
      b.classList.remove('bg-white', 'border', 'border-slate-200', 'text-slate-800', 'shadow-sm', 'font-semibold');
      b.classList.add('font-medium');
    });
    btn.classList.add('bg-white', 'border', 'border-slate-200', 'text-slate-800', 'shadow-sm', 'font-semibold');
    btn.classList.remove('font-medium');
    activeFilter = btn.getAttribute('data-filter');
    renderTable();
  });
});

// Search input
searchInput.addEventListener('input', (e) => {
  searchQuery = e.target.value.trim();
  renderTable();
});

// Copy all anomalous order IDs
btnCopyAnomalies.addEventListener('click', () => {
  const anomalies = reconResults.filter(r => r.status === 'GHOST' || r.status === 'OVERDUE' || r.status === 'MISMATCH');
  if (anomalies.length === 0) {
    alert('ไม่พบรายการที่มีปัญหาในชุดข้อมูลนี้ครับ');
    return;
  }
  const text = anomalies.map(r => `${r.orderId} \t (${r.status}: ${r.notes})`).join('\n');
  navigator.clipboard.writeText(text).then(() => {
    alert(`คัดลอก Order ID ที่ผิดปกติทั้งหมด ${anomalies.length} รายการ เรียบร้อยแล้ว! นำไปวางใน Ticket ซัพพอร์ตได้ทันที`);
  });
});

// Export to Excel
btnExportExcel.addEventListener('click', () => {
  if (reconResults.length === 0) {
    alert('ไม่มีข้อมูลให้ส่งออก');
    return;
  }

  const wb = XLSX.utils.book_new();

  // Summary sheet
  const summaryData = [
    ['รายงานการตรวจสอบและกระทบยอดเงินเข้า (Platform Payout Reconciler)', ''],
    ['วันที่สร้างรายงาน', new Date().toLocaleString('th-TH')],
    ['', ''],
    ['รายการ', 'จำนวน', 'ยอดเงิน (THB)'],
    ['ยอดเงินโอนเข้าจริง (Income)', incomeData.length, incomeData.reduce((a, b) => a + b.actualAmount, 0)],
    ['ตรงกันตามนัด (Matched)', reconResults.filter(r => r.status === 'MATCHED').length, reconResults.filter(r => r.status === 'MATCHED').reduce((a, b) => a + b.actualAmount, 0)],
    ['🚨 เงินเข้าไม่ตรงนัด (Ghost Payouts)', reconResults.filter(r => r.status === 'GHOST').length, reconResults.filter(r => r.status === 'GHOST').reduce((a, b) => a + b.actualAmount, 0)],
    ['⏳ ถึงกำหนดแต่ไม่โอน (Overdue/Held)', reconResults.filter(r => r.status === 'OVERDUE').length, reconResults.filter(r => r.status === 'OVERDUE').reduce((a, b) => a + b.estimatedAmount, 0)],
    ['⚠️ ยอดเงินไม่ตรง (Discrepancy)', reconResults.filter(r => r.status === 'MISMATCH').length, reconResults.filter(r => r.status === 'MISMATCH').reduce((a, b) => a + Math.abs(b.diff), 0)]
  ];
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, wsSummary, 'สรุปภาพรวม');

  // Anomalies Sheet (Ghost + Overdue + Mismatch)
  const anomalies = reconResults.filter(r => r.status !== 'MATCHED').map(r => ({
    'สถานะ': r.status === 'GHOST' ? 'เงินเข้าไม่ตรงนัด (Ghost)' : (r.status === 'OVERDUE' ? 'ถึงกำหนดแต่ไม่โอน' : 'ยอดไม่ตรง'),
    'หมายเลขคำสั่งซื้อ': r.orderId,
    'ประเภท': r.type,
    'ยอดรอเข้า (ประมาณการ)': r.estimatedAmount,
    'ยอดโอนจริง': r.actualAmount,
    'ส่วนต่าง': r.diff,
    'กำหนดชำระ': r.payoutTime,
    'เหตุผลในระบบ': r.reason,
    'หมายเหตุ/ข้อสังเกต': r.notes
  }));
  if (anomalies.length > 0) {
    const wsAnomalies = XLSX.utils.json_to_sheet(anomalies);
    XLSX.utils.book_append_sheet(wb, wsAnomalies, 'รายการผิดปกติ_สำหรับเปิดตั๋ว');
  }

  // All Items Sheet
  const allData = reconResults.map(r => ({
    'สถานะ': r.status,
    'หมายเลขคำสั่งซื้อ': r.orderId,
    'ประเภท': r.type,
    'ยอดรอเข้า (ประมาณการ)': r.estimatedAmount,
    'ยอดโอนจริง': r.actualAmount,
    'ส่วนต่าง': r.diff,
    'กำหนดชำระ': r.payoutTime,
    'เหตุผล': r.reason,
    'วันที่สร้าง': r.orderCreatedDate,
    'หมายเหตุ': r.notes
  }));
  const wsAll = XLSX.utils.json_to_sheet(allData);
  XLSX.utils.book_append_sheet(wb, wsAll, 'รายการทั้งหมด');

  XLSX.writeFile(wb, `รายงานกระทบยอดเงินเข้า_${new Date().toISOString().slice(0, 10)}.xlsx`);
});

// Reset Button
btnReset.addEventListener('click', () => {
  pendingData = [];
  incomeData = [];
  reconResults = [];
  pendingWorkbook = null;
  incomeWorkbook = null;
  filePendingInput.value = '';
  fileIncomeInput.value = '';
  pendingFileInfo.classList.add('hidden');
  incomeFileInfo.classList.add('hidden');
  metricsPanel.classList.add('hidden');
  resultsPanel.classList.add('hidden');
  document.getElementById('badge-pending').innerText = 'ยังไม่เลือกไฟล์';
  document.getElementById('badge-pending').className = 'text-xs text-slate-400';
  document.getElementById('badge-income').innerText = 'ยังไม่เลือกไฟล์';
  document.getElementById('badge-income').className = 'text-xs text-slate-400';
});

// Load Demo Dataset (Case วันที่ 22)
btnLoadDemo.addEventListener('click', () => {
  // 1. Simulate real pending data from September 2569 sheet
  pendingData = [
    { orderId: '586196895880873180', estimatedAmount: 1391.23, estimatedPayoutTime: 'ส่งมอบแล้ว + 3 วัน', reason: 'รอการจัดส่งพัสดุ', type: 'คำสั่งซื้อ', orderCreatedDate: '2026/09/22' },
    { orderId: '586196843500111234', estimatedAmount: 1798.36, estimatedPayoutTime: 'ส่งมอบแล้ว + 3 วัน', reason: 'รอการจัดส่งพัสดุ', type: 'คำสั่งซื้อ', orderCreatedDate: '2026/09/22' },
    { orderId: '585822627988932574', estimatedAmount: 1653.67, estimatedPayoutTime: '2026/09/22', reason: 'จัดส่งสำเร็จแล้ว รอการชำระเงิน', type: 'คำสั่งซื้อ', orderCreatedDate: '2026/09/18' },
    { orderId: '585822816578405579', estimatedAmount: 1728.46, estimatedPayoutTime: '2026/09/22', reason: 'จัดส่งสำเร็จแล้ว รอการชำระเงิน', type: 'คำสั่งซื้อ', orderCreatedDate: '2026/09/18' },
    { orderId: '585819127154575208', estimatedAmount: 1668.47, estimatedPayoutTime: '2026/09/22', reason: 'จัดส่งสำเร็จแล้ว รอการชำระเงิน', type: 'คำสั่งซื้อ', orderCreatedDate: '2026/09/18' },
    { orderId: '585817701549901332', estimatedAmount: 2698.48, estimatedPayoutTime: '2026/09/23', reason: 'จัดส่งสำเร็จแล้ว รอการชำระเงิน', type: 'คำสั่งซื้อ', orderCreatedDate: '2026/09/19' },
    { orderId: '585817700378249085', estimatedAmount: 1670.07, estimatedPayoutTime: '2026/09/23', reason: 'จัดส่งสำเร็จแล้ว รอการชำระเงิน', type: 'คำสั่งซื้อ', orderCreatedDate: '2026/09/19' }
  ];

  // 2. Simulate Income file for the 22nd evening payout
  // Notice: Includes a Ghost payout (unlisted order released after ticket!), a matched order, and an amount mismatch!
  incomeData = [
    { orderId: '585822627988932574', actualAmount: 1653.67, payoutTime: '2026/09/22 17:45:00', type: 'คำสั่งซื้อ' },
    { orderId: '585822816578405579', actualAmount: 1580.00, payoutTime: '2026/09/22 17:45:00', type: 'คำสั่งซื้อ' }, // mismatch
    // GHOST ORDER (No pending schedule for the 22nd, released after ticket!)
    { orderId: '585799120485901233', actualAmount: 3450.50, payoutTime: '2026/09/22 17:45:00', type: 'คำสั่งซื้อ (ปลดล็อกหลัง Ticket)' },
    { orderId: '585788910234125890', actualAmount: 1200.00, payoutTime: '2026/09/22 17:45:00', type: 'รายการปรับยอด/ชดเชย' }
  ];

  document.getElementById('pending-filename').innerText = 'ตัวอย่าง_TikTok_รอชำระเงิน_9.2569.xlsx';
  document.getElementById('pending-stats').innerText = `${pendingData.length} รายการ`;
  pendingFileInfo.classList.remove('hidden');
  document.getElementById('badge-pending').innerText = 'ข้อมูลจำลอง';

  document.getElementById('income-filename').innerText = 'ตัวอย่าง_TikTok_Income_22_ก.ย..xlsx';
  document.getElementById('income-stats').innerText = `${incomeData.length} รายการ`;
  incomeFileInfo.classList.remove('hidden');
  document.getElementById('badge-income').innerText = 'ข้อมูลจำลอง';

  runReconciliation();
  alert('โหลดตัวอย่างข้อมูลจำลองจำลองเคสวันที่ 22 เรียบร้อยแล้ว! ลองดูในตารางและกดสลับแท็บเพื่อดูรายการ Ghost, Overdue และ Mismatch');
});
