/**
 * Side Panel 逻辑
 * 职责：Excel 加载、列选择、专利列表、操作控制
 */

// ========== DOM 元素 ==========
const els = {
  loadCard: document.getElementById('loadCard'),
  fileDrop: document.getElementById('fileDrop'),
  fileInput: document.getElementById('fileInput'),
  columnCard: document.getElementById('columnCard'),
  columnSelect: document.getElementById('columnSelect'),
  btnConfirmColumn: document.getElementById('btnConfirmColumn'),
  mainCard: document.getElementById('mainCard'),
  currentPatent: document.getElementById('currentPatent'),
  progressLabel: document.getElementById('progressLabel'),
  progressPercent: document.getElementById('progressPercent'),
  progressFill: document.getElementById('progressFill'),
  btnFill: document.getElementById('btnFill'),
  btnPrint: document.getElementById('btnPrint'),
  btnReset: document.getElementById('btnReset'),
  btnNext: document.getElementById('btnNext'),
  btnExtractFees: document.getElementById('btnExtractFees'),
  btnExtractAccumulate: document.getElementById('btnExtractAccumulate'),
  btnExportAccumulated: document.getElementById('btnExportAccumulated'),
  btnExportExcelDue: document.getElementById('btnExportExcelDue'),
  btnExportExcelPaid: document.getElementById('btnExportExcelPaid'),
  btnExportExcelPayment: document.getElementById('btnExportExcelPayment'),
  btnExportExcelDueAll: document.getElementById('btnExportExcelDueAll'),
  accumulatedCount: document.getElementById('accumulatedCount'),
  feeCard: document.getElementById('feeCard'),
  feeContent: document.getElementById('feeContent'),
  btnExportJson: document.getElementById('btnExportJson'),
  listCard: document.getElementById('listCard'),
  patentList: document.getElementById('patentList'),
  modeAppNo: document.getElementById('modeAppNo'),
  modePubNo: document.getElementById('modePubNo'),
  btnSettings: document.getElementById('btnSettings'),
  btnBackMain: document.getElementById('btnBackMain'),
  settingsPage: document.getElementById('settingsPage'),
  settingsAccumulatedCount: document.getElementById('settingsAccumulatedCount'),
};

let workbook = null;
let allColumns = [];
let allRows = [];
let confirmMode = 'column'; // 'column' | 'sheet'
let currentSearchMode = 'appNo'; // 'appNo' | 'pubNo'

// ========== 初始化 ==========
async function init() {
  bindEvents();
  await restoreState();
}

function bindEvents() {
  // 文件拖拽
  els.fileDrop.addEventListener('click', () => els.fileInput.click());
  els.fileDrop.addEventListener('dragover', (e) => {
    e.preventDefault();
    els.fileDrop.classList.add('dragover');
  });
  els.fileDrop.addEventListener('dragleave', () => {
    els.fileDrop.classList.remove('dragover');
  });
  els.fileDrop.addEventListener('drop', (e) => {
    e.preventDefault();
    els.fileDrop.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });
  els.fileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) handleFile(e.target.files[0]);
  });

  // Sheet 选择时立即解析显示列信息
  els.columnSelect.addEventListener('change', () => {
    if (confirmMode === 'sheet') {
      const idx = parseInt(els.columnSelect.value);
      if (!isNaN(idx)) {
        parseSheet(pendingSheetNames[idx]);
      }
    }
  });

  // 列选择确认
  els.btnConfirmColumn.addEventListener('click', () => {
    if (confirmMode === 'sheet') handleConfirmSheet();
    else handleConfirmColumn();
  });

  // 检索方式切换
  els.modeAppNo.addEventListener('click', () => setSearchMode('appNo'));
  els.modePubNo.addEventListener('click', () => setSearchMode('pubNo'));

  // 设置页面导航
  els.btnSettings.addEventListener('click', showSettings);
  els.btnBackMain.addEventListener('click', showMain);

  // 操作按钮
  els.btnFill.addEventListener('click', handleFill);
  els.btnPrint.addEventListener('click', handlePrint);
  els.btnReset.addEventListener('click', handleReset);
  els.btnNext.addEventListener('click', handleNextPatent);
  els.btnExtractFees.addEventListener('click', handleExtractFees);
  els.btnExportJson.addEventListener('click', handleExportJson);
  els.btnExtractAccumulate.addEventListener('click', handleExtractAccumulate);
  els.btnExportAccumulated.addEventListener('click', handleExportAccumulated);
  els.btnExportExcelDue.addEventListener('click', handleExportExcelDue);
  els.btnExportExcelPaid.addEventListener('click', handleExportExcelPaid);
  els.btnExportExcelPayment.addEventListener('click', handleExportExcelPayment);
  els.btnExportExcelDueAll.addEventListener('click', handleExportExcelDueAll);

  // 监听状态变化
  chrome.storage.onChanged.addListener((changes) => {
    if (changes.patentQueryState) {
      updateUI(changes.patentQueryState.newValue);
    }
  });

}

// ========== 页面导航 ==========
function showSettings() {
  document.querySelector('.content').classList.add('hidden');
  els.settingsPage.classList.remove('hidden');
  // 同步累积计数
  els.settingsAccumulatedCount.textContent = els.accumulatedCount.textContent;
}

function showMain() {
  els.settingsPage.classList.add('hidden');
  document.querySelector('.content').classList.remove('hidden');
}

// ========== 检索方式切换 ==========
function setSearchMode(mode) {
  currentSearchMode = mode;
  chrome.runtime.sendMessage({ action: 'SET_SEARCH_MODE', data: { mode } });
  updateModeUI(mode);
}

function updateModeUI(mode) {
  if (mode === 'pubNo') {
    els.modePubNo.style.background = '#1a73e8';
    els.modePubNo.style.color = 'white';
    els.modeAppNo.style.background = '#fff';
    els.modeAppNo.style.color = '#555';
    els.btnFill.textContent = '📝 填入公告号';
  } else {
    els.modeAppNo.style.background = '#1a73e8';
    els.modeAppNo.style.color = 'white';
    els.modePubNo.style.background = '#fff';
    els.modePubNo.style.color = '#555';
    els.btnFill.textContent = '📝 填入专利号';
  }
}

// ========== 文件处理 ==========
function handleFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: 'array' });

      // 多 sheet 时让用户选择，单 sheet 直接解析
      if (workbook.SheetNames.length > 1) {
        showSheetSelector(workbook.SheetNames);
      } else {
        parseSheet(workbook.SheetNames[0]);
      }
    } catch (err) {
      showToast('文件解析失败: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

// ========== Sheet 选择 ==========
let pendingSheetNames = [];
let isMultiSheet = false;

function showSheetSelector(sheetNames) {
  pendingSheetNames = sheetNames;
  isMultiSheet = true;
  confirmMode = 'sheet';
  els.columnCard.classList.remove('hidden');
  els.columnCard.querySelector('.card-title').innerHTML = '📑 选择工作表 <button id="btnBackSheet" style="margin-left: auto; background: none; border: none; cursor: pointer; font-size: 14px; padding: 0 4px; color: #1a73e8;" data-tip="重选工作表">← 重选工作表</button>';
  document.getElementById('btnBackSheet').addEventListener('click', () => {
    showSheetSelector(pendingSheetNames);
  });
  els.columnSelect.innerHTML = '<option value="">请选择工作表</option>';
  sheetNames.forEach((name, i) => {
    const option = document.createElement('option');
    option.value = i;
    option.textContent = name;
    els.columnSelect.appendChild(option);
  });
}

function handleConfirmSheet() {
  const idx = parseInt(els.columnSelect.value);
  if (isNaN(idx)) { showToast('请选择一个工作表', 'error'); return; }
  confirmMode = 'column';
  parseSheet(pendingSheetNames[idx]);
}

// ========== 解析 Sheet ==========
function parseSheet(sheetName) {
  confirmMode = 'column';
  const sheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  if (jsonData.length < 2) {
    showToast('工作表至少需要表头行和一行数据', 'error');
    return;
  }

  // 自动检测表头行：找第一个有 3 个以上非空单元格的行
  let headerRowIndex = -1;
  for (let i = 0; i < Math.min(jsonData.length, 20); i++) {
    const nonNullCount = (jsonData[i] || []).filter(v => v !== null && v !== undefined && String(v).trim() !== '').length;
    if (nonNullCount >= 3) {
      headerRowIndex = i;
      break;
    }
  }

  if (headerRowIndex === -1) {
    showToast('未找到有效的表头行', 'error');
    return;
  }

  // 处理合并单元格：将 merged cell 的值传播到整个合并区域
  const merges = sheet['!merges'] || [];
  const headerRow = jsonData[headerRowIndex] || [];
  const filledHeader = [...headerRow];

  for (const merge of merges) {
    // 合并区域覆盖了表头行
    if (merge.s.r <= headerRowIndex && merge.e.r >= headerRowIndex) {
      const topLeftValue = getCell(sheet, merge.s.r, merge.s.c);
      for (let col = merge.s.c; col <= merge.e.c; col++) {
        if (!filledHeader[col] || String(filledHeader[col]).trim() === '') {
          filledHeader[col] = topLeftValue;
        }
      }
    }
  }

  // 提取列名和数据
  allColumns = filledHeader.map((col, i) => ({
    name: String(col || `列${i + 1}`).trim(),
    index: i
  }));
  allRows = jsonData.slice(headerRowIndex + 1);

  // 过滤全空行
  allRows = allRows.filter(row => row && row.some(v => v !== null && v !== undefined && String(v).trim() !== ''));

  // 填充列选择下拉框
  if (isMultiSheet) {
    els.columnCard.querySelector('.card-title').innerHTML = '📌 选择专利号列 <button id="btnBackSheet" style="margin-left: auto; background: none; border: none; cursor: pointer; font-size: 14px; padding: 0 4px; color: #1a73e8;" data-tip="重选工作表">← 重选工作表</button>';
    document.getElementById('btnBackSheet').addEventListener('click', () => {
      showSheetSelector(pendingSheetNames);
    });
  } else {
    els.columnCard.querySelector('.card-title').textContent = '📌 选择专利号列';
  }
  els.columnSelect.innerHTML = '<option value="">请选择包含专利号的列</option>';
  allColumns.forEach((col) => {
    const option = document.createElement('option');
    option.value = col.index;
    option.textContent = col.name;
    els.columnSelect.appendChild(option);
  });

  // 智能猜测：找包含"专利号""申请号""注册号"的列
  const guessIndex = allColumns.findIndex((col) =>
    /专利号|申请号|注册号|patent/i.test(col.name)
  );
  if (guessIndex >= 0) {
    els.columnSelect.value = allColumns[guessIndex].index;
  }

  els.columnCard.classList.remove('hidden');
  showToast(`已加载 ${allRows.length} 行数据，共 ${allColumns.length} 列`);
}

// ========== 获取单元格值 ==========
function getCell(sheet, row, col) {
  const addr = XLSX.utils.encode_cell({ r: row, c: col });
  const cell = sheet[addr];
  return cell ? cell.v : null;
}

// ========== 列选择确认 ==========
async function handleConfirmColumn() {
  const colIndex = parseInt(els.columnSelect.value);
  if (isNaN(colIndex)) {
    showToast('请选择一列', 'error');
    return;
  }

  const columnName = allColumns[colIndex].name;

  // 提取专利号（过滤空值）
  const patents = allRows
    .map((row) => row[colIndex])
    .filter((val) => val !== undefined && val !== null && String(val).trim() !== '')
    .map((val) => String(val).trim());

  if (patents.length === 0) {
    showToast('所选列中没有找到数据', 'error');
    return;
  }

  // 发送到 background
  const resp = await chrome.runtime.sendMessage({
    action: 'LOAD_PATENTS',
    data: { patents, columnName }
  });

  if (resp?.success) {
    showToast(`已加载 ${patents.length} 个专利号`);
    els.mainCard.classList.remove('hidden');
    els.listCard.classList.remove('hidden');
    updateUI({ patents, currentIndex: 0 });
  }
}

// ========== 填入专利号 ==========
async function handleFill() {
  try {
    els.btnFill.disabled = true;
    els.btnFill.textContent = '⏳ 填入中...';
    const resp = await chrome.runtime.sendMessage({ action: 'FILL_PATENT' });
    if (resp?.error) {
      showToast(resp.error, 'error');
    } else {
      showToast(currentSearchMode === 'pubNo' ? '公告号已填入' : '专利号已填入');
    }
  } catch (e) {
    showToast('填入失败: ' + e.message, 'error');
  } finally {
    els.btnFill.disabled = false;
    els.btnFill.textContent = currentSearchMode === 'pubNo' ? '📝 填入公告号' : '📝 填入专利号';
  }
}

// ========== 原生打印 ==========
async function handlePrint() {
  try {
    els.btnPrint.disabled = true;
    els.btnPrint.textContent = '⏳ 打印中...';

    // 在用户点击的上下文中复制文件名（需要用户手势）
    const stateResp = await chrome.runtime.sendMessage({ action: 'GET_STATE' });
    if (stateResp?.success) {
      const { patents, currentIndex } = stateResp.data;
      const patent = patents[currentIndex] || '';
      const seq = String(currentIndex + 1).padStart(3, '0');
      const filename = `${seq}-${patent}`;
      await navigator.clipboard.writeText(filename);
    }

    // 触发打印
    const resp = await chrome.runtime.sendMessage({ action: 'PRINT_PAGE' });
    if (resp?.error) {
      showToast(resp.error, 'error');
    } else {
      // 打印对话框已关闭，自动标记并提示可撤销
      const stateResp2 = await chrome.runtime.sendMessage({ action: 'GET_STATE' });
      const { patents: p, currentIndex: ci } = stateResp2.data;
      const patent = p[ci] || '';
      await chrome.runtime.sendMessage({ action: 'MARK_PRINTED', data: { patent } });
      showToast('已标记为已打印', 'success', {
        label: '撤销',
        callback: async () => {
          await chrome.runtime.sendMessage({ action: 'UNMARK_PRINTED', data: { patent } });
        }
      });
    }
  } catch (e) {
    showToast('打印失败: ' + e.message, 'error');
  } finally {
    els.btnPrint.disabled = false;
    els.btnPrint.textContent = '🖨️ 打印页面';
  }
}

// ========== 提取费用信息 ==========
let lastFeeData = null;

async function handleExtractFees() {
  try {
    els.btnExtractFees.disabled = true;
    els.btnExtractFees.textContent = '⏳ 提取中...';
    const resp = await chrome.runtime.sendMessage({ action: 'EXTRACT_FEES' });
    if (resp?.error) {
      showToast(resp.error, 'error');
    } else {
      lastFeeData = resp.data;
      renderFeeInfo(resp.data);
      showToast('费用信息提取成功');
    }
  } catch (e) {
    showToast('提取失败: ' + e.message, 'error');
  } finally {
    els.btnExtractFees.disabled = false;
    els.btnExtractFees.textContent = '💰 提取费用信息';
  }
}

function renderFeeInfo(data) {
  els.feeCard.classList.remove('hidden');
  let html = '';

  // 基本信息
  if (data.basicInfo && Object.keys(data.basicInfo).length > 0) {
    html += '<div style="margin-bottom: 10px;"><b>专利基本信息</b></div>';
    for (const [key, val] of Object.entries(data.basicInfo)) {
      html += `<div style="color: #555;">${escapeHtml(key)}: ${escapeHtml(val)}</div>`;
    }
  }

  // 应缴费用
  if (data.feesDue && data.feesDue.length > 0) {
    html += '<div style="margin: 10px 0 6px;"><b>应缴费用</b> (共 ' + data.feesDue.length + ' 项)</div>';
    html += '<table style="width: 100%; border-collapse: collapse; font-size: 11px;">';
    html += '<tr style="background: #f5f5f5;"><th style="padding: 4px; border: 1px solid #e8e8e8; text-align: left;">费用种类</th><th style="padding: 4px; border: 1px solid #e8e8e8; text-align: right;">金额</th><th style="padding: 4px; border: 1px solid #e8e8e8; text-align: center;">缴费截止日</th><th style="padding: 4px; border: 1px solid #e8e8e8; text-align: center;">费用状态</th></tr>';
    for (const fee of data.feesDue) {
      html += `<tr><td style="padding: 4px; border: 1px solid #e8e8e8;">${escapeHtml(fee.feeType)}</td><td style="padding: 4px; border: 1px solid #e8e8e8; text-align: right;">${escapeHtml(fee.amount)}</td><td style="padding: 4px; border: 1px solid #e8e8e8; text-align: center;">${escapeHtml(fee.deadline)}</td><td style="padding: 4px; border: 1px solid #e8e8e8; text-align: center;">${escapeHtml(fee.status)}</td></tr>`;
    }
    html += '</table>';
  }

  // 已缴费用
  if (data.feesPaid && data.feesPaid.length > 0) {
    html += '<div style="margin: 10px 0 6px;"><b>已缴费用</b> (共 ' + data.feesPaid.length + ' 项)</div>';
    html += '<table style="width: 100%; border-collapse: collapse; font-size: 11px;">';
    html += '<tr style="background: #f5f5f5;"><th style="padding: 4px; border: 1px solid #e8e8e8; text-align: left;">费用类别</th><th style="padding: 4px; border: 1px solid #e8e8e8; text-align: right;">应缴金额</th><th style="padding: 4px; border: 1px solid #e8e8e8; text-align: center;">缴费日期</th><th style="padding: 4px; border: 1px solid #e8e8e8; text-align: left;">缴费人姓名</th><th style="padding: 4px; border: 1px solid #e8e8e8; text-align: left;">票据代码</th><th style="padding: 4px; border: 1px solid #e8e8e8; text-align: left;">票据号码</th></tr>';
    for (const fee of data.feesPaid) {
      html += `<tr><td style="padding: 4px; border: 1px solid #e8e8e8;">${escapeHtml(fee.feeType)}</td><td style="padding: 4px; border: 1px solid #e8e8e8; text-align: right;">${escapeHtml(fee.amount)}</td><td style="padding: 4px; border: 1px solid #e8e8e8; text-align: center;">${escapeHtml(fee.payDate)}</td><td style="padding: 4px; border: 1px solid #e8e8e8;">${escapeHtml(fee.payer)}</td><td style="padding: 4px; border: 1px solid #e8e8e8;">${escapeHtml(fee.invoiceCode)}</td><td style="padding: 4px; border: 1px solid #e8e8e8;">${escapeHtml(fee.invoiceNo)}</td></tr>`;
    }
    html += '</table>';
  }

  if (!html) {
    html = '<div style="color: #999; text-align: center; padding: 16px;">未识别到费用信息，请确保已打开费用查询页面</div>';
  }

  els.feeContent.innerHTML = html;
}

function handleExportJson() {
  if (!lastFeeData) {
    showToast('请先提取费用信息', 'error');
    return;
  }
  const json = JSON.stringify(lastFeeData, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const patentNo = lastFeeData.basicInfo?.patentNumber || 'unknown';
  a.download = `${patentNo}_费用信息.json`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('JSON 已导出');
}

// ========== 提取并累积 ==========
async function handleExtractAccumulate() {
  try {
    els.btnExtractAccumulate.disabled = true;
    els.btnExtractAccumulate.textContent = '⏳ 提取中...';
    const resp = await chrome.runtime.sendMessage({ action: 'EXTRACT_AND_ACCUMULATE' });
    if (resp?.error) {
      showToast(resp.error, 'error');
    } else {
      lastFeeData = resp.data;
      renderFeeInfo(resp.data);
      const accText = `已累积 ${resp.accumulatedCount} 个专利`;
      els.accumulatedCount.textContent = accText;
      els.settingsAccumulatedCount.textContent = accText;

      // 自动跳转到下一个专利
      const stateResp = await chrome.runtime.sendMessage({ action: 'GET_STATE' });
      const prevIndex = stateResp?.data?.currentIndex ?? 0;
      const prevPatent = stateResp?.data?.patents[prevIndex] || '';
      await chrome.runtime.sendMessage({ action: 'NEXT_PATENT' });
      await restoreState();

      showToast(`已累积 ${resp.accumulatedCount} 个专利，已自动跳转`, 'success', {
        label: '返回上一个',
        callback: async () => {
          await chrome.runtime.sendMessage({ action: 'SET_INDEX', data: { index: prevIndex } });
          showToast(`已返回 ${prevPatent}`);
        }
      });
    }
  } catch (e) {
    showToast('提取失败: ' + e.message, 'error');
  } finally {
    els.btnExtractAccumulate.disabled = false;
    els.btnExtractAccumulate.textContent = '📦 提取并累积';
  }
}

// ========== 导出累积 JSON ==========
async function handleExportAccumulated() {
  try {
    const resp = await chrome.runtime.sendMessage({ action: 'GET_ACCUMULATED' });
    if (!resp?.success || Object.keys(resp.data).length === 0) {
      showToast('暂无累积数据，请先提取并累积', 'error');
      return;
    }
    const json = JSON.stringify(resp.data, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const now = new Date();
    const dateStr = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
    a.download = `专利费用信息_${dateStr}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast(`已导出 ${Object.keys(resp.data).length} 个专利的费用信息`);
  } catch (e) {
    showToast('导出失败: ' + e.message, 'error');
  }
}

// ========== 导出应缴年费 Excel ==========
async function handleExportExcelDue() {
  try {
    const resp = await chrome.runtime.sendMessage({ action: 'GET_ACCUMULATED' });
    if (!resp?.success || Object.keys(resp.data).length === 0) {
      showToast('暂无累积数据，请先提取并累积', 'error');
      return;
    }

    const rows = [['序号', '专利申请号', '应缴费用项目', '应缴费用数额', '缴费期限']];
    let idx = 1;
    for (const [patentNo, info] of Object.entries(resp.data)) {
      const due = info.feesDue?.[0];
      if (!due) continue;
      rows.push([
        idx,
        patentNo,
        due.feeType || '',
        due.amount || '',
        due.deadline || ''
      ]);
      idx++;
    }

    if (rows.length === 1) {
      showToast('没有应缴费用数据可导出', 'error');
      return;
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '应缴年费');
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    XLSX.writeFile(wb, `专利最近应缴费用_${dateStr}.xlsx`);
    showToast(`已导出 ${rows.length - 1} 条应缴年费记录`);
  } catch (e) {
    showToast('导出失败: ' + e.message, 'error');
  }
}

// ========== 导出成本核算 Excel ==========
async function handleExportExcelPaid() {
  try {
    const resp = await chrome.runtime.sendMessage({ action: 'GET_ACCUMULATED' });
    if (!resp?.success || Object.keys(resp.data).length === 0) {
      showToast('暂无累积数据，请先提取并累积', 'error');
      return;
    }

    const rows = [['序号', '专利申请号', '已缴费用项目', '已缴费用数额', '缴费日期']];
    const merges = [];
    let idx = 1;

    for (const [patentNo, info] of Object.entries(resp.data)) {
      const feesPaid = info.feesPaid || [];
      if (feesPaid.length === 0) continue;

      const startRow = rows.length;
      feesPaid.forEach((fee) => {
        rows.push([
          idx,
          patentNo,
          fee.feeType || '',
          fee.amount || '',
          fee.payDate || ''
        ]);
      });

      if (feesPaid.length > 1) {
        merges.push({
          s: { r: startRow, c: 0 },
          e: { r: startRow + feesPaid.length - 1, c: 0 }
        });
        merges.push({
          s: { r: startRow, c: 1 },
          e: { r: startRow + feesPaid.length - 1, c: 1 }
        });
      }
      idx++;
    }

    if (rows.length === 1) {
      showToast('没有已缴费用数据可导出', 'error');
      return;
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    if (merges.length > 0) {
      ws['!merges'] = merges;
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '成本核算');
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    XLSX.writeFile(wb, `专利成本核算_${dateStr}.xlsx`);
    showToast(`已导出 ${rows.length - 1} 条已缴费用记录`);
  } catch (e) {
    showToast('导出失败: ' + e.message, 'error');
  }
}

// ========== 导出网上缴费 Excel ==========
async function handleExportExcelPayment() {
  try {
    const resp = await chrome.runtime.sendMessage({ action: 'GET_ACCUMULATED' });
    if (!resp?.success || Object.keys(resp.data).length === 0) {
      showToast('暂无累积数据，请先提取并累积', 'error');
      return;
    }

    const rows = [['序号', '申请号/专利号/国际申请号/海牙转交编号', '业务类型', '票据抬头', '统一社会信用代码', '费用种类', '外币金额', '费用金额（人民币）', '备注']];
    let idx = 1;

    for (const [patentNo, info] of Object.entries(resp.data)) {
      const feesDue = info.feesDue || [];
      const normalizedNo = normalizePatentNumber(info.basicInfo?.patentNumber || patentNo);
      const annualFee = feesDue.find(f => (f.feeType || '').includes('年费'));

      if (annualFee) {
        const patentType = inferPatentType(annualFee.feeType);
        rows.push([
          idx,
          normalizedNo,
          patentType,
          '', // 票据抬头
          '', // 统一社会信用代码
          annualFee.feeType || '',
          '', // 外币金额
          annualFee.amount || '',
          ''  // 备注
        ]);
      } else {
        rows.push([
          idx,
          normalizedNo,
          '', // 业务类型
          '', // 票据抬头
          '', // 统一社会信用代码
          '', // 费用种类
          '', // 外币金额
          '', // 费用金额
          '无应缴年费'
        ]);
      }
      idx++;
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '网上缴费');
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    XLSX.writeFile(wb, `专利应缴年费_网上缴费标准格式_${dateStr}.xlsx`);
    showToast(`已导出 ${rows.length - 1} 条网上缴费记录`);
  } catch (e) {
    showToast('导出失败: ' + e.message, 'error');
  }
}

// ========== 导出全部应缴费用 Excel ==========
async function handleExportExcelDueAll() {
  try {
    const resp = await chrome.runtime.sendMessage({ action: 'GET_ACCUMULATED' });
    if (!resp?.success || Object.keys(resp.data).length === 0) {
      showToast('暂无累积数据，请先提取并累积', 'error');
      return;
    }

    const rows = [['序号', '专利申请号', '应缴费用项目', '应缴费用数额', '缴费期限']];
    const merges = [];
    let idx = 1;

    for (const [patentNo, info] of Object.entries(resp.data)) {
      const feesDue = info.feesDue || [];

      if (feesDue.length === 0) {
        rows.push([idx, patentNo, '无', '', '']);
        idx++;
        continue;
      }

      const startRow = rows.length;
      feesDue.forEach((fee) => {
        rows.push([
          idx,
          patentNo,
          fee.feeType || '',
          fee.amount || '',
          fee.deadline || ''
        ]);
      });

      if (feesDue.length > 1) {
        merges.push({
          s: { r: startRow, c: 0 },
          e: { r: startRow + feesDue.length - 1, c: 0 }
        });
        merges.push({
          s: { r: startRow, c: 1 },
          e: { r: startRow + feesDue.length - 1, c: 1 }
        });
      }
      idx++;
    }

    if (rows.length === 1) {
      showToast('没有数据可导出', 'error');
      return;
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    if (merges.length > 0) {
      ws['!merges'] = merges;
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '全部应缴费用');
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    XLSX.writeFile(wb, `专利全部应缴费用_${dateStr}.xlsx`);
    showToast(`已导出 ${Object.keys(resp.data).length} 个专利的应缴费用记录`);
  } catch (e) {
    showToast('导出失败: ' + e.message, 'error');
  }
}

function inferPatentType(feeType) {
  if (!feeType) return '';
  if (feeType.includes('实用新型')) return '实用新型';
  if (feeType.includes('外观设计')) return '外观设计';
  if (feeType.includes('发明')) return '发明';
  return '';
}

function normalizePatentNumber(pn) {
  if (!pn) return '';
  let normalized = pn.replace(/^CN/i, '');
  normalized = normalized.replace(/\./g, '');
  normalized = normalized.replace(/\D/g, '');
  return normalized;
}

// ========== 工具函数 ==========
function escapeHtml(text) {
  if (text === null || text === undefined) return '';
  const div = document.createElement('div');
  div.textContent = String(text);
  return div.innerHTML;
}

// ========== 重置插件 ==========
async function handleReset() {
  // 1. 清除后台状态
  await chrome.runtime.sendMessage({ action: 'RESET_ALL' });

  // 2. 清除本地文件与解析状态
  workbook = null;
  allColumns = [];
  allRows = [];
  lastFeeData = null;

  // 3. 重置 UI 到初始模式
  els.mainCard.classList.add('hidden');
  els.listCard.classList.add('hidden');
  els.columnCard.classList.add('hidden');
  els.feeCard.classList.add('hidden');
  els.settingsPage.classList.add('hidden');
  els.loadCard.classList.remove('hidden');
  document.querySelector('.content').classList.remove('hidden');

  // 4. 清空显示内容
  els.patentList.innerHTML = '';
  els.currentPatent.textContent = '-';
  els.progressLabel.textContent = '0 / 0';
  els.progressPercent.textContent = '0%';
  els.progressFill.style.width = '0%';
  els.accumulatedCount.textContent = '';
  els.settingsAccumulatedCount.textContent = '';
  els.fileInput.value = '';

  // 重置检索方式为申请号
  currentSearchMode = 'appNo';
  updateModeUI('appNo');

  showToast('插件已重置，请重新上传表格');
}

// ========== 下一个专利 ==========
async function handleNextPatent() {
  try {
    els.btnNext.disabled = true;
    els.btnNext.textContent = '⏳ 切换中...';
    const stateResp = await chrome.runtime.sendMessage({ action: 'GET_STATE' });
    const total = stateResp?.data?.patents?.length || 0;
    const oldIndex = stateResp?.data?.currentIndex ?? 0;

    const resp = await chrome.runtime.sendMessage({ action: 'NEXT_PATENT' });
    if (resp?.error) {
      showToast(resp.error, 'error');
    } else {
      await restoreState();
      if (oldIndex >= total) {
        showToast('已经是最后一个专利');
      } else {
        showToast(`已切换到第 ${resp.currentIndex + 1} 个专利`);
      }
    }
  } catch (e) {
    showToast('切换失败: ' + e.message, 'error');
  } finally {
    els.btnNext.disabled = false;
    els.btnNext.textContent = '⏭️ 下一个专利';
  }
}


// ========== 更新 UI ==========
function updateUI(state) {
  if (!state) return;
  const { patents, currentIndex } = state;

  if (!patents || patents.length === 0) return;

  // 隐藏初始模块，显示主操作区
  els.loadCard.classList.add('hidden');
  els.columnCard.classList.add('hidden');
  els.mainCard.classList.remove('hidden');
  els.listCard.classList.remove('hidden');

  // 当前专利号
  if (currentIndex < patents.length) {
    els.currentPatent.textContent = patents[currentIndex];
  } else {
    els.currentPatent.textContent = '全部完成';
  }

  // 进度
  const total = patents.length;
  const done = Math.min(currentIndex, total);
  const percent = total > 0 ? Math.round((done / total) * 100) : 0;
  els.progressLabel.textContent = `${done} / ${total}`;
  els.progressPercent.textContent = `${percent}%`;
  els.progressFill.style.width = `${percent}%`;

  // 专利列表
  renderPatentList(patents, currentIndex, state.accumulatedFees || {}, state.printedPatents || {});

  // 累积计数
  const accCount = state.accumulatedFees ? Object.keys(state.accumulatedFees).length : 0;
  const accText = accCount > 0 ? `已累积 ${accCount} 个专利` : '';
  els.accumulatedCount.textContent = accText;
  els.settingsAccumulatedCount.textContent = accText;
}

function renderPatentList(patents, currentIndex, accumulatedFees, printedPatents) {
  els.patentList.innerHTML = '';
  patents.forEach((patent, i) => {
    const item = document.createElement('div');
    item.className = 'patent-item';
    if (i === currentIndex) item.classList.add('active');
    if (i < currentIndex) item.classList.add('done');

    const seq = document.createElement('span');
    seq.className = 'seq';
    seq.textContent = String(i + 1).padStart(3, '0');

    const name = document.createElement('span');
    name.textContent = patent;

    item.appendChild(seq);
    item.appendChild(name);

    // 固定位置图标区：左打印，右累积
    const icons = document.createElement('span');
    icons.style.cssText = 'margin-left: auto; display: flex; gap: 2px; flex-shrink: 0; width: 32px;';

    const printSlot = document.createElement('span');
    printSlot.style.cssText = 'width: 15px; text-align: center; font-size: 11px;';
    if (printedPatents[patent]) {
      printSlot.textContent = '🖨️';
      printSlot.title = '已打印';
    }
    icons.appendChild(printSlot);

    const accSlot = document.createElement('span');
    accSlot.style.cssText = 'width: 15px; text-align: center; font-size: 11px;';
    if (accumulatedFees[patent]) {
      accSlot.textContent = '📦';
      accSlot.title = '已提取并累积';
    }
    icons.appendChild(accSlot);

    item.appendChild(icons);

    // 点击跳转到指定专利
    item.addEventListener('click', async () => {
      await chrome.runtime.sendMessage({ action: 'SET_INDEX', data: { index: i } });
    });

    els.patentList.appendChild(item);
  });

  // 滚动到当前项
  const activeItem = els.patentList.querySelector('.active');
  if (activeItem) {
    activeItem.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }
}

// ========== 恢复状态 ==========
async function restoreState() {
  try {
    const resp = await chrome.runtime.sendMessage({ action: 'GET_STATE' });
    if (resp?.success && resp.data.patents.length > 0) {
      updateUI(resp.data);
      if (resp.data.searchMode) {
        currentSearchMode = resp.data.searchMode;
        updateModeUI(resp.data.searchMode);
      }
      const count = resp.data.accumulatedFees ? Object.keys(resp.data.accumulatedFees).length : 0;
      if (count > 0) {
        const accText = `已累积 ${count} 个专利`;
        els.accumulatedCount.textContent = accText;
        els.settingsAccumulatedCount.textContent = accText;
      }
    }
  } catch (e) {
    // 静默
  }
}

// ========== 工具函数 ==========
function showToast(message, type = 'success', action = null) {
  const existing = document.querySelectorAll('.toast');
  const topOffset = 60 + existing.length * 48;
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.style.top = topOffset + 'px';

  const msgSpan = document.createElement('span');
  msgSpan.textContent = message;
  toast.appendChild(msgSpan);

  if (action) {
    const btn = document.createElement('button');
    btn.textContent = action.label;
    btn.style.cssText = 'margin-left: 10px; background: rgba(255,255,255,0.3); border: 1px solid rgba(255,255,255,0.6); color: white; padding: 2px 8px; border-radius: 4px; cursor: pointer; font-size: 12px;';
    btn.addEventListener('click', async () => {
      await action.callback();
      toast.remove();
      document.querySelectorAll('.toast').forEach((t, i) => {
        t.style.top = (60 + i * 48) + 'px';
      });
    });
    toast.appendChild(btn);
  }

  document.body.appendChild(toast);
  setTimeout(() => {
    toast.remove();
    document.querySelectorAll('.toast').forEach((t, i) => {
      t.style.top = (60 + i * 48) + 'px';
    });
  }, action ? 5000 : 3000);
}

// ========== 启动 ==========
init();
