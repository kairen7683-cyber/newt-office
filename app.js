/* ============================================================================
 * NEWT Web Office — Instant Native Excel · Word · PowerPoint Suite
 * 0ms Render Engine (No heavy external grid libraries)
 * ============================================================================ */

let currentMode = 'excel'; // 'excel' | 'word' | 'ppt'

// ==================== EXCEL STATE ====================
let excelSheets = [];
let activeSheetIdx = 0;
let selectedRow = 0;
let selectedCol = 0;
let editingCellInput = null;

// ==================== PPT STATE ====================
let pptDeck = [];
let currentSlideIndex = 0;
let selectedElementId = null;
let presentationIndex = 0;

// LocalStorage Keys
const STORAGE_KEYS = {
  TITLE: 'newt_office_doc_title_v2',
  EXCEL_DATA: 'newt_office_excel_v2',
  WORD_HTML: 'newt_office_word_v2',
  PPT_DECK: 'newt_office_ppt_v2'
};

/* ==================== INITIALIZATION ==================== */
window.addEventListener('DOMContentLoaded', () => {
  const savedTitle = localStorage.getItem(STORAGE_KEYS.TITLE);
  if (savedTitle) {
    document.getElementById('docTitleInput').value = savedTitle;
  }
  document.getElementById('docTitleInput').addEventListener('input', (e) => {
    localStorage.setItem(STORAGE_KEYS.TITLE, e.target.value.trim() || '제목 없는 문서');
    markAutoSaved();
  });

  document.getElementById('globalFileInput').addEventListener('change', handleGlobalFileOpen);

  // Initialize 3 Engines Instantly
  initExcelEngine();
  initWordEngine();
  initPptEngine();

  // Global keyboard handler
  window.addEventListener('keydown', handleGlobalKeyDown);
});

function showToast(msg) {
  const toast = document.getElementById('officeToast');
  toast.textContent = msg;
  toast.classList.remove('hidden');
  clearTimeout(window.__toastTimer);
  window.__toastTimer = setTimeout(() => {
    toast.classList.add('hidden');
  }, 2600);
}

function markAutoSaved() {
  const badge = document.getElementById('autoSaveBadge');
  const now = new Date();
  badge.textContent = `자동 저장 (${now.toTimeString().slice(0, 5)})`;
}

function getCleanDocTitle() {
  const val = document.getElementById('docTitleInput').value.trim();
  return val || 'NEWT_문서';
}

/* ==================== MODE SWITCHING ==================== */
function switchOfficeMode(mode) {
  currentMode = mode;

  document.querySelectorAll('.suite-tab').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.mode === mode);
  });

  document.querySelectorAll('.office-view').forEach((sec) => {
    sec.classList.remove('active');
  });
  document.getElementById(`view-${mode}`).classList.add('active');

  const badge = document.getElementById('brandBadge');
  const dlBtn = document.getElementById('primaryDownloadBtn');
  const dlLabel = document.getElementById('downloadBtnLabel');

  if (mode === 'excel') {
    badge.style.background = 'var(--excel-primary)';
    badge.textContent = 'X';
    dlBtn.style.background = 'var(--excel-primary)';
    dlLabel.textContent = '💾 .xlsx 다운로드';
    renderExcelGrid();
  } else if (mode === 'word') {
    badge.style.background = 'var(--word-primary)';
    badge.textContent = 'W';
    dlBtn.style.background = 'var(--word-primary)';
    dlLabel.textContent = '💾 .doc 다운로드';
  } else if (mode === 'ppt') {
    badge.style.background = 'var(--ppt-primary)';
    badge.textContent = 'P';
    dlBtn.style.background = 'var(--ppt-primary)';
    dlLabel.textContent = '💾 .pptx 다운로드';
    renderPptWorkspace();
  }
}

function downloadCurrentDocument() {
  if (currentMode === 'excel') exportExcelXLSX();
  else if (currentMode === 'word') exportWordDoc();
  else if (currentMode === 'ppt') exportPptxFile();
}

function createNewDocument() {
  if (currentMode === 'excel') {
    if (!confirm('현재 시트를 비우고 새 엑셀 문서를 만드시겠습니까?')) return;
    excelSheets = [createBlankExcelSheet('Sheet1', 35, 14)];
    activeSheetIdx = 0;
    selectedRow = 0;
    selectedCol = 0;
    saveExcelState();
    renderExcelWorkspace();
    showToast('새 빈 엑셀 시트가 생성되었습니다.');
  } else if (currentMode === 'word') {
    if (!confirm('현재 워드 문서를 비우고 새로 작성하시겠습니까?')) return;
    document.getElementById('wordPageEditor').innerHTML =
      '<h1>새 문서 제목</h1><p>여기에 내용을 입력하세요...</p>';
    saveWordState();
    updateWordCounters();
    showToast('새 워드 문서가 생성되었습니다.');
  } else if (currentMode === 'ppt') {
    if (!confirm('현재 슬라이드를 초기화하고 새 프레젠테이션을 만드시겠습니까?')) return;
    pptDeck = [createDefaultTitleSlide()];
    currentSlideIndex = 0;
    selectedElementId = null;
    savePptState();
    renderPptWorkspace();
    showToast('새 프레젠테이션이 생성되었습니다.');
  }
}

/* ============================================================================
 * 1. NATIVE INSTANT EXCEL ENGINE (0ms Grid, Formulas, Multi-Sheet, .xlsx/.csv)
 * ============================================================================ */
function colIdxToName(c) {
  let name = '';
  let n = c;
  while (n >= 0) {
    name = String.fromCharCode((n % 26) + 65) + name;
    n = Math.floor(n / 26) - 1;
  }
  return name;
}

function colNameToIdx(letters) {
  let col = 0;
  const up = letters.toUpperCase();
  for (let i = 0; i < up.length; i++) {
    col = col * 26 + (up.charCodeAt(i) - 64);
  }
  return col - 1;
}

function cellCoordToAddr(r, c) {
  return `${colIdxToName(c)}${r + 1}`;
}

function addrToCellCoord(addr) {
  const m = String(addr).trim().match(/^([A-Za-z]+)(\d+)$/);
  if (!m) return null;
  return {
    c: colNameToIdx(m[1]),
    r: parseInt(m[2], 10) - 1
  };
}

function createBlankExcelSheet(name = 'Sheet1', rows = 35, cols = 14) {
  return {
    name,
    rows,
    cols,
    cells: {} // key: "r,c" -> { raw: "", bold: false, align: "", color: "", bg: "", comma: true }
  };
}

function initExcelEngine() {
  const saved = localStorage.getItem(STORAGE_KEYS.EXCEL_DATA);
  if (saved) {
    try {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0 && parsed[0].cells) {
        excelSheets = parsed;
      }
    } catch (e) {}
  }

  if (!excelSheets || excelSheets.length === 0) {
    loadExcelSample('budget', true);
  } else {
    renderExcelWorkspace();
  }

  const formulaInput = document.getElementById('excelFormulaInput');
  formulaInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      setCellRawValue(selectedRow, selectedCol, formulaInput.value);
      selectExcelCell(Math.min(getActiveSheet().rows - 1, selectedRow + 1), selectedCol);
    }
  });
  formulaInput.addEventListener('input', () => {
    setCellRawValue(selectedRow, selectedCol, formulaInput.value, false);
  });
}

function getActiveSheet() {
  return excelSheets[activeSheetIdx] || excelSheets[0];
}

function saveExcelState() {
  localStorage.setItem(STORAGE_KEYS.EXCEL_DATA, JSON.stringify(excelSheets));
  markAutoSaved();
}

function setCellRawValue(r, c, rawVal, fullRender = true) {
  const sheet = getActiveSheet();
  const key = `${r},${c}`;
  if (!sheet.cells[key]) {
    sheet.cells[key] = { raw: '', comma: true };
  }
  sheet.cells[key].raw = String(rawVal ?? '');
  saveExcelState();
  if (fullRender) {
    renderExcelGrid();
    updateExcelFormulaBarAndStats();
  }
}

function getCellObj(sheet, r, c) {
  return sheet.cells[`${r},${c}`] || null;
}

/* ---- Excel Formula Evaluator (=SUM, =AVERAGE, =MIN, =MAX, =COUNT, =IF, arithmetic) ---- */
function evaluateCellValue(sheet, r, c, visited = new Set()) {
  const key = `${r},${c}`;
  if (visited.has(key)) return '#REF!';
  visited.add(key);

  const cell = sheet.cells[key];
  if (!cell || cell.raw === '' || cell.raw === null || cell.raw === undefined) {
    return '';
  }

  const rawStr = String(cell.raw).trim();
  if (!rawStr.startsWith('=')) {
    const num = Number(rawStr.replace(/,/g, ''));
    if (rawStr !== '' && !isNaN(num) && /^-?\d+(\.\d+)?$/.test(rawStr.replace(/,/g, ''))) {
      return num;
    }
    return rawStr;
  }

  try {
    let expr = rawStr.slice(1).trim();

    // 1. Expand Range Functions: SUM(A1:B5), AVERAGE(C2:C6), MIN(...), MAX(...), COUNT(...)
    expr = expr.replace(/\b(SUM|AVERAGE|AVG|MIN|MAX|COUNT)\s*\(\s*([A-Za-z]+\d+)\s*:\s*([A-Za-z]+\d+)\s*\)/gi, (_, fn, startAddr, endAddr) => {
      const s = addrToCellCoord(startAddr);
      const e = addrToCellCoord(endAddr);
      if (!s || !e) return '0';
      const r1 = Math.min(s.r, e.r);
      const r2 = Math.max(s.r, e.r);
      const c1 = Math.min(s.c, e.c);
      const c2 = Math.max(s.c, e.c);
      const nums = [];
      for (let rr = r1; rr <= r2; rr++) {
        for (let cc = c1; cc <= c2; cc++) {
          const val = evaluateCellValue(sheet, rr, cc, new Set(visited));
          const n = Number(val);
          if (val !== '' && !isNaN(n)) nums.push(n);
        }
      }
      const upFn = fn.toUpperCase();
      if (upFn === 'SUM') return String(nums.reduce((a, b) => a + b, 0));
      if (upFn === 'AVERAGE' || upFn === 'AVG') return nums.length ? String(nums.reduce((a, b) => a + b, 0) / nums.length) : '0';
      if (upFn === 'MIN') return nums.length ? String(Math.min(...nums)) : '0';
      if (upFn === 'MAX') return nums.length ? String(Math.max(...nums)) : '0';
      if (upFn === 'COUNT') return String(nums.length);
      return '0';
    });

    // 2. Replace Single Cell References (e.g. C2, D5)
    expr = expr.replace(/\b([A-Za-z]{1,2}\d{1,3})\b/g, (addr) => {
      const coord = addrToCellCoord(addr);
      if (!coord) return '0';
      const val = evaluateCellValue(sheet, coord.r, coord.c, new Set(visited));
      const n = Number(val);
      return val !== '' && !isNaN(n) ? String(n) : '0';
    });

    // Safe arithmetic evaluation
    if (!/^[0-9+\-*/().\s]+$/.test(expr)) {
      return '#수식오류';
    }
    const result = Function(`"use strict"; return (${expr});`)();
    if (typeof result === 'number' && isFinite(result)) {
      return Math.round(result * 100) / 100;
    }
    return result;
  } catch (err) {
    return '#ERROR';
  }
}

function formatDisplayValue(sheet, r, c) {
  const cell = getCellObj(sheet, r, c);
  const val = evaluateCellValue(sheet, r, c);
  if (val === '' || val === null || val === undefined) return '';
  if (typeof val === 'number') {
    if (cell && cell.comma === false) return String(val);
    return val.toLocaleString('ko-KR');
  }
  return String(val);
}

function renderExcelWorkspace() {
  renderExcelSheetTabs();
  renderExcelGrid();
  updateExcelFormulaBarAndStats();
}

function renderExcelSheetTabs() {
  const container = document.getElementById('excelSheetTabs');
  container.innerHTML = '';

  excelSheets.forEach((sh, idx) => {
    const btn = document.createElement('button');
    btn.className = `sheet-tab-btn ${idx === activeSheetIdx ? 'active' : ''}`;
    btn.textContent = sh.name;
    btn.onclick = () => {
      activeSheetIdx = idx;
      selectedRow = 0;
      selectedCol = 0;
      renderExcelWorkspace();
    };
    btn.ondblclick = () => {
      const newName = prompt('시트 이름을 입력하세요:', sh.name);
      if (newName && newName.trim()) {
        sh.name = newName.trim();
        saveExcelState();
        renderExcelSheetTabs();
      }
    };
    container.appendChild(btn);
  });
}

function addNewExcelSheet() {
  const newName = `Sheet${excelSheets.length + 1}`;
  excelSheets.push(createBlankExcelSheet(newName, 35, 14));
  activeSheetIdx = excelSheets.length - 1;
  selectedRow = 0;
  selectedCol = 0;
  saveExcelState();
  renderExcelWorkspace();
  showToast(`'${newName}' 시트가 추가되었습니다. (시트 탭 더블클릭 시 이름 변경)`);
}

function renderExcelGrid() {
  const sheet = getActiveSheet();
  const table = document.getElementById('excelGridTable');
  if (!sheet || !table) return;

  let html = '<thead><tr><th class="corner-header"></th>';
  for (let c = 0; c < sheet.cols; c++) {
    const activeCls = c === selectedCol ? 'active-header' : '';
    html += `<th class="${activeCls}" data-col="${c}">${colIdxToName(c)}</th>`;
  }
  html += '</tr></thead><tbody>';

  for (let r = 0; r < sheet.rows; r++) {
    const rowActiveCls = r === selectedRow ? 'active-header' : '';
    html += `<tr><th class="row-header ${rowActiveCls}" data-row="${r}">${r + 1}</th>`;

    for (let c = 0; c < sheet.cols; c++) {
      const cell = getCellObj(sheet, r, c);
      const disp = formatDisplayValue(sheet, r, c);
      const isSelected = r === selectedRow && c === selectedCol;
      const isNum = typeof evaluateCellValue(sheet, r, c) === 'number';

      let styleStr = '';
      if (cell) {
        if (cell.bg) styleStr += `background-color:${cell.bg};`;
        if (cell.color) styleStr += `color:${cell.color};`;
        if (cell.bold) styleStr += `font-weight:700;`;
        if (cell.align) {
          styleStr += `text-align:${cell.align};`;
        } else if (isNum) {
          styleStr += `text-align:right;`;
        }
      } else if (isNum) {
        styleStr += `text-align:right;`;
      }

      const hasFormula = cell && String(cell.raw).startsWith('=');
      const titleAttr = hasFormula ? `title="수식: ${escapeHtml(cell.raw)}"` : '';

      html += `<td class="excel-cell ${isSelected ? 'selected-cell' : ''}" data-r="${r}" data-c="${c}" style="${styleStr}" ${titleAttr}>${escapeHtml(disp)}</td>`;
    }
    html += '</tr>';
  }
  html += '</tbody>';

  table.innerHTML = html;

  // Bind click & dblclick on table via delegation
  table.onclick = (e) => {
    const td = e.target.closest('td.excel-cell');
    if (!td) return;
    const r = parseInt(td.dataset.r, 10);
    const c = parseInt(td.dataset.c, 10);
    if (editingCellInput && (r !== selectedRow || c !== selectedCol)) {
      commitInlineCellEdit();
    }
    selectExcelCell(r, c);
  };

  table.ondblclick = (e) => {
    const td = e.target.closest('td.excel-cell');
    if (!td) return;
    const r = parseInt(td.dataset.r, 10);
    const c = parseInt(td.dataset.c, 10);
    selectExcelCell(r, c);
    openInlineCellEditor(td, r, c);
  };
}

function selectExcelCell(r, c) {
  const sheet = getActiveSheet();
  selectedRow = Math.max(0, Math.min(sheet.rows - 1, r));
  selectedCol = Math.max(0, Math.min(sheet.cols - 1, c));

  // Fast DOM class update without rebuilding full table
  const table = document.getElementById('excelGridTable');
  table.querySelectorAll('.selected-cell').forEach((el) => el.classList.remove('selected-cell'));
  table.querySelectorAll('.active-header').forEach((el) => el.classList.remove('active-header'));

  const targetTd = table.querySelector(`td.excel-cell[data-r="${selectedRow}"][data-c="${selectedCol}"]`);
  if (targetTd) {
    targetTd.classList.add('selected-cell');
    targetTd.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  const colTh = table.querySelector(`thead th[data-col="${selectedCol}"]`);
  if (colTh) colTh.classList.add('active-header');
  const rowTh = table.querySelector(`tbody th[data-row="${selectedRow}"]`);
  if (rowTh) rowTh.classList.add('active-header');

  updateExcelFormulaBarAndStats();
}

function openInlineCellEditor(td, r, c, initialChar = null) {
  if (editingCellInput) commitInlineCellEdit();

  const sheet = getActiveSheet();
  const cell = getCellObj(sheet, r, c);
  const currentRaw = initialChar !== null ? initialChar : (cell ? cell.raw : '');

  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'excel-cell-editor';
  input.value = currentRaw;

  td.appendChild(input);
  input.focus();
  if (initialChar === null) {
    input.select();
  } else {
    input.setSelectionRange(input.value.length, input.value.length);
  }

  editingCellInput = { input, r, c };

  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      commitInlineCellEdit();
      selectExcelCell(r + 1, c);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      commitInlineCellEdit();
      selectExcelCell(r, c + (e.shiftKey ? -1 : 1));
    } else if (e.key === 'Escape') {
      e.preventDefault();
      editingCellInput = null;
      renderExcelGrid();
    }
  });

  input.addEventListener('blur', () => {
    if (editingCellInput) {
      commitInlineCellEdit();
    }
  });
}

function commitInlineCellEdit() {
  if (!editingCellInput) return;
  const { input, r, c } = editingCellInput;
  const val = input.value;
  editingCellInput = null;
  setCellRawValue(r, c, val, true);
}

function updateExcelFormulaBarAndStats() {
  const sheet = getActiveSheet();
  const addr = cellCoordToAddr(selectedRow, selectedCol);
  document.getElementById('excelCellAddressBox').textContent = addr;

  const cell = getCellObj(sheet, selectedRow, selectedCol);
  const formulaInput = document.getElementById('excelFormulaInput');
  if (document.activeElement !== formulaInput) {
    formulaInput.value = cell ? cell.raw : '';
  }

  // Compute sheet numeric summary
  let sum = 0;
  let count = 0;
  Object.keys(sheet.cells).forEach((k) => {
    const [r, c] = k.split(',').map(Number);
    const val = evaluateCellValue(sheet, r, c);
    if (typeof val === 'number' && !isNaN(val)) {
      sum += val;
      count++;
    }
  });
  const avg = count > 0 ? Math.round(sum / count) : 0;

  document.getElementById('excelSummaryStats').textContent =
    `선택 셀: ${addr}  |  시트 숫자 합계: ${sum.toLocaleString('ko-KR')}  |  평균: ${avg.toLocaleString('ko-KR')}  |  데이터 셀: ${count}개`;
}

function formatSelectedExcelCells(prop, value) {
  const sheet = getActiveSheet();
  const key = `${selectedRow},${selectedCol}`;
  if (!sheet.cells[key]) sheet.cells[key] = { raw: '', comma: true };
  const cell = sheet.cells[key];

  if (prop === 'bold') cell.bold = !cell.bold;
  else if (prop === 'align') cell.align = value;
  else if (prop === 'color') cell.color = value;
  else if (prop === 'bg') cell.bg = value;
  else if (prop === 'comma') cell.comma = cell.comma === false ? true : false;

  saveExcelState();
  renderExcelGrid();
}

function insertExcelFunction(fnName) {
  const sheet = getActiveSheet();
  // Find numeric cells above current cell in same column
  let startR = 1;
  const endR = Math.max(1, selectedRow);
  const colLetter = colIdxToName(selectedCol);
  const formula = `=${fnName}(${colLetter}${startR + 1}:${colLetter}${endR})`;
  setCellRawValue(selectedRow, selectedCol, formula, true);
  showToast(`${cellCoordToAddr(selectedRow, selectedCol)} 셀에 ${formula} 수식이 적용되었습니다.`);
}

function addExcelRows(count = 10) {
  const sheet = getActiveSheet();
  sheet.rows += count;
  saveExcelState();
  renderExcelGrid();
  showToast(`행 ${count}개가 추가되었습니다. (총 ${sheet.rows}행)`);
}

function addExcelCols(count = 5) {
  const sheet = getActiveSheet();
  sheet.cols += count;
  saveExcelState();
  renderExcelGrid();
  showToast(`열 ${count}개가 추가되었습니다. (총 ${sheet.cols}열)`);
}

function populateSheetFromAOA(sheetName, aoa, headerColor = '#107c41') {
  const rows = Math.max(35, aoa.length + 15);
  const cols = Math.max(12, (aoa[0]?.length || 8) + 4);
  const sheet = createBlankExcelSheet(sheetName, rows, cols);

  for (let r = 0; r < aoa.length; r++) {
    for (let c = 0; c < aoa[r].length; c++) {
      const val = aoa[r][c];
      if (val === undefined || val === null || val === '') continue;
      const isHeader = r === 0;
      const isTotalRow = r === aoa.length - 1 && String(aoa[r][0]).includes('합계');
      sheet.cells[`${r},${c}`] = {
        raw: String(val),
        bold: isHeader || isTotalRow,
        bg: isHeader ? headerColor : (isTotalRow ? '#ecfdf5' : ''),
        color: isHeader ? '#ffffff' : '#0f172a',
        align: isHeader ? 'center' : (typeof val === 'number' || String(val).startsWith('=') ? 'right' : 'left'),
        comma: true
      };
    }
  }
  return sheet;
}

function loadExcelSample(type, silent = false) {
  let aoa = [];
  let sheetName = 'Sheet1';
  let docTitle = '월간예산_가계부';

  if (type === 'budget') {
    sheetName = '월간예산_가계부';
    docTitle = '2026_월간예산_가계부';
    aoa = [
      ['항목 구분', '세부 지출 내역', '월 예산 (원)', '실제 지출 (원)', '잔여 예산 (수식)', '비고 / 결제수단'],
      ['고정지출', '통신비 및 인터넷 요금', 85000, 82000, '=C2-D2', '자동이체'],
      ['고정지출', 'MS 365 오피스 구독료 해지!', 11900, 0, '=C3-D3', 'NEWT 웹오피스로 100% 절감'],
      ['변동생활비', '식비 / 마트 장보기', 600000, 535000, '=C4-D4', '신용카드'],
      ['변동생활비', '교통비 / 차량 주유비', 200000, 185000, '=C5-D5', '체크카드'],
      ['문화/자기계발', '도서 구입 및 운동', 150000, 120000, '=C6-D6', ''],
      ['저축/투자', '미국 배당 ETF 정기 적립', 1000000, 1000000, '=C7-D7', '자동매수 완료'],
      ['합계', '월간 총계 (자동 수식 계산)', '=SUM(C2:C7)', '=SUM(D2:D7)', '=SUM(E2:E7)', '수식: =SUM(C2:C7)']
    ];
  } else if (type === 'invoice') {
    sheetName = '거래명세서';
    docTitle = '표준_거래명세서';
    aoa = [
      ['품목코드', '품목명 / 규격', '수량', '단가 (원)', '공급가액 (수식)', '부가세 10% (수식)', '합계금액 (수식)'],
      ['ITM-101', '비즈니스 노트북 16GB / 512GB', 2, 1250000, '=C2*D2', '=E2*0.1', '=E2+F2'],
      ['ITM-102', '27인치 4K UHD 모니터', 2, 420000, '=C3*D3', '=E3*0.1', '=E3+F3'],
      ['ITM-103', '무선 저소음 키보드/마우스 세트', 4, 65000, '=C4*D4', '=E4*0.1', '=E4+F4'],
      ['ITM-104', 'USB-C 멀티 도킹 스테이션', 2, 89000, '=C5*D5', '=E5*0.1', '=E5+F5'],
      ['합계', '총 청구 금액', '=SUM(C2:C5)', '', '=SUM(E2:E5)', '=SUM(F2:F5)', '=SUM(G2:G5)']
    ];
  } else if (type === 'calc') {
    sheetName = '복리수익_시뮬레이션';
    docTitle = '자산_복리수익_계산표';
    aoa = [
      ['투자 연차', '기초 투자금 (원)', '연간 추가납입 (원)', '연 목표수익률 (%)', '연말 평가금액 (수식)', '순수익금 (수식)'],
      [1, 20000000, 12000000, 12, '=(B2+C2)*1.12', '=E2-(B2+C2)'],
      [2, '=E2', 12000000, 12, '=(B3+C3)*1.12', '=E3-(B3+C3)'],
      [3, '=E3', 12000000, 12, '=(B4+C4)*1.12', '=E4-(B4+C4)'],
      [4, '=E4', 12000000, 12, '=(B5+C5)*1.12', '=E5-(B5+C5)'],
      [5, '=E5', 12000000, 12, '=(B6+C6)*1.12', '=E6-(B6+C6)'],
      ['합계', '5년 누적 성과', '=SUM(C2:C6)', '', '=MAX(E2:E6)', '=SUM(F2:F6)']
    ];
  }

  excelSheets = [populateSheetFromAOA(sheetName, aoa)];
  activeSheetIdx = 0;
  selectedRow = 1;
  selectedCol = 2;
  document.getElementById('docTitleInput').value = docTitle;
  saveExcelState();
  renderExcelWorkspace();
  if (!silent) showToast(`'${sheetName}' 시트가 즉시 로드되었습니다.`);
}

/* ---- Excel Keyboard Navigation ---- */
function handleGlobalKeyDown(e) {
  // Ignore if typing inside an input/textarea/word editor
  const tag = document.activeElement?.tagName;
  const isWordEditor = document.activeElement?.id === 'wordPageEditor';

  if (currentMode === 'excel' && !editingCellInput && tag !== 'INPUT' && tag !== 'TEXTAREA' && !isWordEditor) {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      selectExcelCell(selectedRow + 1, selectedCol);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      selectExcelCell(selectedRow - 1, selectedCol);
    } else if (e.key === 'ArrowRight' || e.key === 'Tab') {
      e.preventDefault();
      selectExcelCell(selectedRow, selectedCol + (e.shiftKey ? -1 : 1));
    } else if (e.key === 'ArrowLeft') {
      e.preventDefault();
      selectExcelCell(selectedRow, selectedCol - 1);
    } else if (e.key === 'Enter' || e.key === 'F2') {
      e.preventDefault();
      const td = document.querySelector(`td.excel-cell[data-r="${selectedRow}"][data-c="${selectedCol}"]`);
      if (td) openInlineCellEditor(td, selectedRow, selectedCol);
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      setCellRawValue(selectedRow, selectedCol, '', true);
    } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      const td = document.querySelector(`td.excel-cell[data-r="${selectedRow}"][data-c="${selectedCol}"]`);
      if (td) {
        e.preventDefault();
        openInlineCellEditor(td, selectedRow, selectedCol, e.key);
      }
    }
  } else if (currentMode === 'ppt') {
    if (e.key === 'F5') {
      e.preventDefault();
      startSlideShow();
    } else if (e.key === 'Escape') {
      const modal = document.getElementById('pptPresentationModal');
      if (!modal.classList.contains('hidden')) closeSlideShow();
    } else if (!document.getElementById('pptPresentationModal').classList.contains('hidden')) {
      if (e.key === 'ArrowRight' || e.key === ' ') stepPresentation(1);
      if (e.key === 'ArrowLeft') stepPresentation(-1);
    }
  }
}

/* ---- Excel Import & Export (.xlsx / .csv) ---- */
async function importExcelFile(file) {
  if (typeof XLSX === 'undefined') {
    showToast('엑셀 파서 라이브러리를 불러오는 중입니다. 1초 후 다시 시도해 주세요.');
    return;
  }
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array', cellFormula: true });

  excelSheets = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
    const rows = Math.max(35, aoa.length + 15);
    const cols = Math.max(14, (aoa[0]?.length || 10) + 5);
    const sheet = createBlankExcelSheet(name, rows, cols);

    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = ws[addr];
        if (!cell) continue;
        const rawVal = cell.f ? `=${cell.f}` : (cell.v !== undefined ? String(cell.v) : '');
        if (rawVal !== '') {
          sheet.cells[`${R},${C}`] = {
            raw: rawVal,
            bold: R === 0,
            bg: R === 0 ? '#f1f5f9' : '',
            comma: true
          };
        }
      }
    }
    return sheet;
  });

  activeSheetIdx = 0;
  selectedRow = 0;
  selectedCol = 0;
  saveExcelState();
  renderExcelWorkspace();
  showToast(`'${file.name}' (${excelSheets.length}개 시트) 불러오기 완료!`);
}

function buildWorkbookFromState() {
  const wb = XLSX.utils.book_new();

  excelSheets.forEach((sheet, idx) => {
    let maxR = 0;
    let maxC = 0;
    Object.keys(sheet.cells).forEach((k) => {
      const cell = sheet.cells[k];
      if (cell && cell.raw !== '') {
        const [r, c] = k.split(',').map(Number);
        if (r > maxR) maxR = r;
        if (c > maxC) maxC = c;
      }
    });

    const aoa = [];
    for (let r = 0; r <= maxR; r++) {
      const rowArr = [];
      for (let c = 0; c <= maxC; c++) {
        const val = evaluateCellValue(sheet, r, c);
        rowArr.push(val);
      }
      aoa.push(rowArr);
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Attach formulas to SheetJS worksheet
    for (let r = 0; r <= maxR; r++) {
      for (let c = 0; c <= maxC; c++) {
        const cell = getCellObj(sheet, r, c);
        if (cell && String(cell.raw).startsWith('=')) {
          const addr = XLSX.utils.encode_cell({ r, c });
          const evaluated = evaluateCellValue(sheet, r, c);
          ws[addr] = {
            t: typeof evaluated === 'number' ? 'n' : 's',
            v: evaluated,
            f: String(cell.raw).slice(1)
          };
        }
      }
    }

    XLSX.utils.book_append_sheet(wb, ws, sheet.name || `Sheet${idx + 1}`);
  });

  return wb;
}

function exportExcelXLSX() {
  if (typeof XLSX === 'undefined') {
    exportExcelCSV();
    return;
  }
  const wb = buildWorkbookFromState();
  const fileName = `${getCleanDocTitle()}.xlsx`;
  XLSX.writeFile(wb, fileName);
  showToast(`'${fileName}' 엑셀 파일 다운로드 완료!`);
}

function exportExcelCSV() {
  const sheet = getActiveSheet();
  let maxR = 0;
  let maxC = 0;
  Object.keys(sheet.cells).forEach((k) => {
    if (sheet.cells[k]?.raw !== '') {
      const [r, c] = k.split(',').map(Number);
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    }
  });

  const lines = [];
  for (let r = 0; r <= maxR; r++) {
    const rowCols = [];
    for (let c = 0; c <= maxC; c++) {
      const val = String(evaluateCellValue(sheet, r, c) ?? '').replace(/"/g, '""');
      rowCols.push(`"${val}"`);
    }
    lines.push(rowCols.join(','));
  }

  const blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
  triggerBlobDownload(blob, `${getCleanDocTitle()}.csv`);
  showToast('CSV 파일 다운로드 완료!');
}

/* ==================== GLOBAL FILE OPEN ==================== */
async function handleGlobalFileOpen(event) {
  const file = event.target.files[0];
  if (!file) return;
  event.target.value = '';

  const ext = file.name.split('.').pop().toLowerCase();
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  document.getElementById('docTitleInput').value = baseName;
  localStorage.setItem(STORAGE_KEYS.TITLE, baseName);

  try {
    if (['xlsx', 'xls', 'csv'].includes(ext)) {
      switchOfficeMode('excel');
      await importExcelFile(file);
    } else if (['docx'].includes(ext)) {
      switchOfficeMode('word');
      await importDocxFile(file);
    } else if (['doc', 'html', 'htm', 'txt'].includes(ext)) {
      switchOfficeMode('word');
      const text = await file.text();
      const editor = document.getElementById('wordPageEditor');
      if (ext === 'txt') {
        editor.innerHTML = text
          .split(/\r?\n/)
          .map((line) => `<p>${escapeHtml(line) || '&nbsp;'}</p>`)
          .join('');
      } else {
        editor.innerHTML = text;
      }
      saveWordState();
      updateWordCounters();
      showToast(`'${file.name}' 문서를 불러왔습니다.`);
    } else if (['pptx'].includes(ext)) {
      switchOfficeMode('ppt');
      await importPptxFile(file);
    }
  } catch (err) {
    console.error(err);
    showToast('파일 열기 오류: ' + err.message);
  }
}

/* ============================================================================
 * 2. WORD ENGINE (Rich A4 Editor + .docx Import + MS Word .doc Export)
 * ============================================================================ */
function initWordEngine() {
  const editor = document.getElementById('wordPageEditor');
  const savedHtml = localStorage.getItem(STORAGE_KEYS.WORD_HTML);

  if (savedHtml) {
    editor.innerHTML = savedHtml;
  } else {
    loadWordTemplate('report', true);
  }

  editor.addEventListener('input', () => {
    saveWordState();
    updateWordCounters();
  });

  updateWordCounters();
}

function saveWordState() {
  const editor = document.getElementById('wordPageEditor');
  localStorage.setItem(STORAGE_KEYS.WORD_HTML, editor.innerHTML);
  markAutoSaved();
}

function updateWordCounters() {
  const editor = document.getElementById('wordPageEditor');
  const text = editor.innerText || '';
  const totalChars = text.replace(/\n/g, '').length;
  const noSpaceChars = text.replace(/\s+/g, '').length;
  const words = text.trim() ? text.trim().split(/\s+/).length : 0;

  document.getElementById('wordCharCount').textContent =
    `글자 수: ${totalChars.toLocaleString()}자 (공백 제외 ${noSpaceChars.toLocaleString()}자) · 단어: ${words.toLocaleString()}개`;
}

function execWordCmd(cmd, value = null) {
  document.getElementById('wordPageEditor').focus();
  document.execCommand(cmd, false, value);
  saveWordState();
}

function applyWordFontSize(sizePt) {
  document.execCommand('fontSize', false, '7');
  const fontElements = document.getElementById('wordPageEditor').querySelectorAll('font[size="7"]');
  fontElements.forEach((el) => {
    el.removeAttribute('size');
    el.style.fontSize = sizePt;
  });
  saveWordState();
}

function insertWordTable() {
  const rows = parseInt(prompt('표의 행(줄) 수를 입력하세요:', '3'), 10) || 3;
  const cols = parseInt(prompt('표의 열(칸) 수를 입력하세요:', '3'), 10) || 3;

  let html = '<table><thead><tr>';
  for (let c = 0; c < cols; c++) html += `<th>항목 ${c + 1}</th>`;
  html += '</tr></thead><tbody>';
  for (let r = 0; r < rows - 1; r++) {
    html += '<tr>';
    for (let c = 0; c < cols; c++) html += '<td>내용 입력</td>';
    html += '</tr>';
  }
  html += '</tbody></table><p><br></p>';
  execWordCmd('insertHTML', html);
}

function insertWordImage() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/*';
  input.onchange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      execWordCmd('insertHTML', `<p><img src="${reader.result}" style="max-width:100%; height:auto;" /></p>`);
    };
    reader.readAsDataURL(file);
  };
  input.click();
}

function insertWordHR() {
  execWordCmd('insertHTML', '<hr style="border:none; border-top:1px solid #cbd5e1; margin:16px 0;" /><p><br></p>');
}

function loadWordTemplate(type, silent = false) {
  const editor = document.getElementById('wordPageEditor');
  const today = new Date().toISOString().slice(0, 10);

  if (type === 'report') {
    editor.innerHTML = `
      <h1>2026년 프로젝트 추진 기획 및 업무 보고서</h1>
      <p style="color:#475569;"><strong>작성일자:</strong> ${today} &nbsp;|&nbsp; <strong>작성자:</strong> 기획팀 &nbsp;|&nbsp; <strong>문서등급:</strong> 일반</p>
      <hr style="border:none; border-top:2px solid #1e3a8a; margin:14px 0 20px;" />

      <h2>1. 추진 배경 및 목적</h2>
      <p>상용 MS 오피스 구독 비용을 100% 절감하고, 어느 노트북에서나 웹 주소 접속 즉시 0.1초 만에 엑셀·워드·파워포인트 문서를 편집·저장할 수 있는 초경량 웹 오피스 환경을 구축합니다.</p>

      <h2>2. 주요 실행 계획 및 비용 비교</h2>
      <table>
        <thead>
          <tr>
            <th>구분</th>
            <th>기존 방식 (MS 365 구독)</th>
            <th>신규 도입 (NEWT 웹 오피스)</th>
            <th>절감 효과</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td><strong>연간 구독료</strong></td>
            <td>매년 정기 결제 발생</td>
            <td><strong>0원 (평생 무료)</strong></td>
            <td>100% 비용 절감</td>
          </tr>
          <tr>
            <td><strong>실행 속도</strong></td>
            <td>무거운 설치 프로그램 구동</td>
            <td><strong>브라우저 접속 즉시 0.1초 로딩</strong></td>
            <td>초고속 업무 처리</td>
          </tr>
          <tr>
            <td><strong>호환 파일</strong></td>
            <td>.xlsx, .docx, .pptx</td>
            <td><strong>.xlsx, .docx/.doc, .pptx 완벽 지원</strong></td>
            <td>100% 호환</td>
          </tr>
        </tbody>
      </table>
    `;
  } else if (type === 'minutes') {
    editor.innerHTML = `
      <h1>주간 실무 회의록</h1>
      <table>
        <tbody>
          <tr>
            <th style="width:20%;">회의 일시</th>
            <td style="width:30%;">${today} 10:00 ~ 11:00</td>
            <th style="width:20%;">회의 장소</th>
            <td style="width:30%;">대회의실</td>
          </tr>
          <tr>
            <th>참석자</th>
            <td colspan="3">홍길동 팀장, 김철수 선임, 이영희 매니저</td>
          </tr>
          <tr>
            <th>회의 안건</th>
            <td colspan="3"><strong>신규 사업 마일스톤 점검 및 리소스 배분</strong></td>
          </tr>
        </tbody>
      </table>
      <h2>1. 주요 결정 사항</h2>
      <ul>
        <li>핵심 과제 A: 금주 금요일까지 예산안 엑셀 시트 확정</li>
        <li>핵심 과제 B: 차주 화요일 고객사 제안 PPT 발표 진행</li>
      </ul>
    `;
  } else if (type === 'contract') {
    editor.innerHTML = `
      <h1 style="text-align:center;">업무 협약 및 표준 합의서</h1>
      <p style="text-align:center; color:#64748b;">문서번호: 제 2026-001 호</p>
      <p><br></p>
      <p><strong>갑 (발주자):</strong> ________________________ (인)</p>
      <p><strong>을 (수행자):</strong> ________________________ (인)</p>
      <h2>제 1 조 (목적)</h2>
      <p>본 합의서는 '갑'과 '을' 간의 프로젝트 수행 범위 및 상호 권리·의무를 규정함을 목적으로 한다.</p>
      <h2>제 2 조 (계약 기간 및 금액)</h2>
      <ul>
        <li><strong>계약 기간:</strong> ${today} ~ 2026-12-31</li>
        <li><strong>계약 금액:</strong> 일금 __________________ 원정 (₩ ____________)</li>
      </ul>
    `;
  }

  saveWordState();
  updateWordCounters();
  if (!silent) showToast('문서 양식을 불러왔습니다.');
}

async function importDocxFile(file) {
  if (typeof mammoth === 'undefined') {
    showToast('DOCX 파서를 불러오는 중입니다. 잠시 후 다시 시도해 주세요.');
    return;
  }
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer });
  document.getElementById('wordPageEditor').innerHTML = result.value || '<p>(빈 문서)</p>';
  saveWordState();
  updateWordCounters();
  showToast(`워드 파일 '${file.name}' 열기 완료!`);
}

function exportWordDoc() {
  const content = document.getElementById('wordPageEditor').innerHTML;
  const title = getCleanDocTitle();
  const fullWordHtml = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office'
          xmlns:w='urn:schemas-microsoft-com:office:word'
          xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
      <meta charset='utf-8'>
      <title>${escapeHtml(title)}</title>
      <style>
        @page { size: 21cm 29.7cm; margin: 2cm; }
        body { font-family: 'Malgun Gothic', '맑은 고딕', sans-serif; font-size: 11pt; line-height: 1.65; color: #1e293b; }
        table { width: 100%; border-collapse: collapse; margin: 12px 0; }
        th, td { border: 1px solid #64748b; padding: 6px 8px; }
        th { background-color: #f1f5f9; font-weight: bold; }
      </style>
    </head>
    <body>${content}</body>
    </html>
  `;
  const blob = new Blob(['\uFEFF', fullWordHtml], { type: 'application/msword;charset=utf-8' });
  triggerBlobDownload(blob, `${title}.doc`);
  showToast(`'${title}.doc' 다운로드 완료!`);
}

function printWordDocument() {
  window.print();
}

/* ============================================================================
 * 3. POWERPOINT ENGINE (16:9 Slide Editor + PptxGenJS .pptx Export & Import)
 * ============================================================================ */
function initPptEngine() {
  const saved = localStorage.getItem(STORAGE_KEYS.PPT_DECK);
  if (saved) {
    try {
      pptDeck = JSON.parse(saved);
    } catch (e) {
      pptDeck = null;
    }
  }
  if (!pptDeck || !Array.isArray(pptDeck) || pptDeck.length === 0) {
    loadPptTemplateDeck('pitch', true);
  } else {
    renderPptWorkspace();
  }
}

function createDefaultTitleSlide() {
  return {
    id: 'slide_' + Date.now(),
    bg: '#ffffff',
    notes: '',
    elements: [
      {
        id: 'el_' + Date.now() + '_1',
        type: 'text',
        x: 70,
        y: 160,
        w: 820,
        h: 90,
        text: '프레젠테이션 제목을 입력하세요',
        fontSize: 42,
        color: '#0f172a',
        bg: 'transparent',
        bold: true,
        align: 'left'
      },
      {
        id: 'el_' + Date.now() + '_2',
        type: 'text',
        x: 70,
        y: 265,
        w: 750,
        h: 60,
        text: '부제목 또는 발표자 정보를 여기에 작성합니다',
        fontSize: 20,
        color: '#475569',
        bg: 'transparent',
        bold: false,
        align: 'left'
      }
    ]
  };
}

function loadPptTemplateDeck(templateName, silent = false) {
  if (templateName === 'pitch') {
    pptDeck = [
      {
        id: 's1',
        bg: '#0f172a',
        notes: '표지 슬라이드',
        elements: [
          {
            id: 'e1',
            type: 'box',
            x: 70,
            y: 120,
            w: 160,
            h: 36,
            text: '2026 STRATEGY',
            fontSize: 13,
            color: '#ffffff',
            bg: '#c43e1c',
            bold: true,
            align: 'center'
          },
          {
            id: 'e2',
            type: 'text',
            x: 70,
            y: 180,
            w: 820,
            h: 110,
            text: 'NEWT Web Office 비즈니스 제안서',
            fontSize: 44,
            color: '#ffffff',
            bg: 'transparent',
            bold: true,
            align: 'left'
          },
          {
            id: 'e3',
            type: 'text',
            x: 70,
            y: 300,
            w: 760,
            h: 70,
            text: 'MS 오피스 구독 없이 모든 노트북에서 0.1초 만에 실행되는 웹 오피스',
            fontSize: 20,
            color: '#94a3b8',
            bg: 'transparent',
            bold: false,
            align: 'left'
          }
        ]
      },
      {
        id: 's2',
        bg: '#ffffff',
        notes: '핵심 장점 3가지',
        elements: [
          {
            id: 'e21',
            type: 'text',
            x: 60,
            y: 45,
            w: 840,
            h: 55,
            text: '1. 핵심 도입 효과 (Key Benefits)',
            fontSize: 30,
            color: '#0f172a',
            bg: 'transparent',
            bold: true,
            align: 'left'
          },
          {
            id: 'e22',
            type: 'box',
            x: 60,
            y: 140,
            w: 260,
            h: 300,
            text: '💰 구독료 0원\n\n매년 나가는 MS 365 정기 결제 비용을 100% 절감합니다.',
            fontSize: 18,
            color: '#1e293b',
            bg: '#f1f5f9',
            bold: false,
            align: 'left'
          },
          {
            id: 'e23',
            type: 'box',
            x: 350,
            y: 140,
            w: 260,
            h: 300,
            text: '⚡ 0.1초 즉시 로딩\n\n외부 무거운 설치 없이 웹 주소만 열면 바로 작동합니다.',
            fontSize: 18,
            color: '#1e293b',
            bg: '#eff6ff',
            bold: false,
            align: 'left'
          },
          {
            id: 'e24',
            type: 'box',
            x: 640,
            y: 140,
            w: 260,
            h: 300,
            text: '📂 완벽한 파일 호환\n\n.xlsx, .docx, .pptx 파일을 그대로 열고 저장합니다.',
            fontSize: 18,
            color: '#1e293b',
            bg: '#ecfdf5',
            bold: false,
            align: 'left'
          }
        ]
      },
      {
        id: 's3',
        bg: '#f8fafc',
        notes: '지원 기능 요약 표',
        elements: [
          {
            id: 'e31',
            type: 'text',
            x: 60,
            y: 45,
            w: 840,
            h: 55,
            text: '2. 기능 및 포맷 지원 현황',
            fontSize: 30,
            color: '#0f172a',
            bg: 'transparent',
            bold: true,
            align: 'left'
          },
          {
            id: 'e32',
            type: 'table',
            x: 60,
            y: 135,
            w: 840,
            h: 260,
            rows: [
              ['오피스 모듈', '지원 기능', '출력 파일 포맷'],
              ['엑셀 (Spreadsheet)', '수식 계산(SUM/AVERAGE), 다중 시트, 셀 서식', '표준 .xlsx / .csv'],
              ['워드 (Document)', 'A4 페이지 편집, 표/이미지 삽입, 양식 제공', 'MS 호환 .doc / PDF'],
              ['파워포인트 (Slide)', '16:9 슬라이드 편집, 발표 모드(F5)', '표준 PowerPoint .pptx']
            ],
            fontSize: 16,
            color: '#0f172a',
            bg: '#ffffff'
          }
        ]
      }
    ];
  } else if (templateName === 'weekly') {
    pptDeck = [
      {
        id: 'w1',
        bg: '#eff6ff',
        notes: '',
        elements: [
          {
            id: 'we1',
            type: 'text',
            x: 70,
            y: 190,
            w: 820,
            h: 90,
            text: '주간 업무 실적 및 차주 계획',
            fontSize: 42,
            color: '#1e3a8a',
            bg: 'transparent',
            bold: true,
            align: 'left'
          }
        ]
      },
      {
        id: 'w2',
        bg: '#ffffff',
        notes: '',
        elements: [
          {
            id: 'we21',
            type: 'text',
            x: 60,
            y: 45,
            w: 840,
            h: 55,
            text: '금주 주요 진행 실적',
            fontSize: 30,
            color: '#0f172a',
            bg: 'transparent',
            bold: true,
            align: 'left'
          },
          {
            id: 'we22',
            type: 'box',
            x: 60,
            y: 130,
            w: 840,
            h: 320,
            text: '• 핵심 과제 1: 초고속 네이티브 엑셀 스프레드시트 엔진 적용 완료\n\n• 핵심 과제 2: 신규 고객사 제안서 작성 완료',
            fontSize: 22,
            color: '#1e293b',
            bg: '#f8fafc',
            bold: false,
            align: 'left'
          }
        ]
      }
    ];
  }

  currentSlideIndex = 0;
  selectedElementId = null;
  savePptState();
  renderPptWorkspace();
  if (!silent) showToast('PPT 템플릿을 불러왔습니다.');
}

function savePptState() {
  localStorage.setItem(STORAGE_KEYS.PPT_DECK, JSON.stringify(pptDeck));
  markAutoSaved();
}

function renderPptWorkspace() {
  renderPptThumbnails();
  renderCurrentSlideStage();
  renderPptInspector();
}

function renderPptThumbnails() {
  const listEl = document.getElementById('pptThumbnailList');
  if (!listEl) return;
  listEl.innerHTML = '';

  pptDeck.forEach((slide, idx) => {
    const card = document.createElement('div');
    card.className = `slide-thumb-card ${idx === currentSlideIndex ? 'active' : ''}`;
    card.onclick = () => {
      currentSlideIndex = idx;
      selectedElementId = null;
      renderPptWorkspace();
    };

    const firstText = slide.elements.find((e) => e.text)?.text || `슬라이드 ${idx + 1}`;
    const isDark = slide.bg === '#0f172a';

    card.innerHTML = `
      <div class="thumb-header">
        <span>SLIDE ${idx + 1}</span>
        <span>${slide.elements.length}개 요소</span>
      </div>
      <div class="thumb-preview" style="background:${slide.bg}; color:${isDark ? '#fff' : '#0f172a'};">
        <div style="font-weight:700; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
          ${escapeHtml(firstText.split('\n')[0])}
        </div>
      </div>
    `;
    listEl.appendChild(card);
  });
}

function renderCurrentSlideStage() {
  const stage = document.getElementById('pptSlideStage');
  const slide = pptDeck[currentSlideIndex];
  if (!stage || !slide) return;

  stage.style.background = slide.bg || '#ffffff';
  stage.innerHTML = '';

  const bgSelect = document.getElementById('pptBgThemeSelect');
  if (bgSelect) bgSelect.value = slide.bg || '#ffffff';

  const notesArea = document.getElementById('pptSpeakerNotes');
  if (notesArea) notesArea.value = slide.notes || '';

  slide.elements.forEach((el) => {
    const item = document.createElement('div');
    item.className = `ppt-item ${el.id === selectedElementId ? 'selected' : ''}`;
    item.style.left = `${el.x}px`;
    item.style.top = `${el.y}px`;
    item.style.width = `${el.w}px`;
    item.style.height = `${el.h}px`;
    item.style.background = el.bg || 'transparent';
    item.style.color = el.color || '#0f172a';
    item.style.fontSize = `${el.fontSize || 20}px`;
    item.style.fontWeight = el.bold ? '700' : '400';
    item.style.textAlign = el.align || 'left';
    if (el.type === 'box') {
      item.style.borderRadius = '10px';
      item.style.padding = '16px';
    }

    item.addEventListener('mousedown', (e) => {
      e.stopPropagation();
      if (selectedElementId !== el.id) {
        selectedElementId = el.id;
        renderCurrentSlideStage();
        renderPptInspector();
      }
      startDraggingElement(e, el);
    });

    if (el.type === 'text' || el.type === 'box') {
      const editable = document.createElement('div');
      editable.contentEditable = 'true';
      editable.style.whiteSpace = 'pre-wrap';
      editable.innerText = el.text || '';
      editable.addEventListener('mousedown', (e) => {
        if (document.activeElement === editable) e.stopPropagation();
      });
      editable.addEventListener('input', () => {
        el.text = editable.innerText;
        savePptState();
        renderPptThumbnails();
      });
      item.appendChild(editable);
    } else if (el.type === 'image') {
      const img = document.createElement('img');
      img.src = el.src;
      img.style.width = '100%';
      img.style.height = '100%';
      img.style.objectFit = 'contain';
      img.draggable = false;
      item.appendChild(img);
    } else if (el.type === 'table') {
      const tbl = document.createElement('table');
      (el.rows || []).forEach((r, rIdx) => {
        const tr = document.createElement('tr');
        r.forEach((cellVal, cIdx) => {
          const cellEl = document.createElement(rIdx === 0 ? 'th' : 'td');
          cellEl.contentEditable = 'true';
          cellEl.innerText = cellVal;
          cellEl.addEventListener('input', () => {
            el.rows[rIdx][cIdx] = cellEl.innerText;
            savePptState();
          });
          tr.appendChild(cellEl);
        });
        tbl.appendChild(tr);
      });
      item.appendChild(tbl);
    }

    stage.appendChild(item);
  });
}

function startDraggingElement(e, el) {
  if (e.target.isContentEditable && document.activeElement === e.target) return;
  const startX = e.clientX;
  const startY = e.clientY;
  const origX = el.x;
  const origY = el.y;

  function onMove(moveEvt) {
    el.x = Math.max(0, Math.min(920, origX + (moveEvt.clientX - startX)));
    el.y = Math.max(0, Math.min(500, origY + (moveEvt.clientY - startY)));
    renderCurrentSlideStage();
  }
  function onUp() {
    window.removeEventListener('mousemove', onMove);
    window.removeEventListener('mouseup', onUp);
    savePptState();
    renderPptInspector();
  }
  window.addEventListener('mousemove', onMove);
  window.addEventListener('mouseup', onUp);
}

function deselectPptElement(e) {
  if (e.target.id === 'pptSlideStage' || e.target.classList.contains('ppt-stage-wrapper')) {
    selectedElementId = null;
    renderCurrentSlideStage();
    renderPptInspector();
  }
}

function renderPptInspector() {
  const body = document.getElementById('pptInspectorBody');
  if (!body) return;
  const slide = pptDeck[currentSlideIndex];
  const el = slide?.elements.find((item) => item.id === selectedElementId);

  if (!el) {
    body.innerHTML = `<p class="inspector-empty">슬라이드의 텍스트나 도형을 클릭하면 크기·색상·정렬을 조절할 수 있습니다.</p>`;
    return;
  }

  body.innerHTML = `
    <div class="prop-grid-2">
      <div class="prop-row">
        <label>가로 위치 (X)</label>
        <input type="number" value="${Math.round(el.x)}" onchange="updateSelectedPptProp('x', Number(this.value))" />
      </div>
      <div class="prop-row">
        <label>세로 위치 (Y)</label>
        <input type="number" value="${Math.round(el.y)}" onchange="updateSelectedPptProp('y', Number(this.value))" />
      </div>
      <div class="prop-row">
        <label>너비 (W)</label>
        <input type="number" value="${Math.round(el.w)}" onchange="updateSelectedPptProp('w', Number(this.value))" />
      </div>
      <div class="prop-row">
        <label>높이 (H)</label>
        <input type="number" value="${Math.round(el.h)}" onchange="updateSelectedPptProp('h', Number(this.value))" />
      </div>
    </div>
    <div class="prop-grid-2">
      <div class="prop-row">
        <label>글자 크기 (px)</label>
        <input type="number" value="${el.fontSize || 20}" onchange="updateSelectedPptProp('fontSize', Number(this.value))" />
      </div>
      <div class="prop-row">
        <label>정렬</label>
        <select onchange="updateSelectedPptProp('align', this.value)">
          <option value="left" ${el.align === 'left' ? 'selected' : ''}>왼쪽</option>
          <option value="center" ${el.align === 'center' ? 'selected' : ''}>가운데</option>
          <option value="right" ${el.align === 'right' ? 'selected' : ''}>오른쪽</option>
        </select>
      </div>
    </div>
    <div class="prop-grid-2">
      <div class="prop-row">
        <label>글자 색상</label>
        <input type="color" value="${el.color || '#0f172a'}" onchange="updateSelectedPptProp('color', this.value)" style="height:32px;" />
      </div>
      <div class="prop-row">
        <label>배경 색상</label>
        <input type="color" value="${el.bg && el.bg !== 'transparent' ? el.bg : '#ffffff'}" onchange="updateSelectedPptProp('bg', this.value)" style="height:32px;" />
      </div>
    </div>
    <div style="display:flex; gap:8px; margin-top:8px;">
      <button class="tool-btn" style="flex:1;" onclick="updateSelectedPptProp('bold', !${Boolean(el.bold)})"><b>B</b> 굵게</button>
      <button class="tool-btn danger-hover" style="flex:1;" onclick="deleteSelectedPptElement()">🗑️ 삭제</button>
    </div>
  `;
}

function updateSelectedPptProp(key, val) {
  const slide = pptDeck[currentSlideIndex];
  const el = slide?.elements.find((item) => item.id === selectedElementId);
  if (!el) return;
  el[key] = val;
  savePptState();
  renderCurrentSlideStage();
  renderPptInspector();
}

function deleteSelectedPptElement() {
  const slide = pptDeck[currentSlideIndex];
  if (!slide || !selectedElementId) return;
  slide.elements = slide.elements.filter((e) => e.id !== selectedElementId);
  selectedElementId = null;
  savePptState();
  renderPptWorkspace();
}

function addPptSlide() {
  pptDeck.splice(currentSlideIndex + 1, 0, {
    id: 'slide_' + Date.now(),
    bg: '#ffffff',
    notes: '',
    elements: [
      {
        id: 'el_' + Date.now() + '_1',
        type: 'text',
        x: 60,
        y: 45,
        w: 840,
        h: 60,
        text: `슬라이드 ${pptDeck.length + 1} 제목`,
        fontSize: 32,
        color: '#0f172a',
        bg: 'transparent',
        bold: true,
        align: 'left'
      },
      {
        id: 'el_' + Date.now() + '_2',
        type: 'box',
        x: 60,
        y: 130,
        w: 840,
        h: 320,
        text: '• 핵심 내용을 여기에 작성하세요.',
        fontSize: 20,
        color: '#334155',
        bg: '#f8fafc',
        bold: false,
        align: 'left'
      }
    ]
  });
  currentSlideIndex++;
  selectedElementId = null;
  savePptState();
  renderPptWorkspace();
}

function duplicatePptSlide() {
  const current = pptDeck[currentSlideIndex];
  if (!current) return;
  const clone = JSON.parse(JSON.stringify(current));
  clone.id = 'slide_' + Date.now();
  pptDeck.splice(currentSlideIndex + 1, 0, clone);
  currentSlideIndex++;
  savePptState();
  renderPptWorkspace();
}

function deletePptSlide() {
  if (pptDeck.length <= 1) {
    showToast('최소 1장의 슬라이드는 남아있어야 합니다.');
    return;
  }
  pptDeck.splice(currentSlideIndex, 1);
  currentSlideIndex = Math.max(0, currentSlideIndex - 1);
  selectedElementId = null;
  savePptState();
  renderPptWorkspace();
}

function addPptElement(type) {
  const slide = pptDeck[currentSlideIndex];
  if (!slide) return;
  const newId = 'el_' + Date.now();
  if (type === 'text') {
    slide.elements.push({
      id: newId,
      type: 'text',
      x: 120,
      y: 150,
      w: 500,
      h: 70,
      text: '새 텍스트 상자',
      fontSize: 24,
      color: slide.bg === '#0f172a' ? '#ffffff' : '#0f172a',
      bg: 'transparent',
      bold: false,
      align: 'left'
    });
  } else if (type === 'box') {
    slide.elements.push({
      id: newId,
      type: 'box',
      x: 150,
      y: 160,
      w: 320,
      h: 180,
      text: '강조 카드 / 도형 블록',
      fontSize: 20,
      color: '#0f172a',
      bg: '#eff6ff',
      bold: true,
      align: 'center'
    });
  } else if (type === 'table') {
    slide.elements.push({
      id: newId,
      type: 'table',
      x: 100,
      y: 160,
      w: 720,
      h: 200,
      rows: [
        ['구분', '주요 항목', '비고'],
        ['항목 1', '세부 내용을 입력하세요', '완료'],
        ['항목 2', '세부 내용을 입력하세요', '진행중']
      ],
      fontSize: 16,
      color: '#0f172a',
      bg: '#ffffff'
    });
  }
  selectedElementId = newId;
  savePptState();
  renderPptWorkspace();
}

function triggerPptImageUpload() {
  document.getElementById('pptImgInput').click();
}

function handlePptImageFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  event.target.value = '';
  const reader = new FileReader();
  reader.onload = () => {
    const slide = pptDeck[currentSlideIndex];
    const newId = 'el_' + Date.now();
    slide.elements.push({
      id: newId,
      type: 'image',
      x: 200,
      y: 120,
      w: 420,
      h: 280,
      src: reader.result
    });
    selectedElementId = newId;
    savePptState();
    renderPptWorkspace();
  };
  reader.readAsDataURL(file);
}

function changeSlideBackground(bgHex) {
  const slide = pptDeck[currentSlideIndex];
  if (!slide) return;
  slide.bg = bgHex;
  savePptState();
  renderPptWorkspace();
}

function saveCurrentSpeakerNotes(val) {
  const slide = pptDeck[currentSlideIndex];
  if (!slide) return;
  slide.notes = val;
  savePptState();
}

function startSlideShow() {
  presentationIndex = currentSlideIndex;
  const modal = document.getElementById('pptPresentationModal');
  modal.classList.remove('hidden');
  if (modal.requestFullscreen) modal.requestFullscreen().catch(() => {});
  renderPresentationSlide();
}

function closeSlideShow() {
  const modal = document.getElementById('pptPresentationModal');
  modal.classList.add('hidden');
  if (document.fullscreenElement && document.exitFullscreen) {
    document.exitFullscreen().catch(() => {});
  }
}

function stepPresentation(delta) {
  const next = presentationIndex + delta;
  if (next < 0 || next >= pptDeck.length) return;
  presentationIndex = next;
  renderPresentationSlide();
}

function renderPresentationSlide() {
  const host = document.getElementById('presentationStageHost');
  document.getElementById('presentationPageIndicator').textContent =
    `${presentationIndex + 1} / ${pptDeck.length}`;

  const slide = pptDeck[presentationIndex];
  const scale = Math.min(window.innerWidth / 960, window.innerHeight / 540) * 0.94;

  const stageClone = document.createElement('div');
  stageClone.className = 'ppt-slide-stage';
  stageClone.style.background = slide.bg || '#ffffff';
  stageClone.style.transform = `scale(${scale})`;

  slide.elements.forEach((el) => {
    const item = document.createElement('div');
    item.className = 'ppt-item';
    item.style.left = `${el.x}px`;
    item.style.top = `${el.y}px`;
    item.style.width = `${el.w}px`;
    item.style.height = `${el.h}px`;
    item.style.background = el.bg || 'transparent';
    item.style.color = el.color || '#0f172a';
    item.style.fontSize = `${el.fontSize || 20}px`;
    item.style.fontWeight = el.bold ? '700' : '400';
    item.style.textAlign = el.align || 'left';
    if (el.type === 'box') {
      item.style.borderRadius = '10px';
      item.style.padding = '16px';
    }
    if (el.type === 'text' || el.type === 'box') {
      item.style.whiteSpace = 'pre-wrap';
      item.innerText = el.text || '';
    } else if (el.type === 'image') {
      item.innerHTML = `<img src="${el.src}" style="width:100%;height:100%;object-fit:contain;" />`;
    } else if (el.type === 'table') {
      let tblHtml = '<table>';
      (el.rows || []).forEach((r, rIdx) => {
        tblHtml += '<tr>';
        r.forEach((c) => {
          tblHtml += rIdx === 0 ? `<th>${escapeHtml(c)}</th>` : `<td>${escapeHtml(c)}</td>`;
        });
        tblHtml += '</tr>';
      });
      tblHtml += '</table>';
      item.innerHTML = tblHtml;
    }
    stageClone.appendChild(item);
  });

  host.innerHTML = '';
  host.appendChild(stageClone);
}

function exportPptxFile() {
  if (typeof PptxGenJS === 'undefined') {
    showToast('PPTX 엔진을 준비 중입니다. 1초 후 다시 눌러주세요.');
    return;
  }
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  const pxToIn = (px) => Number((px / 96).toFixed(3));

  pptDeck.forEach((s) => {
    const slide = pptx.addSlide();
    if (s.bg && s.bg !== 'transparent') {
      slide.background = { color: s.bg.replace('#', '') };
    }
    if (s.notes) slide.addNotes(s.notes);

    s.elements.forEach((el) => {
      const x = pxToIn(el.x);
      const y = pxToIn(el.y);
      const w = pxToIn(el.w);
      const h = pxToIn(el.h);

      if (el.type === 'text' || el.type === 'box') {
        const opts = {
          x,
          y,
          w,
          h,
          fontSize: Math.round((el.fontSize || 20) * 0.75),
          color: (el.color || '#0f172a').replace('#', ''),
          bold: Boolean(el.bold),
          align: el.align || 'left',
          valign: 'top',
          fontFace: 'Malgun Gothic'
        };
        if (el.bg && el.bg !== 'transparent') {
          opts.fill = { color: el.bg.replace('#', '') };
        }
        slide.addText(el.text || '', opts);
      } else if (el.type === 'image' && el.src) {
        slide.addImage({ data: el.src, x, y, w, h });
      } else if (el.type === 'table' && Array.isArray(el.rows)) {
        slide.addTable(el.rows, {
          x,
          y,
          w,
          fontSize: 13,
          border: { pt: 1, color: '94a3b8' },
          fill: { color: 'FFFFFF' },
          fontFace: 'Malgun Gothic'
        });
      }
    });
  });

  const fileName = `${getCleanDocTitle()}.pptx`;
  pptx.writeFile({ fileName }).then(() => {
    showToast(`'${fileName}' 다운로드 완료!`);
  });
}

async function importPptxFile(file) {
  if (typeof JSZip === 'undefined') {
    showToast('PPTX 리더를 불러오는 중입니다. 1초 후 다시 시도해 주세요.');
    return;
  }
  const zip = await JSZip.loadAsync(file);
  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => parseInt(a.match(/slide(\d+)\.xml/)[1], 10) - parseInt(b.match(/slide(\d+)\.xml/)[1], 10));

  if (slideFiles.length === 0) {
    showToast('PPTX 내부에서 슬라이드를 찾지 못했습니다.');
    return;
  }

  const importedDeck = [];
  for (let i = 0; i < slideFiles.length; i++) {
    const xmlStr = await zip.files[slideFiles[i]].async('string');
    const xmlDoc = new DOMParser().parseFromString(xmlStr, 'application/xml');
    const shapes = Array.from(xmlDoc.getElementsByTagName('p:sp'));
    const elements = [];
    let autoY = 60;

    shapes.forEach((sp, spIdx) => {
      const lines = Array.from(sp.getElementsByTagName('a:p'))
        .map((p) => Array.from(p.getElementsByTagName('a:t')).map((t) => t.textContent).join(''))
        .filter((line) => line.trim().length > 0);
      if (lines.length === 0) return;

      const isTitle = spIdx === 0;
      elements.push({
        id: `imp_${i}_${spIdx}`,
        type: isTitle ? 'text' : 'box',
        x: 60,
        y: autoY,
        w: 840,
        h: isTitle ? 70 : Math.min(300, Math.max(80, lines.length * 42)),
        text: lines.join('\n'),
        fontSize: isTitle ? 32 : 20,
        color: '#0f172a',
        bg: isTitle ? 'transparent' : '#f8fafc',
        bold: isTitle,
        align: 'left'
      });
      autoY += isTitle ? 90 : Math.min(310, Math.max(95, lines.length * 45));
    });

    importedDeck.push({
      id: `imp_slide_${i}`,
      bg: '#ffffff',
      notes: '',
      elements
    });
  }

  pptDeck = importedDeck;
  currentSlideIndex = 0;
  selectedElementId = null;
  savePptState();
  renderPptWorkspace();
  showToast(`PPTX '${file.name}' (${importedDeck.length}장) 열기 완료!`);
}

/* ==================== HELPERS ==================== */
function triggerBlobDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 200);
}

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
