/* ============================================================================
 * Microsoft 365 Authentic Web Office Suite — Excel · Word · PowerPoint
 * Instant Native 0ms Render Engine + Ribbon Commands + Status Bar
 * ============================================================================ */

let currentMode = 'excel'; // 'excel' | 'word' | 'ppt'
let activeRibbonTab = 'home'; // 'home' | 'insert' | 'layout' | 'special' | 'templates'
let workspaceZoom = 100;

// ==================== EXCEL STATE ====================
let excelSheets = [];
let activeSheetIdx = 0;
let selectedRow = 0;
let selectedCol = 0;
let editingCellInput = null;
let excelClipboard = null;
let showExcelGridlines = true;

// ==================== PPT STATE ====================
let pptDeck = [];
let currentSlideIndex = 0;
let selectedElementId = null;
let presentationIndex = 0;

// LocalStorage Keys
const STORAGE_KEYS = {
  TITLE: 'newt_office_doc_title_v3',
  EXCEL_DATA: 'newt_office_excel_v3',
  WORD_HTML: 'newt_office_word_v3',
  PPT_DECK: 'newt_office_ppt_v3'
};

/* ==================== INITIALIZATION ==================== */
window.addEventListener('DOMContentLoaded', () => {
  const savedTitle = localStorage.getItem(STORAGE_KEYS.TITLE);
  if (savedTitle) {
    document.getElementById('docTitleInput').value = savedTitle;
  }
  document.getElementById('docTitleInput').addEventListener('input', (e) => {
    localStorage.setItem(STORAGE_KEYS.TITLE, e.target.value.trim() || '문서1');
    markAutoSaved();
  });

  document.getElementById('globalFileInput').addEventListener('change', handleGlobalFileOpen);

  // Close waffle menu when clicking outside
  window.addEventListener('click', (e) => {
    const dd = document.getElementById('waffleDropdown');
    if (dd && !dd.classList.contains('hidden')) {
      if (!e.target.closest('#waffleDropdown') && !e.target.closest('.waffle-btn')) {
        dd.classList.add('hidden');
      }
    }
  });

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
  }, 2500);
}

function markAutoSaved() {
  const badge = document.getElementById('autoSaveBadge');
  if (!badge) return;
  const now = new Date();
  badge.textContent = `· 저장됨 (${now.toTimeString().slice(0, 5)})`;
}

function getCleanDocTitle() {
  const val = document.getElementById('docTitleInput').value.trim();
  return val || '문서1';
}

function triggerGlobalFileOpen() {
  document.getElementById('globalFileInput').click();
}

function toggleWaffleMenu(e) {
  e.stopPropagation();
  document.getElementById('waffleDropdown').classList.toggle('hidden');
}

function closeWaffleMenu() {
  document.getElementById('waffleDropdown').classList.add('hidden');
}

function openFileBackstage() {
  document.getElementById('fileBackstageModal').classList.remove('hidden');
}

function closeFileBackstage() {
  document.getElementById('fileBackstageModal').classList.add('hidden');
}

function handleUndo() {
  if (currentMode === 'word') {
    execWordCmd('undo');
  } else {
    document.execCommand('undo');
  }
}

function handleRedo() {
  if (currentMode === 'word') {
    execWordCmd('redo');
  } else {
    document.execCommand('redo');
  }
}

/* ==================== RIBBON & APP MODE SWITCHING ==================== */
function switchOfficeMode(mode) {
  currentMode = mode;

  // 1. Body Theme Class (MS Excel Green / MS Word Blue / MS PPT Orange)
  document.body.className = `theme-${mode}`;

  // 2. Titlebar App Brand & Chip Active State
  const brandEl = document.getElementById('msAppBrandName');
  const extEl = document.getElementById('docExtBadge');
  const dlLabel = document.getElementById('downloadBtnLabel');
  const layoutTabLabel = document.getElementById('layoutTabLabel');
  const specialTabLabel = document.getElementById('specialTabLabel');

  document.querySelectorAll('.app-switch-chip').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.mode === mode);
  });

  // 3. Switch Active App Ribbon Group
  document.querySelectorAll('.app-ribbon-group').forEach((grp) => {
    grp.classList.toggle('active', grp.dataset.appRibbon === mode);
  });

  // 4. Switch Workspace View
  document.querySelectorAll('.office-view').forEach((sec) => {
    sec.classList.remove('active');
  });
  document.getElementById(`view-${mode}`).classList.add('active');

  if (mode === 'excel') {
    brandEl.textContent = 'Excel';
    extEl.textContent = '.xlsx';
    dlLabel.textContent = '⬇ .xlsx 저장';
    layoutTabLabel.textContent = '페이지 레이아웃';
    specialTabLabel.textContent = '수식 및 데이터';
    renderExcelWorkspace();
  } else if (mode === 'word') {
    brandEl.textContent = 'Word';
    extEl.textContent = '.doc';
    dlLabel.textContent = '⬇ .doc 저장';
    layoutTabLabel.textContent = '레이아웃';
    specialTabLabel.textContent = '검토 및 인쇄';
    updateWordCounters();
  } else if (mode === 'ppt') {
    brandEl.textContent = 'PowerPoint';
    extEl.textContent = '.pptx';
    dlLabel.textContent = '⬇ .pptx 저장';
    layoutTabLabel.textContent = '디자인 테마';
    specialTabLabel.textContent = '슬라이드 쇼';
    renderPptWorkspace();
  }

  // Ensure current ribbon tab panel is shown
  switchRibbonTab(activeRibbonTab);
}

function switchRibbonTab(tabName) {
  activeRibbonTab = tabName;

  document.querySelectorAll('.ribbon-tab').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.rtab === tabName);
  });

  // Shared templates panel vs app-specific panel
  const sharedPanel = document.querySelector('.shared-templates-panel');
  const activeAppGroup = document.querySelector(`.app-ribbon-group[data-app-ribbon="${currentMode}"]`);

  if (tabName === 'templates') {
    document.querySelectorAll('.app-ribbon-group').forEach((g) => g.classList.remove('active'));
    if (sharedPanel) sharedPanel.classList.add('active');
  } else {
    if (sharedPanel) sharedPanel.classList.remove('active');
    document.querySelectorAll('.app-ribbon-group').forEach((g) => {
      g.classList.toggle('active', g.dataset.appRibbon === currentMode);
    });
    if (activeAppGroup) {
      activeAppGroup.querySelectorAll('.ribbon-panel').forEach((p) => {
        p.classList.toggle('active', p.dataset.rpanel === tabName);
      });
    }
  }
}

/* ==================== WORKSPACE ZOOM SLIDER ==================== */
function setWorkspaceZoom(val) {
  workspaceZoom = Math.max(60, Math.min(150, Number(val) || 100));
  document.getElementById('workspaceZoomSlider').value = workspaceZoom;
  document.getElementById('workspaceZoomLabel').textContent = `${workspaceZoom}%`;

  const scale = workspaceZoom / 100;
  const excelTable = document.getElementById('excelGridTable');
  const wordPage = document.getElementById('wordPageEditor');
  const pptStage = document.getElementById('pptSlideStage');

  if (excelTable) {
    excelTable.style.zoom = `${workspaceZoom}%`;
  }
  if (wordPage) {
    wordPage.style.transform = `scale(${scale})`;
  }
  if (pptStage) {
    pptStage.style.transform = `scale(${scale})`;
  }
}

function adjustWorkspaceZoom(delta) {
  setWorkspaceZoom(workspaceZoom + delta);
}

function downloadCurrentDocument() {
  if (currentMode === 'excel') exportExcelXLSX();
  else if (currentMode === 'word') exportWordDoc();
  else if (currentMode === 'ppt') exportPptxFile();
}

function createNewDocument() {
  if (currentMode === 'excel') {
    if (!confirm('현재 시트를 비우고 새 빈 엑셀 통합 문서를 만드시겠습니까?')) return;
    excelSheets = [createBlankExcelSheet('Sheet1', 40, 15)];
    activeSheetIdx = 0;
    selectedRow = 0;
    selectedCol = 0;
    saveExcelState();
    renderExcelWorkspace();
    showToast('새 Excel 통합 문서가 생성되었습니다.');
  } else if (currentMode === 'word') {
    if (!confirm('현재 워드 문서를 비우고 새 문서를 작성하시겠습니까?')) return;
    document.getElementById('wordPageEditor').innerHTML =
      '<h1>새 문서 제목</h1><p>여기에 본문 내용을 입력하세요...</p>';
    saveWordState();
    updateWordCounters();
    showToast('새 Word 문서가 생성되었습니다.');
  } else if (currentMode === 'ppt') {
    if (!confirm('현재 슬라이드를 초기화하고 새 프레젠테이션을 만드시겠습니까?')) return;
    pptDeck = [createDefaultTitleSlide()];
    currentSlideIndex = 0;
    selectedElementId = null;
    savePptState();
    renderPptWorkspace();
    showToast('새 PowerPoint 프레젠테이션이 생성되었습니다.');
  }
}

/* ============================================================================
 * 1. NATIVE INSTANT EXCEL ENGINE
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

function createBlankExcelSheet(name = 'Sheet1', rows = 40, cols = 15) {
  return {
    name,
    rows,
    cols,
    cells: {}
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
      hideFormulaAutocomplete();
      confirmFormulaBarEdit();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      hideFormulaAutocomplete();
      cancelFormulaBarEdit();
    }
  });
  formulaInput.addEventListener('input', () => {
    setCellRawValue(selectedRow, selectedCol, formulaInput.value, false);
    updateFormulaAutocomplete(formulaInput.value);
  });
  formulaInput.addEventListener('blur', () => {
    setTimeout(hideFormulaAutocomplete, 180);
  });
}

function confirmFormulaBarEdit() {
  const formulaInput = document.getElementById('excelFormulaInput');
  setCellRawValue(selectedRow, selectedCol, formulaInput.value, true);
  selectExcelCell(Math.min(getActiveSheet().rows - 1, selectedRow + 1), selectedCol);
}

function cancelFormulaBarEdit() {
  updateExcelFormulaBarAndStats();
  document.getElementById('excelFormulaInput').blur();
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
  invalidateHyperFormulaCache();
  saveExcelState();
  if (fullRender) {
    renderExcelGrid();
    updateExcelFormulaBarAndStats();
  }
}

function getCellObj(sheet, r, c) {
  return sheet.cells[`${r},${c}`] || null;
}

/* ============================================================================
 * FULL 400+ EXCEL FORMULA ENGINE (HyperFormula + Native Built-in Engine)
 * Supports: VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, IF, IFS, IFERROR, SWITCH,
 * SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, COUNTA, ROUND, ROUNDUP, ROUNDDOWN,
 * LEFT, RIGHT, MID, LEN, TRIM, CONCAT, SUBSTITUTE, FIND, UPPER, LOWER, TEXT,
 * TODAY, NOW, DATE, YEAR, MONTH, DAY, DATEDIF, EOMONTH, PMT, PV, FV, NPV, etc.
 * ============================================================================ */

let hfInstance = null;
let hfCacheDirty = true;

function invalidateHyperFormulaCache() {
  hfCacheDirty = true;
}

function onHyperFormulaLoaded() {
  invalidateHyperFormulaCache();
  const badge = document.getElementById('formulaEngineBadge');
  if (badge) {
    badge.textContent = '⚡ 400+ 오리지널 엑셀 함수 엔진(HyperFormula) 활성';
  }
  if (currentMode === 'excel') {
    renderExcelGrid();
    updateExcelFormulaBarAndStats();
  }
}

function rebuildHyperFormulaIfNeeded() {
  if (typeof HyperFormula === 'undefined') return null;
  if (!hfCacheDirty && hfInstance) return hfInstance;

  try {
    const sheetsObj = {};
    excelSheets.forEach((sh) => {
      let maxR = Math.min(sh.rows, 60);
      let maxC = Math.min(sh.cols, 26);
      Object.keys(sh.cells).forEach((k) => {
        const [r, c] = k.split(',').map(Number);
        if (r + 1 > maxR) maxR = r + 1;
        if (c + 1 > maxC) maxC = c + 1;
      });

      const matrix = [];
      for (let r = 0; r < maxR; r++) {
        const rowArr = [];
        for (let c = 0; c < maxC; c++) {
          const cell = sh.cells[`${r},${c}`];
          if (!cell || cell.raw === '' || cell.raw === null || cell.raw === undefined) {
            rowArr.push(null);
          } else {
            const rawStr = String(cell.raw).trim();
            if (rawStr.startsWith('=')) {
              rowArr.push(rawStr);
            } else {
              const cleaned = rawStr.replace(/,/g, '');
              const num = Number(cleaned);
              if (cleaned !== '' && !isNaN(num) && /^-?\d+(\.\d+)?$/.test(cleaned)) {
                rowArr.push(num);
              } else {
                rowArr.push(rawStr);
              }
            }
          }
        }
        matrix.push(rowArr);
      }
      sheetsObj[sh.name] = matrix;
    });

    hfInstance = HyperFormula.buildFromSheets(sheetsObj, {
      licenseKey: 'gpl-v3'
    });
    hfCacheDirty = false;
    return hfInstance;
  } catch (e) {
    console.warn('HyperFormula build fallback:', e);
    return null;
  }
}

// Helper: Split top-level arguments by comma respecting quotes & nested parentheses
function splitFormulaArgs(argString) {
  const args = [];
  let current = '';
  let depth = 0;
  let inQuotes = false;

  for (let i = 0; i < argString.length; i++) {
    const ch = argString[i];
    if (ch === '"') {
      inQuotes = !inQuotes;
      current += ch;
    } else if (!inQuotes && ch === '(') {
      depth++;
      current += ch;
    } else if (!inQuotes && ch === ')') {
      depth--;
      current += ch;
    } else if (!inQuotes && depth === 0 && ch === ',') {
      args.push(current.trim());
      current = '';
    } else {
      current += ch;
    }
  }
  if (current.trim() !== '' || argString.endsWith(',')) {
    args.push(current.trim());
  }
  return args;
}

function getRangeMatrixValues(sheet, rangeStr, visited) {
  const clean = rangeStr.replace(/\$/g, '').trim();
  const parts = clean.split(':');
  if (parts.length !== 2) return [];
  const s = addrToCellCoord(parts[0]);
  const e = addrToCellCoord(parts[1]);
  if (!s || !e) return [];
  const r1 = Math.min(s.r, e.r);
  const r2 = Math.max(s.r, e.r);
  const c1 = Math.min(s.c, e.c);
  const c2 = Math.max(s.c, e.c);

  const matrix = [];
  for (let r = r1; r <= r2; r++) {
    const row = [];
    for (let c = c1; c <= c2; c++) {
      row.push(evaluateCellValue(sheet, r, c, new Set(visited)));
    }
    matrix.push(row);
  }
  return matrix;
}

function flattenMatrix(mat) {
  const out = [];
  mat.forEach((row) => row.forEach((v) => out.push(v)));
  return out;
}

function matchesExcelCriterion(cellVal, criterionRaw) {
  const crit = String(criterionRaw).replace(/^"|"$/g, '').trim();
  const m = crit.match(/^(>=|<=|<>|>|<|=)(.*)$/);
  if (m) {
    const op = m[1];
    const rhsRaw = m[2].trim();
    const lhsNum = Number(cellVal);
    const rhsNum = Number(rhsRaw);
    const bothNum = cellVal !== '' && rhsRaw !== '' && !isNaN(lhsNum) && !isNaN(rhsNum);
    const a = bothNum ? lhsNum : String(cellVal);
    const b = bothNum ? rhsNum : rhsRaw;
    if (op === '>') return a > b;
    if (op === '>=') return a >= b;
    if (op === '<') return a < b;
    if (op === '<=') return a <= b;
    if (op === '<>') return a !== b;
    if (op === '=') return a === b;
  }
  const n1 = Number(cellVal);
  const n2 = Number(crit);
  if (cellVal !== '' && crit !== '' && !isNaN(n1) && !isNaN(n2)) {
    return n1 === n2;
  }
  return String(cellVal).toLowerCase() === crit.toLowerCase();
}

function evalSingleFormulaNode(sheet, exprStr, visited) {
  const s = String(exprStr).trim();
  if (s === '') return '';

  // Quoted string literal
  if (s.startsWith('"') && s.endsWith('"') && s.length >= 2) {
    return s.slice(1, -1);
  }
  // Boolean literals
  if (s.toUpperCase() === 'TRUE') return true;
  if (s.toUpperCase() === 'FALSE') return false;

  // Number literal
  if (/^-?\d+(\.\d+)?$/.test(s)) {
    return Number(s);
  }

  // Direct Cell Reference (e.g. A1, $C$2)
  const cleanCell = s.replace(/\$/g, '');
  if (/^[A-Za-z]{1,2}\d{1,3}$/.test(cleanCell)) {
    const coord = addrToCellCoord(cleanCell);
    if (coord) return evaluateCellValue(sheet, coord.r, coord.c, new Set(visited));
  }

  // Function call: FNNAME(...)
  const fnMatch = s.match(/^([A-Za-z0-9_]+)\s*\(([\s\S]*)\)$/);
  if (fnMatch) {
    const fn = fnMatch[1].toUpperCase();
    const argsRaw = splitFormulaArgs(fnMatch[2]);

    const evalArg = (idx) => (idx < argsRaw.length ? evalSingleFormulaNode(sheet, argsRaw[idx], visited) : '');
    const getFlatValues = (rawArg) => {
      if (rawArg.includes(':')) {
        return flattenMatrix(getRangeMatrixValues(sheet, rawArg, visited));
      }
      return [evalSingleFormulaNode(sheet, rawArg, visited)];
    };
    const getNumsFromArgs = () => {
      const nums = [];
      argsRaw.forEach((arg) => {
        getFlatValues(arg).forEach((v) => {
          const n = Number(v);
          if (v !== '' && v !== null && v !== false && v !== true && !isNaN(n)) nums.push(n);
        });
      });
      return nums;
    };

    if (fn === 'SUM') return getNumsFromArgs().reduce((a, b) => a + b, 0);
    if (fn === 'AVERAGE' || fn === 'AVG') {
      const nums = getNumsFromArgs();
      return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
    }
    if (fn === 'MIN') {
      const nums = getNumsFromArgs();
      return nums.length ? Math.min(...nums) : 0;
    }
    if (fn === 'MAX') {
      const nums = getNumsFromArgs();
      return nums.length ? Math.max(...nums) : 0;
    }
    if (fn === 'COUNT') return getNumsFromArgs().length;
    if (fn === 'COUNTA') {
      let cnt = 0;
      argsRaw.forEach((arg) => {
        getFlatValues(arg).forEach((v) => {
          if (v !== '' && v !== null && v !== undefined) cnt++;
        });
      });
      return cnt;
    }
    if (fn === 'COUNTBLANK') {
      let cnt = 0;
      argsRaw.forEach((arg) => {
        getFlatValues(arg).forEach((v) => {
          if (v === '' || v === null || v === undefined) cnt++;
        });
      });
      return cnt;
    }
    if (fn === 'MEDIAN') {
      const nums = getNumsFromArgs().sort((a, b) => a - b);
      if (!nums.length) return 0;
      const mid = Math.floor(nums.length / 2);
      return nums.length % 2 !== 0 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
    }
    if (fn === 'PRODUCT') {
      const nums = getNumsFromArgs();
      return nums.length ? nums.reduce((a, b) => a * b, 1) : 0;
    }
    if (fn === 'ROUND') {
      const n = Number(evalArg(0)) || 0;
      const d = Number(evalArg(1)) || 0;
      const f = Math.pow(10, d);
      return Math.round(n * f) / f;
    }
    if (fn === 'ROUNDUP') {
      const n = Number(evalArg(0)) || 0;
      const d = Number(evalArg(1)) || 0;
      const f = Math.pow(10, d);
      return (n >= 0 ? Math.ceil(n * f) : Math.floor(n * f)) / f;
    }
    if (fn === 'ROUNDDOWN' || fn === 'TRUNC') {
      const n = Number(evalArg(0)) || 0;
      const d = Number(evalArg(1)) || 0;
      const f = Math.pow(10, d);
      return Math.trunc(n * f) / f;
    }
    if (fn === 'INT') return Math.floor(Number(evalArg(0)) || 0);
    if (fn === 'ABS') return Math.abs(Number(evalArg(0)) || 0);
    if (fn === 'SQRT') return Math.sqrt(Number(evalArg(0)) || 0);
    if (fn === 'POWER') return Math.pow(Number(evalArg(0)) || 0, Number(evalArg(1)) || 0);
    if (fn === 'MOD') return (Number(evalArg(0)) || 0) % (Number(evalArg(1)) || 1);

    // Logical: IF, IFS, IFERROR, AND, OR, NOT
    if (fn === 'IF') {
      const cond = Boolean(evalArg(0));
      return cond ? evalArg(1) : (argsRaw.length > 2 ? evalArg(2) : false);
    }
    if (fn === 'IFS') {
      for (let i = 0; i < argsRaw.length - 1; i += 2) {
        if (Boolean(evalArg(i))) return evalArg(i + 1);
      }
      return '#N/A';
    }
    if (fn === 'IFERROR') {
      const v = evalArg(0);
      return String(v).startsWith('#') ? evalArg(1) : v;
    }
    if (fn === 'AND') return argsRaw.every((_, i) => Boolean(evalArg(i)));
    if (fn === 'OR') return argsRaw.some((_, i) => Boolean(evalArg(i)));
    if (fn === 'NOT') return !Boolean(evalArg(0));

    // Conditional Aggregation: SUMIF, COUNTIF, AVERAGEIF
    if (fn === 'COUNTIF') {
      const rangeVals = getFlatValues(argsRaw[0] || '');
      const crit = evalArg(1);
      return rangeVals.filter((v) => matchesExcelCriterion(v, crit)).length;
    }
    if (fn === 'SUMIF') {
      const rangeVals = getFlatValues(argsRaw[0] || '');
      const crit = evalArg(1);
      const sumVals = argsRaw.length > 2 ? getFlatValues(argsRaw[2]) : rangeVals;
      let total = 0;
      rangeVals.forEach((v, idx) => {
        if (matchesExcelCriterion(v, crit)) {
          const n = Number(sumVals[idx]);
          if (!isNaN(n)) total += n;
        }
      });
      return total;
    }
    if (fn === 'AVERAGEIF') {
      const rangeVals = getFlatValues(argsRaw[0] || '');
      const crit = evalArg(1);
      const avgVals = argsRaw.length > 2 ? getFlatValues(argsRaw[2]) : rangeVals;
      let total = 0;
      let cnt = 0;
      rangeVals.forEach((v, idx) => {
        if (matchesExcelCriterion(v, crit)) {
          const n = Number(avgVals[idx]);
          if (!isNaN(n)) {
            total += n;
            cnt++;
          }
        }
      });
      return cnt ? total / cnt : 0;
    }

    // Lookup: VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH
    if (fn === 'VLOOKUP') {
      const lookupVal = evalArg(0);
      const mat = getRangeMatrixValues(sheet, argsRaw[1] || '', visited);
      const colIdx = (Number(evalArg(2)) || 1) - 1;
      for (let r = 0; r < mat.length; r++) {
        if (String(mat[r][0]).toLowerCase() === String(lookupVal).toLowerCase()) {
          return mat[r][colIdx] !== undefined ? mat[r][colIdx] : '#REF!';
        }
      }
      return '#N/A';
    }
    if (fn === 'HLOOKUP') {
      const lookupVal = evalArg(0);
      const mat = getRangeMatrixValues(sheet, argsRaw[1] || '', visited);
      const rowIdx = (Number(evalArg(2)) || 1) - 1;
      if (mat.length === 0) return '#N/A';
      for (let c = 0; c < mat[0].length; c++) {
        if (String(mat[0][c]).toLowerCase() === String(lookupVal).toLowerCase()) {
          return mat[rowIdx]?.[c] !== undefined ? mat[rowIdx][c] : '#REF!';
        }
      }
      return '#N/A';
    }
    if (fn === 'XLOOKUP') {
      const lookupVal = evalArg(0);
      const lookArr = getFlatValues(argsRaw[1] || '');
      const retArr = getFlatValues(argsRaw[2] || '');
      const notFound = argsRaw.length > 3 ? evalArg(3) : '#N/A';
      for (let i = 0; i < lookArr.length; i++) {
        if (String(lookArr[i]).toLowerCase() === String(lookupVal).toLowerCase()) {
          return retArr[i] !== undefined ? retArr[i] : notFound;
        }
      }
      return notFound;
    }
    if (fn === 'MATCH') {
      const lookupVal = evalArg(0);
      const lookArr = getFlatValues(argsRaw[1] || '');
      for (let i = 0; i < lookArr.length; i++) {
        if (String(lookArr[i]).toLowerCase() === String(lookupVal).toLowerCase()) {
          return i + 1;
        }
      }
      return '#N/A';
    }
    if (fn === 'INDEX') {
      const mat = getRangeMatrixValues(sheet, argsRaw[0] || '', visited);
      const rIdx = (Number(evalArg(1)) || 1) - 1;
      const cIdx = argsRaw.length > 2 ? (Number(evalArg(2)) || 1) - 1 : 0;
      return mat[rIdx]?.[cIdx] !== undefined ? mat[rIdx][cIdx] : '#REF!';
    }

    // Text functions: LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, CONCAT, CONCATENATE, SUBSTITUTE
    if (fn === 'LEFT') {
      const str = String(evalArg(0));
      const n = argsRaw.length > 1 ? Number(evalArg(1)) : 1;
      return str.slice(0, n);
    }
    if (fn === 'RIGHT') {
      const str = String(evalArg(0));
      const n = argsRaw.length > 1 ? Number(evalArg(1)) : 1;
      return str.slice(Math.max(0, str.length - n));
    }
    if (fn === 'MID') {
      const str = String(evalArg(0));
      const start = Math.max(1, Number(evalArg(1)) || 1) - 1;
      const len = Number(evalArg(2)) || 0;
      return str.slice(start, start + len);
    }
    if (fn === 'LEN') return String(evalArg(0)).length;
    if (fn === 'TRIM') return String(evalArg(0)).trim().replace(/\s+/g, ' ');
    if (fn === 'UPPER') return String(evalArg(0)).toUpperCase();
    if (fn === 'LOWER') return String(evalArg(0)).toLowerCase();
    if (fn === 'CONCAT' || fn === 'CONCATENATE') {
      return argsRaw.map((_, i) => String(evalArg(i))).join('');
    }
    if (fn === 'SUBSTITUTE') {
      const str = String(evalArg(0));
      const oldTxt = String(evalArg(1));
      const newTxt = String(evalArg(2));
      return str.split(oldTxt).join(newTxt);
    }
    if (fn === 'VALUE') return Number(String(evalArg(0)).replace(/,/g, '')) || 0;

    // Date & Financial: TODAY, NOW, YEAR, MONTH, DAY, PMT
    if (fn === 'TODAY') return new Date().toISOString().slice(0, 10);
    if (fn === 'NOW') return new Date().toISOString().slice(0, 16).replace('T', ' ');
    if (fn === 'YEAR') return new Date(String(evalArg(0))).getFullYear() || new Date().getFullYear();
    if (fn === 'MONTH') return (new Date(String(evalArg(0))).getMonth() + 1) || (new Date().getMonth() + 1);
    if (fn === 'DAY') return new Date(String(evalArg(0))).getDate() || new Date().getDate();
    if (fn === 'PMT') {
      const rate = Number(evalArg(0)) || 0;
      const nper = Number(evalArg(1)) || 1;
      const pv = Number(evalArg(2)) || 0;
      if (rate === 0) return -(pv / nper);
      return Math.round((-(pv * rate) / (1 - Math.pow(1 + rate, -nper))) * 100) / 100;
    }
  }

  // Comparison expressions (e.g. D2>C2, A2="고정지출")
  const cmpMatch = s.match(/^([^><=]+)(>=|<=|<>|>|<|=)([^><=]+)$/);
  if (cmpMatch) {
    const lhs = evalSingleFormulaNode(sheet, cmpMatch[1], visited);
    const op = cmpMatch[2];
    const rhs = evalSingleFormulaNode(sheet, cmpMatch[3], visited);
    const n1 = Number(lhs);
    const n2 = Number(rhs);
    const bothNum = lhs !== '' && rhs !== '' && !isNaN(n1) && !isNaN(n2);
    const a = bothNum ? n1 : String(lhs);
    const b = bothNum ? n2 : String(rhs);
    if (op === '>') return a > b;
    if (op === '>=') return a >= b;
    if (op === '<') return a < b;
    if (op === '<=') return a <= b;
    if (op === '<>') return a !== b;
    if (op === '=') return a === b;
  }

  return null;
}

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
    const cleaned = rawStr.replace(/,/g, '');
    const num = Number(cleaned);
    if (cleaned !== '' && !isNaN(num) && /^-?\d+(\.\d+)?$/.test(cleaned)) {
      return num;
    }
    return rawStr;
  }

  const exprBody = rawStr.slice(1).trim();

  // 1. Try Native Built-in Rich Evaluator First (Handles XLOOKUP, VLOOKUP, SUMIF, COUNTIF, IF, etc. immediately at 0ms)
  try {
    const directRes = evalSingleFormulaNode(sheet, exprBody, visited);
    if (directRes !== null && directRes !== undefined) {
      if (typeof directRes === 'number' && isFinite(directRes)) {
        return Math.round(directRes * 100) / 100;
      }
      return directRes;
    }
  } catch (e) {}

  // 2. Try HyperFormula (400+ Full Original Excel Functions) if loaded
  if (typeof HyperFormula !== 'undefined') {
    const hf = rebuildHyperFormulaIfNeeded();
    if (hf) {
      try {
        const sheetId = hf.getSheetId(sheet.name);
        if (sheetId !== undefined) {
          const hfVal = hf.getCellValue({ sheet: sheetId, col: c, row: r });
          if (hfVal !== null && hfVal !== undefined) {
            if (typeof hfVal === 'object' && hfVal.value) {
              return hfVal.value; // Error code like #VALUE!, #NAME?
            }
            if (typeof hfVal === 'number' && isFinite(hfVal)) {
              return Math.round(hfVal * 100) / 100;
            }
            return hfVal;
          }
        }
      } catch (e) {}
    }
  }

  // 3. Fallback arithmetic with inline cell & function substitution
  try {
    let expr = exprBody;

    expr = expr.replace(/\b(SUM|AVERAGE|AVG|MIN|MAX|COUNT)\s*\(\s*([A-Za-z]+\d+)\s*:\s*([A-Za-z]+\d+)\s*\)/gi, (_, fn, startAddr, endAddr) => {
      return String(evalSingleFormulaNode(sheet, `${fn}(${startAddr}:${endAddr})`, visited) || 0);
    });

    expr = expr.replace(/\$?([A-Za-z]{1,2})\$?(\d{1,3})\b/g, (_, colL, rowN) => {
      const coord = addrToCellCoord(`${colL}${rowN}`);
      if (!coord) return '0';
      const val = evaluateCellValue(sheet, coord.r, coord.c, new Set(visited));
      const n = Number(val);
      return val !== '' && !isNaN(n) ? String(n) : '0';
    });

    if (!/^[0-9+\-*/().\s]+$/.test(expr)) {
      return '#VALUE!';
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

/* ==================== EXCEL FUNCTION CATALOG & WIZARD MODAL ==================== */
const EXCEL_FUNCTION_CATALOG = [
  // 수학/삼각 (MATH)
  { name: 'SUM', cat: 'MATH', syntax: '=SUM(C2:C7)', desc: '지정한 범위의 모든 숫자를 더합니다.' },
  { name: 'SUMIF', cat: 'MATH', syntax: '=SUMIF(A2:A7, "고정지출", D2:D7)', desc: '주어진 조건에 맞는 셀들의 합계를 구합니다.' },
  { name: 'SUMIFS', cat: 'MATH', syntax: '=SUMIFS(D2:D7, A2:A7, "고정지출", C2:C7, ">10000")', desc: '여러 조건에 모두 맞는 셀들의 합계를 구합니다.' },
  { name: 'SUMPRODUCT', cat: 'MATH', syntax: '=SUMPRODUCT(C2:C5, D2:D5)', desc: '배열 또는 범위의 대응하는 값끼리 곱한 뒤 그 합계를 구합니다.' },
  { name: 'ROUND', cat: 'MATH', syntax: '=ROUND(E2, 0)', desc: '숫자를 지정한 자릿수로 반올림합니다.' },
  { name: 'ROUNDUP', cat: 'MATH', syntax: '=ROUNDUP(E2, 0)', desc: '숫자를 지정한 자릿수로 무조건 올림합니다.' },
  { name: 'ROUNDDOWN', cat: 'MATH', syntax: '=ROUNDDOWN(E2, 0)', desc: '숫자를 지정한 자릿수로 무조건 내림(버림)합니다.' },
  { name: 'SUBTOTAL', cat: 'MATH', syntax: '=SUBTOTAL(9, C2:C7)', desc: '목록이나 데이터베이스의 부분합(9=합계, 1=평균)을 구합니다.' },
  { name: 'ABS', cat: 'MATH', syntax: '=ABS(E2)', desc: '숫자의 부호를 제외한 절댓값을 반환합니다.' },
  { name: 'INT', cat: 'MATH', syntax: '=INT(E2)', desc: '숫자를 가장 가까운 정수로 내림합니다.' },
  { name: 'MOD', cat: 'MATH', syntax: '=MOD(C2, 10)', desc: '나눗셈의 나머지를 반환합니다.' },
  { name: 'POWER', cat: 'MATH', syntax: '=POWER(1.12, 5)', desc: '거듭제곱 값을 계산합니다.' },
  { name: 'SQRT', cat: 'MATH', syntax: '=SQRT(C2)', desc: '양의 제곱근을 구합니다.' },
  { name: 'PRODUCT', cat: 'MATH', syntax: '=PRODUCT(C2:D2)', desc: '인수로 지정한 숫자를 모두 곱합니다.' },

  // 통계 (STAT)
  { name: 'AVERAGE', cat: 'STAT', syntax: '=AVERAGE(C2:C7)', desc: '지정한 범위의 산술 평균을 구합니다.' },
  { name: 'AVERAGEIF', cat: 'STAT', syntax: '=AVERAGEIF(A2:A7, "고정지출", D2:D7)', desc: '조건에 맞는 셀들의 평균을 구합니다.' },
  { name: 'COUNT', cat: 'STAT', syntax: '=COUNT(C2:C7)', desc: '범위 내에서 숫자가 포함된 셀의 개수를 셉니다.' },
  { name: 'COUNTA', cat: 'STAT', syntax: '=COUNTA(A2:A7)', desc: '범위 내에서 비어 있지 않은 모든 셀의 개수를 셉니다.' },
  { name: 'COUNTBLANK', cat: 'STAT', syntax: '=COUNTBLANK(A2:A10)', desc: '범위 내에서 비어 있는 셀의 개수를 셉니다.' },
  { name: 'COUNTIF', cat: 'STAT', syntax: '=COUNTIF(A2:A7, "고정지출")', desc: '지정한 조건에 맞는 셀의 개수를 구합니다.' },
  { name: 'COUNTIFS', cat: 'STAT', syntax: '=COUNTIFS(A2:A7, "고정지출", D2:D7, ">50000")', desc: '여러 조건을 충족하는 셀의 개수를 구합니다.' },
  { name: 'MAX', cat: 'STAT', syntax: '=MAX(C2:C7)', desc: '범위 내에서 가장 큰 최댓값을 반환합니다.' },
  { name: 'MIN', cat: 'STAT', syntax: '=MIN(C2:C7)', desc: '범위 내에서 가장 작은 최솟값을 반환합니다.' },
  { name: 'MEDIAN', cat: 'STAT', syntax: '=MEDIAN(C2:C7)', desc: '주어진 숫자 집합의 중앙값을 반환합니다.' },
  { name: 'LARGE', cat: 'STAT', syntax: '=LARGE(C2:C7, 2)', desc: '데이터 집합에서 k번째로 큰 값을 반환합니다.' },
  { name: 'SMALL', cat: 'STAT', syntax: '=SMALL(C2:C7, 2)', desc: '데이터 집합에서 k번째로 작은 값을 반환합니다.' },
  { name: 'STDEV.S', cat: 'STAT', syntax: '=STDEV.S(C2:C7)', desc: '표본 집단의 표준 편차를 구합니다.' },

  // 논리 (LOGIC)
  { name: 'IF', cat: 'LOGIC', syntax: '=IF(D2>C2, "예산초과", "정상")', desc: '조건이 참일 때와 거짓일 때 각각 다른 값을 반환합니다.' },
  { name: 'IFS', cat: 'LOGIC', syntax: '=IFS(D2>500000, "고액", D2>100000, "중간", TRUE, "소액")', desc: '여러 조건을 순차적으로 검사하여 첫 번째 참인 결과를 반환합니다.' },
  { name: 'IFERROR', cat: 'LOGIC', syntax: '=IFERROR(C2/D2, 0)', desc: '수식에서 오류(#DIV/0!, #N/A 등)가 발생하면 지정한 대체 값을 반환합니다.' },
  { name: 'AND', cat: 'LOGIC', syntax: '=AND(C2>0, D2>0)', desc: '모든 인수가 참(TRUE)이면 TRUE를 반환합니다.' },
  { name: 'OR', cat: 'LOGIC', syntax: '=OR(C2>500000, D2>500000)', desc: '인수 중 하나라도 참이면 TRUE를 반환합니다.' },
  { name: 'NOT', cat: 'LOGIC', syntax: '=NOT(C2=0)', desc: '논리값의 반대(TRUE↔FALSE)를 반환합니다.' },
  { name: 'SWITCH', cat: 'LOGIC', syntax: '=SWITCH(A2, "고정지출", "필수", "저축/투자", "자산", "기타")', desc: '값 목록과 식을 비교하여 일치하는 첫 번째 결과를 반환합니다.' },

  // 찾기/참조 (LOOKUP)
  { name: 'VLOOKUP', cat: 'LOOKUP', syntax: '=VLOOKUP("식비 / 마트 장보기", B2:E7, 3, 0)', desc: '표의 첫 열에서 값을 찾아 같은 행의 지정한 열 값을 가져옵니다.' },
  { name: 'XLOOKUP', cat: 'LOOKUP', syntax: '=XLOOKUP("통신비 및 인터넷 요금", B2:B7, D2:D7, "없음")', desc: '범위에서 항목을 검색하여 일치하는 항목을 반환하는 최신 검색 함수입니다.' },
  { name: 'HLOOKUP', cat: 'LOOKUP', syntax: '=HLOOKUP("월 예산 (원)", A1:F7, 2, 0)', desc: '표의 첫 행에서 값을 찾아 같은 열의 지정한 행 값을 가져옵니다.' },
  { name: 'INDEX', cat: 'LOOKUP', syntax: '=INDEX(A2:F7, 2, 3)', desc: '테이블이나 범위에서 지정한 행과 열 위치의 값을 반환합니다.' },
  { name: 'MATCH', cat: 'LOOKUP', syntax: '=MATCH("고정지출", A2:A7, 0)', desc: '범위에서 지정한 항목의 상대 위치(순번)를 찾습니다.' },
  { name: 'CHOOSE', cat: 'LOOKUP', syntax: '=CHOOSE(2, "1등급", "2등급", "3등급")', desc: '인덱스 번호를 사용하여 값 목록에서 값을 선택합니다.' },

  // 텍스트 (TEXT)
  { name: 'LEFT', cat: 'TEXT', syntax: '=LEFT(A2, 2)', desc: '텍스트 문자열의 시작(왼쪽)부터 지정한 수만큼 문자를 추출합니다.' },
  { name: 'RIGHT', cat: 'TEXT', syntax: '=RIGHT(A2, 2)', desc: '텍스트 문자열의 끝(오른쪽)부터 지정한 수만큼 문자를 추출합니다.' },
  { name: 'MID', cat: 'TEXT', syntax: '=MID(B2, 1, 4)', desc: '문자열의 지정한 시작 위치부터 원하는 개수만큼 문자를 추출합니다.' },
  { name: 'LEN', cat: 'TEXT', syntax: '=LEN(B2)', desc: '텍스트 문자열의 글자 수를 반환합니다.' },
  { name: 'CONCAT', cat: 'TEXT', syntax: '=CONCAT(A2, " - ", B2)', desc: '여러 텍스트 항목을 하나의 문자열로 결합합니다.' },
  { name: 'SUBSTITUTE', cat: 'TEXT', syntax: '=SUBSTITUTE(B2, "요금", "비용")', desc: '텍스트 문자열에서 기존 텍스트를 새 텍스트로 바꿉니다.' },
  { name: 'TRIM', cat: 'TEXT', syntax: '=TRIM(B2)', desc: '텍스트의 앞뒤 공백 및 중복 공백을 제거합니다.' },
  { name: 'UPPER', cat: 'TEXT', syntax: '=UPPER("newt office")', desc: '영문 텍스트를 모두 대문자로 변환합니다.' },
  { name: 'LOWER', cat: 'TEXT', syntax: '=LOWER("NEWT OFFICE")', desc: '영문 텍스트를 모두 소문자로 변환합니다.' },
  { name: 'VALUE', cat: 'TEXT', syntax: '=VALUE("125000")', desc: '숫자 형식의 텍스트를 실제 숫자로 변환합니다.' },

  // 날짜/시간 (DATE)
  { name: 'TODAY', cat: 'DATE', syntax: '=TODAY()', desc: '오늘 날짜(YYYY-MM-DD)를 반환합니다.' },
  { name: 'NOW', cat: 'DATE', syntax: '=NOW()', desc: '현재 날짜와 시간을 반환합니다.' },
  { name: 'YEAR', cat: 'DATE', syntax: '=YEAR(TODAY())', desc: '날짜에서 연도(Year)를 추출합니다.' },
  { name: 'MONTH', cat: 'DATE', syntax: '=MONTH(TODAY())', desc: '날짜에서 월(1~12)을 추출합니다.' },
  { name: 'DAY', cat: 'DATE', syntax: '=DAY(TODAY())', desc: '날짜에서 일(1~31)을 추출합니다.' },
  { name: 'DATE', cat: 'DATE', syntax: '=DATE(2026, 12, 31)', desc: '연, 월, 일을 조합하여 날짜를 생성합니다.' },
  { name: 'DATEDIF', cat: 'DATE', syntax: '=DATEDIF("2026-01-01", "2026-12-31", "D")', desc: '두 날짜 사이의 일수("D"), 개월수("M"), 연수("Y")를 계산합니다.' },
  { name: 'EOMONTH', cat: 'DATE', syntax: '=EOMONTH("2026-11-01", 0)', desc: '지정한 달의 마지막 날짜(말일)를 구합니다.' },
  { name: 'NETWORKDAYS', cat: 'DATE', syntax: '=NETWORKDAYS("2026-11-01", "2026-11-30")', desc: '두 날짜 사이의 주말을 제외한 실제 근무일수를 계산합니다.' },

  // 재무 (FINANCE)
  { name: 'PMT', cat: 'FINANCE', syntax: '=PMT(0.05/12, 36, 30000000)', desc: '고정 이자율과 정기 상환 방식에 따른 대출 월 상환금을 계산합니다.' },
  { name: 'PV', cat: 'FINANCE', syntax: '=PV(0.05/12, 36, -900000)', desc: '투자 또는 대출의 현재 가치(Present Value)를 구합니다.' },
  { name: 'FV', cat: 'FINANCE', syntax: '=FV(0.06/12, 60, -1000000)', desc: '정기 적금/투자의 미래 가치(Future Value)를 계산합니다.' },
  { name: 'NPV', cat: 'FINANCE', syntax: '=NPV(0.08, C2:C6)', desc: '할인율과 미래 현금 흐름을 사용해 순현재가치를 계산합니다.' },

  // 정보 (INFO)
  { name: 'ISBLANK', cat: 'INFO', syntax: '=ISBLANK(F2)', desc: '셀이 비어 있으면 TRUE를 반환합니다.' },
  { name: 'ISNUMBER', cat: 'INFO', syntax: '=ISNUMBER(C2)', desc: '값이 숫자이면 TRUE를 반환합니다.' },
  { name: 'ISTEXT', cat: 'INFO', syntax: '=ISTEXT(A2)', desc: '값이 텍스트이면 TRUE를 반환합니다.' },
  { name: 'ISERROR', cat: 'INFO', syntax: '=ISERROR(C2/0)', desc: '값이 오류 값이면 TRUE를 반환합니다.' }
];

function openFunctionWizardModal(category = 'ALL') {
  if (currentMode !== 'excel') switchOfficeMode('excel');
  const modal = document.getElementById('functionWizardModal');
  const catSelect = document.getElementById('fxCategorySelect');
  const searchInput = document.getElementById('fxSearchInput');
  if (catSelect) catSelect.value = category;
  if (searchInput) searchInput.value = '';
  modal.classList.remove('hidden');
  renderFunctionWizardList();
  setTimeout(() => searchInput?.focus(), 60);
}

function closeFunctionWizardModal() {
  document.getElementById('functionWizardModal').classList.add('hidden');
}

function renderFunctionWizardList() {
  const listEl = document.getElementById('fxWizardList');
  const q = (document.getElementById('fxSearchInput')?.value || '').trim().toLowerCase();
  const cat = document.getElementById('fxCategorySelect')?.value || 'ALL';

  const catLabels = {
    MATH: '수학/삼각',
    STAT: '통계',
    LOGIC: '논리',
    LOOKUP: '찾기/참조',
    TEXT: '텍스트',
    DATE: '날짜/시간',
    FINANCE: '재무',
    INFO: '정보'
  };

  const filtered = EXCEL_FUNCTION_CATALOG.filter((item) => {
    if (cat !== 'ALL' && item.cat !== cat) return false;
    if (!q) return true;
    return (
      item.name.toLowerCase().includes(q) ||
      item.desc.toLowerCase().includes(q) ||
      item.syntax.toLowerCase().includes(q)
    );
  });

  if (filtered.length === 0) {
    listEl.innerHTML = `<p style="padding:20px;color:#605e5c;font-size:13px;">검색어와 일치하는 기본 카탈로그 함수가 없습니다. 수식 입력줄에 <strong>=${escapeHtml(q.toUpperCase())}(...)</strong> 형태로 직접 입력하시면 400+ HyperFormula 엔진이 자동 계산합니다.</p>`;
    return;
  }

  listEl.innerHTML = filtered
    .map(
      (item) => `
      <div class="fx-card" onclick="applyWizardFunction('${item.syntax.replace(/'/g, "\\'").replace(/"/g, '&quot;')}')">
        <div class="fx-card-left">
          <div class="fx-card-syntax">${escapeHtml(item.syntax)}</div>
          <div class="fx-card-desc"><strong>${escapeHtml(item.name)}</strong> — ${escapeHtml(item.desc)}</div>
        </div>
        <span class="fx-card-cat">${catLabels[item.cat] || item.cat}</span>
      </div>
    `
    )
    .join('');
}

function applyWizardFunction(syntaxExample) {
  closeFunctionWizardModal();
  setCellRawValue(selectedRow, selectedCol, syntaxExample, true);
  const formulaInput = document.getElementById('excelFormulaInput');
  if (formulaInput) {
    formulaInput.value = syntaxExample;
    formulaInput.focus();
  }
  showToast(`${cellCoordToAddr(selectedRow, selectedCol)} 셀에 ${syntaxExample} 수식이 삽입되었습니다.`);
}

/* ---- Formula Bar Live Autocomplete Dropdown ---- */
function updateFormulaAutocomplete(val) {
  const popup = document.getElementById('formulaAutocompletePopup');
  if (!popup) return;
  const trimmed = String(val || '').trim();
  if (!trimmed.startsWith('=')) {
    popup.classList.add('hidden');
    return;
  }

  // Match last function token after '=' or '(' or ','
  const tokenMatch = trimmed.match(/[=,(+\-*/]\s*([A-Za-z]{1,12})$/);
  if (!tokenMatch) {
    popup.classList.add('hidden');
    return;
  }

  const prefix = tokenMatch[1].toUpperCase();
  const matches = EXCEL_FUNCTION_CATALOG.filter((f) => f.name.startsWith(prefix)).slice(0, 8);
  if (matches.length === 0) {
    popup.classList.add('hidden');
    return;
  }

  popup.innerHTML = matches
    .map(
      (f) => `
      <div class="ac-item" onmousedown="insertAutocompletedFunction('${f.name}', '${f.syntax.replace(/'/g, "\\'").replace(/"/g, '&quot;')}')">
        <strong>${f.name}</strong>
        <span>${escapeHtml(f.desc)}</span>
      </div>
    `
    )
    .join('');
  popup.classList.remove('hidden');
}

function hideFormulaAutocomplete() {
  const popup = document.getElementById('formulaAutocompletePopup');
  if (popup) popup.classList.add('hidden');
}

function insertAutocompletedFunction(fnName, sampleSyntax) {
  const formulaInput = document.getElementById('excelFormulaInput');
  const current = formulaInput.value.trim();
  if (/^=[A-Za-z]+$/.test(current)) {
    formulaInput.value = sampleSyntax;
  } else {
    formulaInput.value = current.replace(/([A-Za-z]+)$/, `${fnName}(`);
  }
  setCellRawValue(selectedRow, selectedCol, formulaInput.value, true);
  hideFormulaAutocomplete();
  setTimeout(() => formulaInput.focus(), 20);
}

function formatDisplayValue(sheet, r, c) {
  const cell = getCellObj(sheet, r, c);
  const val = evaluateCellValue(sheet, r, c);
  if (val === '' || val === null || val === undefined) return '';
  if (typeof val === 'number') {
    if (cell && cell.numFmt === 'currency') {
      return '₩ ' + val.toLocaleString('ko-KR');
    }
    if (cell && cell.numFmt === 'percent') {
      return val.toLocaleString('ko-KR') + '%';
    }
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
  if (!container) return;
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
      const newName = prompt('시트 이름 변경:', sh.name);
      if (newName && newName.trim()) {
        sh.name = newName.trim();
        saveExcelState();
        renderExcelSheetTabs();
      }
    };
    container.appendChild(btn);
  });
}

function cycleExcelSheet(delta) {
  const next = activeSheetIdx + delta;
  if (next < 0 || next >= excelSheets.length) return;
  activeSheetIdx = next;
  renderExcelWorkspace();
}

function addNewExcelSheet() {
  const newName = `Sheet${excelSheets.length + 1}`;
  excelSheets.push(createBlankExcelSheet(newName, 40, 15));
  activeSheetIdx = excelSheets.length - 1;
  selectedRow = 0;
  selectedCol = 0;
  saveExcelState();
  renderExcelWorkspace();
  showToast(`'${newName}' 시트가 추가되었습니다.`);
}

function renderExcelGrid() {
  const sheet = getActiveSheet();
  const table = document.getElementById('excelGridTable');
  if (!sheet || !table) return;

  table.classList.toggle('hide-gridlines', !showExcelGridlines);

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
        if (cell.italic) styleStr += `font-style:italic;`;
        if (cell.underline) styleStr += `text-decoration:underline;`;
        if (cell.font) styleStr += `font-family:${cell.font};`;
        if (cell.fontSize) styleStr += `font-size:${cell.fontSize};`;
        if (cell.border) styleStr += `border:1.5px solid #323130;`;
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
  document.getElementById('statusReadyText').textContent = '입력';

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
      document.getElementById('statusReadyText').textContent = '준비';
      renderExcelGrid();
    }
  });

  input.addEventListener('blur', () => {
    if (editingCellInput) commitInlineCellEdit();
  });
}

function commitInlineCellEdit() {
  if (!editingCellInput) return;
  const { input, r, c } = editingCellInput;
  const val = input.value;
  editingCellInput = null;
  document.getElementById('statusReadyText').textContent = '준비';
  setCellRawValue(r, c, val, true);
}

function updateExcelFormulaBarAndStats() {
  const sheet = getActiveSheet();
  const addr = cellCoordToAddr(selectedRow, selectedCol);
  const addrBox = document.getElementById('excelCellAddressBox');
  if (addrBox) addrBox.textContent = addr;

  const cell = getCellObj(sheet, selectedRow, selectedCol);
  const formulaInput = document.getElementById('excelFormulaInput');
  if (formulaInput && document.activeElement !== formulaInput) {
    formulaInput.value = cell ? cell.raw : '';
  }

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

  if (currentMode === 'excel') {
    document.getElementById('statusReadyText').textContent = '준비';
    document.getElementById('statusContextInfo').textContent =
      `셀 ${addr}  ·  시트 ${activeSheetIdx + 1} / ${excelSheets.length}`;
    document.getElementById('statusMetricsText').innerHTML =
      `평균: ${avg.toLocaleString('ko-KR')} &nbsp;&nbsp;&nbsp; 개수: ${count} &nbsp;&nbsp;&nbsp; 합계: ${sum.toLocaleString('ko-KR')}`;
  }
}

/* ---- Excel Ribbon Actions ---- */
function copyExcelCell() {
  const sheet = getActiveSheet();
  const cell = getCellObj(sheet, selectedRow, selectedCol);
  excelClipboard = cell ? JSON.parse(JSON.stringify(cell)) : { raw: '' };
  showToast(`${cellCoordToAddr(selectedRow, selectedCol)} 셀이 복사되었습니다.`);
}

function cutExcelCell() {
  copyExcelCell();
  setCellRawValue(selectedRow, selectedCol, '', true);
}

function pasteExcelCell() {
  if (!excelClipboard) {
    showToast('복사된 셀이 없습니다.');
    return;
  }
  const sheet = getActiveSheet();
  sheet.cells[`${selectedRow},${selectedCol}`] = JSON.parse(JSON.stringify(excelClipboard));
  saveExcelState();
  renderExcelGrid();
  updateExcelFormulaBarAndStats();
}

function formatSelectedExcelCells(prop, value) {
  const sheet = getActiveSheet();
  const key = `${selectedRow},${selectedCol}`;
  if (!sheet.cells[key]) sheet.cells[key] = { raw: '', comma: true };
  const cell = sheet.cells[key];

  if (prop === 'bold') cell.bold = !cell.bold;
  else if (prop === 'italic') cell.italic = !cell.italic;
  else if (prop === 'underline') cell.underline = !cell.underline;
  else if (prop === 'border') cell.border = !cell.border;
  else if (prop === 'font') cell.font = value;
  else if (prop === 'fontSize') cell.fontSize = value;
  else if (prop === 'align') cell.align = value;
  else if (prop === 'color') cell.color = value;
  else if (prop === 'bg') cell.bg = value;
  else if (prop === 'comma') cell.comma = cell.comma === false ? true : false;
  else if (prop === 'clearStyle') {
    sheet.cells[key] = { raw: cell.raw || '', comma: true };
  }

  saveExcelState();
  renderExcelGrid();
}

function applyExcelNumberPreset(preset) {
  const sheet = getActiveSheet();
  const key = `${selectedRow},${selectedCol}`;
  if (!sheet.cells[key]) sheet.cells[key] = { raw: '', comma: true };
  const cell = sheet.cells[key];

  if (preset === 'currency') {
    cell.numFmt = 'currency';
    cell.comma = true;
  } else if (preset === 'percent') {
    cell.numFmt = 'percent';
  } else if (preset === 'plain') {
    cell.numFmt = 'general';
    cell.comma = false;
  } else {
    cell.numFmt = 'general';
    cell.comma = true;
  }
  saveExcelState();
  renderExcelGrid();
}

function insertExcelRowAtCursor() {
  const sheet = getActiveSheet();
  const newCells = {};
  Object.keys(sheet.cells).forEach((k) => {
    const [r, c] = k.split(',').map(Number);
    if (r >= selectedRow) {
      newCells[`${r + 1},${c}`] = sheet.cells[k];
    } else {
      newCells[k] = sheet.cells[k];
    }
  });
  sheet.cells = newCells;
  sheet.rows++;
  saveExcelState();
  renderExcelGrid();
  showToast(`${selectedRow + 1}행 위에 새 행이 삽입되었습니다.`);
}

function deleteExcelRowAtCursor() {
  const sheet = getActiveSheet();
  const newCells = {};
  Object.keys(sheet.cells).forEach((k) => {
    const [r, c] = k.split(',').map(Number);
    if (r < selectedRow) {
      newCells[k] = sheet.cells[k];
    } else if (r > selectedRow) {
      newCells[`${r - 1},${c}`] = sheet.cells[k];
    }
  });
  sheet.cells = newCells;
  saveExcelState();
  renderExcelGrid();
  showToast(`${selectedRow + 1}행이 삭제되었습니다.`);
}

function insertExcelColAtCursor() {
  const sheet = getActiveSheet();
  const newCells = {};
  Object.keys(sheet.cells).forEach((k) => {
    const [r, c] = k.split(',').map(Number);
    if (c >= selectedCol) {
      newCells[`${r},${c + 1}`] = sheet.cells[k];
    } else {
      newCells[k] = sheet.cells[k];
    }
  });
  sheet.cells = newCells;
  sheet.cols++;
  saveExcelState();
  renderExcelGrid();
  showToast(`${colIdxToName(selectedCol)}열 왼쪽에 새 열이 삽입되었습니다.`);
}

function sortExcelColumn(dir = 'asc') {
  const sheet = getActiveSheet();
  // Collect data rows (skip row 0 header)
  let maxR = 0;
  Object.keys(sheet.cells).forEach((k) => {
    const [r] = k.split(',').map(Number);
    if (r > maxR) maxR = r;
  });
  if (maxR <= 1) return;

  const rowsData = [];
  for (let r = 1; r <= maxR; r++) {
    const rowMap = {};
    let hasVal = false;
    for (let c = 0; c < sheet.cols; c++) {
      if (sheet.cells[`${r},${c}`]) {
        rowMap[c] = sheet.cells[`${r},${c}`];
        if (rowMap[c].raw !== '') hasVal = true;
      }
    }
    if (hasVal) {
      rowsData.push({ r, sortVal: evaluateCellValue(sheet, r, selectedCol), rowMap });
    }
  }

  rowsData.sort((a, b) => {
    const va = a.sortVal;
    const vb = b.sortVal;
    if (typeof va === 'number' && typeof vb === 'number') {
      return dir === 'asc' ? va - vb : vb - va;
    }
    return dir === 'asc'
      ? String(va).localeCompare(String(vb), 'ko')
      : String(vb).localeCompare(String(va), 'ko');
  });

  // Keep row 0 header, rewrite rows 1..rowsData.length
  const newCells = {};
  for (let c = 0; c < sheet.cols; c++) {
    if (sheet.cells[`0,${c}`]) newCells[`0,${c}`] = sheet.cells[`0,${c}`];
  }
  rowsData.forEach((item, idx) => {
    const targetR = idx + 1;
    Object.keys(item.rowMap).forEach((c) => {
      newCells[`${targetR},${c}`] = item.rowMap[c];
    });
  });

  sheet.cells = newCells;
  saveExcelState();
  renderExcelGrid();
  showToast(`${colIdxToName(selectedCol)}열 기준으로 정렬되었습니다.`);
}

function insertTodayInExcel() {
  const today = new Date().toISOString().slice(0, 10);
  setCellRawValue(selectedRow, selectedCol, today, true);
}

function toggleExcelGridlines() {
  showExcelGridlines = !showExcelGridlines;
  renderExcelGrid();
}

function insertExcelFunction(fnName) {
  const startR = 1;
  const endR = Math.max(1, selectedRow);
  const colLetter = colIdxToName(selectedCol);
  const formula = `=${fnName}(${colLetter}${startR + 1}:${colLetter}${endR})`;
  setCellRawValue(selectedRow, selectedCol, formula, true);
  showToast(`${cellCoordToAddr(selectedRow, selectedCol)} 셀에 ${formula} 수식이 입력되었습니다.`);
}

function addExcelRows(count = 10) {
  const sheet = getActiveSheet();
  sheet.rows += count;
  saveExcelState();
  renderExcelGrid();
  showToast(`행 ${count}개가 추가되었습니다.`);
}

function addExcelCols(count = 5) {
  const sheet = getActiveSheet();
  sheet.cols += count;
  saveExcelState();
  renderExcelGrid();
  showToast(`열 ${count}개가 추가되었습니다.`);
}

function populateSheetFromAOA(sheetName, aoa, headerColor = '#217346') {
  const rows = Math.max(40, aoa.length + 18);
  const cols = Math.max(14, (aoa[0]?.length || 8) + 5);
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
        bg: isHeader ? headerColor : (isTotalRow ? '#e6f2ec' : ''),
        color: isHeader ? '#ffffff' : '#201f1e',
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
  let docTitle = '2026_월간예산_가계부';

  if (type === 'budget') {
    sheetName = '월간예산_가계부';
    docTitle = '2026_월간예산_가계부';
    aoa = [
      ['항목 구분', '세부 지출 내역', '월 예산 (원)', '실제 지출 (원)', '잔여 예산 (수식)', '비고 / 결제수단'],
      ['고정지출', '통신비 및 인터넷 요금', 85000, 82000, '=C2-D2', '자동이체'],
      ['고정지출', 'MS 365 오피스 구독료 해지!', 11900, 0, '=C3-D3', '웹 오피스 전환 100% 절감'],
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
  } else if (type === 'formulas') {
    sheetName = '실무함수_종합데모';
    docTitle = '엑셀_400함수_실무데모';
    aoa = [
      ['사번 코드', '성명', '소속 부서', '실적 금액 (원)', '목표 달성 (IF)', '부서코드 (LEFT)', '함수 검증 항목', '실제 수식 계산 결과'],
      ['EMP-101', '김민준', '영업팀', 85000000, '=IF(D2>=50000000,"달성","미달")', '=LEFT(A2,3)', '1. 영업팀 실적 합계 (SUMIF)', '=SUMIF(C2:C6,"영업팀",D2:D6)'],
      ['EMP-102', '이서연', '기획팀', 42000000, '=IF(D3>=50000000,"달성","미달")', '=LEFT(A3,3)', '2. 영업팀 인원수 (COUNTIF)', '=COUNTIF(C2:C6,"영업팀")'],
      ['EMP-103', '박도윤', '영업팀', 96000000, '=IF(D4>=50000000,"달성","미달")', '=LEFT(A4,3)', '3. EMP-103 실적 조회 (VLOOKUP)', '=VLOOKUP("EMP-103",A2:D6,4,0)'],
      ['EMP-104', '최지우', '개발팀', 64000000, '=IF(D5>=50000000,"달성","미달")', '=LEFT(A5,3)', '4. 박도윤 부서 조회 (XLOOKUP)', '=XLOOKUP("박도윤",B2:B6,C2:C6,"없음")'],
      ['EMP-105', '정하은', '영업팀', 73000000, '=IF(D6>=50000000,"달성","미달")', '=LEFT(A6,3)', '5. 전체 실적 중앙값 (MEDIAN)', '=MEDIAN(D2:D6)'],
      ['합계', '전체 집계', '=COUNTA(B2:B6)', '=SUM(D2:D6)', '=AVERAGE(D2:D6)', '=TODAY()', '6. 최고 실적자 금액 (MAX)', '=MAX(D2:D6)']
    ];
  }

  invalidateHyperFormulaCache();
  excelSheets = [populateSheetFromAOA(sheetName, aoa)];
  activeSheetIdx = 0;
  selectedRow = 1;
  selectedCol = 2;
  document.getElementById('docTitleInput').value = docTitle;
  saveExcelState();
  renderExcelWorkspace();
  if (!silent) showToast(`'${sheetName}' 서식이 적용되었습니다.`);
}

/* ---- Global Keyboard Navigation ---- */
function handleGlobalKeyDown(e) {
  const tag = document.activeElement?.tagName;
  const isWordEditor = document.activeElement?.id === 'wordPageEditor';

  if (e.ctrlKey && e.key.toLowerCase() === 's') {
    e.preventDefault();
    downloadCurrentDocument();
    return;
  }

  if (currentMode === 'excel' && !editingCellInput && tag !== 'INPUT' && tag !== 'TEXTAREA' && !isWordEditor) {
    if (e.ctrlKey && e.key.toLowerCase() === 'c') {
      e.preventDefault();
      copyExcelCell();
    } else if (e.ctrlKey && e.key.toLowerCase() === 'x') {
      e.preventDefault();
      cutExcelCell();
    } else if (e.ctrlKey && e.key.toLowerCase() === 'v') {
      e.preventDefault();
      pasteExcelCell();
    } else if (e.ctrlKey && e.key.toLowerCase() === 'b') {
      e.preventDefault();
      formatSelectedExcelCells('bold');
    } else if (e.key === 'ArrowDown') {
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
    showToast('엑셀 파서를 로드 중입니다. 1초 후 다시 시도해 주세요.');
    return;
  }
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array', cellFormula: true });

  excelSheets = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
    const rows = Math.max(40, aoa.length + 15);
    const cols = Math.max(15, (aoa[0]?.length || 10) + 5);
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
            bg: R === 0 ? '#f3f2f1' : '',
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
  showToast(`'${file.name}' (${excelSheets.length}개 시트) 열기 완료!`);
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
        rowArr.push(evaluateCellValue(sheet, r, c));
      }
      aoa.push(rowArr);
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);

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
  showToast(`'${fileName}' 저장 완료!`);
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
      showToast(`'${file.name}' 문서를 열었습니다.`);
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
 * 2. MS WORD ENGINE
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
}

function saveWordState() {
  const editor = document.getElementById('wordPageEditor');
  localStorage.setItem(STORAGE_KEYS.WORD_HTML, editor.innerHTML);
  markAutoSaved();
}

function updateWordCounters() {
  const editor = document.getElementById('wordPageEditor');
  if (!editor) return;
  const text = editor.innerText || '';
  const totalChars = text.replace(/\n/g, '').length;
  const noSpaceChars = text.replace(/\s+/g, '').length;
  const words = text.trim() ? text.trim().split(/\s+/).length : 0;

  if (currentMode === 'word') {
    document.getElementById('statusReadyText').textContent = '페이지 1 / 1';
    document.getElementById('statusContextInfo').textContent = '한국어 (대한민국)';
    document.getElementById('statusMetricsText').innerHTML =
      `단어: ${words.toLocaleString()}개 &nbsp;&nbsp;&nbsp; 글자 수: ${totalChars.toLocaleString()}자 (공백 제외 ${noSpaceChars.toLocaleString()}자)`;
  }
}

function alertWordCountReport() {
  const editor = document.getElementById('wordPageEditor');
  const text = editor.innerText || '';
  const totalChars = text.replace(/\n/g, '').length;
  const noSpaceChars = text.replace(/\s+/g, '').length;
  const words = text.trim() ? text.trim().split(/\s+/).length : 0;
  alert(`[Word 단어 수 통계]\n\n• 단어 수: ${words.toLocaleString()}개\n• 문자 수(공백 포함): ${totalChars.toLocaleString()}자\n• 문자 수(공백 제외): ${noSpaceChars.toLocaleString()}자`);
}

function execWordCmd(cmd, value = null) {
  document.getElementById('wordPageEditor').focus();
  document.execCommand(cmd, false, value);
  saveWordState();
  updateWordCounters();
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

function setWordPageMargin(mode) {
  const page = document.getElementById('wordPageEditor');
  if (mode === 'narrow') {
    page.style.padding = '12.7mm 12.7mm';
    showToast('좁은 여백(12.7mm)이 적용되었습니다.');
  } else {
    page.style.padding = '22mm 20mm';
    showToast('보통 여백(22mm)이 적용되었습니다.');
  }
}

function setWordPageOrientation(mode) {
  const page = document.getElementById('wordPageEditor');
  if (mode === 'landscape') {
    page.style.width = '297mm';
    page.style.minHeight = '210mm';
    showToast('가로 용지 방향으로 전환되었습니다.');
  } else {
    page.style.width = '210mm';
    page.style.minHeight = '297mm';
    showToast('세로 용지 방향으로 전환되었습니다.');
  }
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

function insertWordCalloutBox() {
  const html = `<div style="background:#e8eef7; border-left:4px solid #2b579a; padding:12px 16px; margin:14px 0;"><strong>💡 핵심 요약:</strong> 여기에 강조할 내용을 작성하세요.</div><p><br></p>`;
  execWordCmd('insertHTML', html);
}

function loadWordTemplate(type, silent = false) {
  const editor = document.getElementById('wordPageEditor');
  const today = new Date().toISOString().slice(0, 10);

  if (type === 'report') {
    editor.innerHTML = `
      <h1>2026년 프로젝트 추진 기획 및 업무 보고서</h1>
      <p style="color:#475569;"><strong>작성일자:</strong> ${today} &nbsp;|&nbsp; <strong>작성부서:</strong> 전략기획팀 &nbsp;|&nbsp; <strong>문서등급:</strong> 대외비</p>
      <hr style="border:none; border-top:2px solid #2b579a; margin:14px 0 20px;" />

      <h2>1. 추진 배경 및 목적</h2>
      <p>상용 MS 365 오피스 정기 구독료를 100% 절감하고, 어느 노트북에서나 익숙한 MS 오피스 리본 UI 그대로 엑셀·워드·파워포인트 문서를 편집·저장할 수 있는 환경을 구축합니다.</p>

      <h2>2. 주요 실행 계획 및 비용 비교</h2>
      <table>
        <thead>
          <tr>
            <th>구분</th>
            <th>기존 방식 (MS 365 유료 구독)</th>
            <th>신규 도입 (MS 스타일 웹 오피스)</th>
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
            <td><strong>화면 구성</strong></td>
            <td>MS Office 리본 메뉴</td>
            <td><strong>동일한 MS Office 리본 UI 적용</strong></td>
            <td>적응 시간 0초</td>
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
  if (!silent) showToast('Word 문서 서식이 적용되었습니다.');
}

async function importDocxFile(file) {
  if (typeof mammoth === 'undefined') {
    showToast('DOCX 파서를 불러오는 중입니다. 1초 후 다시 시도해 주세요.');
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
  showToast(`'${title}.doc' 저장 완료!`);
}

function printWordDocument() {
  window.print();
}

/* ============================================================================
 * 3. MS POWERPOINT ENGINE
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
        text: '제목을 추가하려면 클릭하세요',
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
        text: '부제목을 추가하려면 클릭하세요',
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
            w: 170,
            h: 36,
            text: '2026 PRESENTATION',
            fontSize: 13,
            color: '#ffffff',
            bg: '#b7472a',
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
            text: '비즈니스 전략 및 도입 제안서',
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
            text: '익숙한 MS Office 365 리본 인터페이스 기반 초경량 웹 오피스',
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
            text: '⚡ 익숙한 리본 UI\n\n실제 MS 오피스와 동일한 홈/삽입/수식 리본 메뉴를 제공합니다.',
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
              ['Excel', '수식 계산(SUM/AVERAGE), 다중 시트, 셀 서식', '표준 .xlsx / .csv'],
              ['Word', 'A4 페이지 편집, 눈금자, 표/이미지 삽입', 'MS 호환 .doc / PDF'],
              ['PowerPoint', '16:9 슬라이드 편집, 발표 모드(F5)', '표준 PowerPoint .pptx']
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
            text: '• 핵심 과제 1: MS 오피스 365 정통 리본 UI 적용 완료\n\n• 핵심 과제 2: 신규 고객사 제안서 작성 완료',
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
  if (!silent) showToast('PowerPoint 템플릿이 적용되었습니다.');
}

function savePptState() {
  localStorage.setItem(STORAGE_KEYS.PPT_DECK, JSON.stringify(pptDeck));
  markAutoSaved();
}

function renderPptWorkspace() {
  renderPptThumbnails();
  renderCurrentSlideStage();
  renderPptInspector();

  if (currentMode === 'ppt') {
    document.getElementById('statusReadyText').textContent =
      `슬라이드 ${currentSlideIndex + 1} / ${pptDeck.length}`;
    document.getElementById('statusContextInfo').textContent = '한국어 (대한민국)';
    document.getElementById('statusMetricsText').innerHTML =
      `개체 수: ${pptDeck[currentSlideIndex]?.elements.length || 0}개 &nbsp;&nbsp;&nbsp; 16:9 와이드스크린`;
  }
}

function renderPptThumbnails() {
  const listEl = document.getElementById('pptThumbnailList');
  if (!listEl) return;
  listEl.innerHTML = '';

  pptDeck.forEach((slide, idx) => {
    const row = document.createElement('div');
    row.className = `slide-thumb-row ${idx === currentSlideIndex ? 'active' : ''}`;
    row.onclick = () => {
      currentSlideIndex = idx;
      selectedElementId = null;
      renderPptWorkspace();
    };

    const firstText = slide.elements.find((e) => e.text)?.text || `슬라이드 ${idx + 1}`;
    const isDark = slide.bg === '#0f172a';

    row.innerHTML = `
      <span class="slide-num-label">${idx + 1}</span>
      <div class="slide-thumb-card" style="background:${slide.bg}; color:${isDark ? '#fff' : '#0f172a'};">
        <div style="font-weight:700; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
          ${escapeHtml(firstText.split('\n')[0])}
        </div>
      </div>
    `;
    listEl.appendChild(row);
  });
}

function renderCurrentSlideStage() {
  const stage = document.getElementById('pptSlideStage');
  const slide = pptDeck[currentSlideIndex];
  if (!stage || !slide) return;

  stage.style.background = slide.bg || '#ffffff';
  stage.innerHTML = '';

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
      item.style.borderRadius = '8px';
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
    body.innerHTML = `<p class="inspector-empty">슬라이드의 텍스트나 도형을 클릭하면 크기·색상·정렬을 정밀하게 편집할 수 있습니다.</p>`;
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
        <input type="color" value="${el.color || '#0f172a'}" onchange="updateSelectedPptProp('color', this.value)" style="height:30px;" />
      </div>
      <div class="prop-row">
        <label>도형 채우기</label>
        <input type="color" value="${el.bg && el.bg !== 'transparent' ? el.bg : '#ffffff'}" onchange="updateSelectedPptProp('bg', this.value)" style="height:30px;" />
      </div>
    </div>
    <div style="display:flex; gap:6px; margin-top:8px;">
      <button class="r-wide-btn" style="flex:1;" onclick="toggleSelectedPptBold()"><b>가</b> 굵게</button>
      <button class="r-wide-btn" style="flex:1;color:#a4262c;" onclick="deleteSelectedPptElement()">🗑️ 삭제</button>
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

function nudgeSelectedPptFontSize(delta) {
  const slide = pptDeck[currentSlideIndex];
  const el = slide?.elements.find((item) => item.id === selectedElementId);
  if (!el) return;
  el.fontSize = Math.max(10, (el.fontSize || 20) + delta);
  savePptState();
  renderCurrentSlideStage();
  renderPptInspector();
}

function toggleSelectedPptBold() {
  const slide = pptDeck[currentSlideIndex];
  const el = slide?.elements.find((item) => item.id === selectedElementId);
  if (!el) return;
  el.bold = !el.bold;
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
      text: '텍스트를 입력하세요',
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
      text: '도형 / 카드 텍스트',
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
      item.style.borderRadius = '8px';
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
    showToast(`'${fileName}' 저장 완료!`);
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
