/* ============================================================================
 * NEWT Web Office — Excel (.xlsx) · Word (.docx/.doc) · PowerPoint (.pptx)
 * 100% Client-Side Serverless Office Suite
 * ============================================================================ */

let currentMode = 'excel'; // 'excel' | 'word' | 'ppt'
let luckysheetInitialized = false;

// PPT State
let pptDeck = [];
let currentSlideIndex = 0;
let selectedElementId = null;
let dragContext = null;
let presentationIndex = 0;

// LocalStorage Keys
const STORAGE_KEYS = {
  TITLE: 'newt_office_doc_title',
  WORD_HTML: 'newt_office_word_html',
  PPT_DECK: 'newt_office_ppt_deck',
  EXCEL_SHEETS: 'newt_office_excel_sheets'
};

/* ==================== INITIALIZATION ==================== */
window.addEventListener('DOMContentLoaded', () => {
  if (window.lucide) lucide.createIcons();

  // Restore document title
  const savedTitle = localStorage.getItem(STORAGE_KEYS.TITLE);
  if (savedTitle) {
    document.getElementById('docTitleInput').value = savedTitle;
  }
  document.getElementById('docTitleInput').addEventListener('input', (e) => {
    localStorage.setItem(STORAGE_KEYS.TITLE, e.target.value.trim() || '제목 없는 문서');
    markAutoSaved();
  });

  // File Input Listener
  document.getElementById('globalFileInput').addEventListener('change', handleGlobalFileOpen);

  // Initialize 3 Engines
  initExcelEngine();
  initWordEngine();
  initPptEngine();

  // Keyboard shortcuts (F5 for SlideShow, Delete in PPT)
  window.addEventListener('keydown', (e) => {
    if (currentMode === 'ppt') {
      if (e.key === 'F5') {
        e.preventDefault();
        startSlideShow();
      } else if (e.key === 'Escape') {
        const modal = document.getElementById('pptPresentationModal');
        if (!modal.classList.contains('hidden')) {
          closeSlideShow();
        }
      } else if (!document.getElementById('pptPresentationModal').classList.contains('hidden')) {
        if (e.key === 'ArrowRight' || e.key === ' ') stepPresentation(1);
        if (e.key === 'ArrowLeft') stepPresentation(-1);
      }
    }
  });
});

function showToast(msg) {
  const toast = document.getElementById('officeToast');
  toast.textContent = msg;
  toast.classList.remove('hidden');
  clearTimeout(window.__toastTimer);
  window.__toastTimer = setTimeout(() => {
    toast.classList.add('hidden');
  }, 2800);
}

function markAutoSaved() {
  const badge = document.getElementById('autoSaveBadge');
  const now = new Date();
  const timeStr = now.toTimeString().slice(0, 5);
  badge.textContent = `자동 저장됨 (${timeStr})`;
}

function getCleanDocTitle() {
  const val = document.getElementById('docTitleInput').value.trim();
  return val || 'NEWT_문서';
}

/* ==================== MODE SWITCHING ==================== */
function switchOfficeMode(mode) {
  currentMode = mode;

  // Update Tabs
  document.querySelectorAll('.suite-tab').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.mode === mode);
  });

  // Update Views
  document.querySelectorAll('.office-view').forEach((sec) => {
    sec.classList.remove('active');
  });
  document.getElementById(`view-${mode}`).classList.add('active');

  // Update Header Brand & Primary Download Button
  const badge = document.getElementById('brandBadge');
  const dlBtn = document.getElementById('primaryDownloadBtn');
  const dlLabel = document.getElementById('downloadBtnLabel');

  if (mode === 'excel') {
    badge.style.background = 'var(--excel-primary)';
    badge.textContent = 'X';
    dlBtn.style.background = 'var(--excel-primary)';
    dlLabel.textContent = '.xlsx 다운로드';
    setTimeout(() => {
      if (window.luckysheet && typeof luckysheet.resize === 'function') {
        luckysheet.resize();
      }
    }, 80);
  } else if (mode === 'word') {
    badge.style.background = 'var(--word-primary)';
    badge.textContent = 'W';
    dlBtn.style.background = 'var(--word-primary)';
    dlLabel.textContent = '.doc 다운로드';
  } else if (mode === 'ppt') {
    badge.style.background = 'var(--ppt-primary)';
    badge.textContent = 'P';
    dlBtn.style.background = 'var(--ppt-primary)';
    dlLabel.textContent = '.pptx 다운로드';
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
    if (!confirm('현재 시트를 초기화하고 새 빈 엑셀 통합 문서를 만드시겠습니까?')) return;
    createLuckysheetInstance([
      {
        name: 'Sheet1',
        color: '#107c41',
        status: 1,
        order: 0,
        celldata: [],
        row: 60,
        column: 26
      }
    ]);
    showToast('새 엑셀 시트가 준비되었습니다.');
  } else if (currentMode === 'word') {
    if (!confirm('현재 워드 문서를 비우고 새로 작성하시겠습니까?')) return;
    const editor = document.getElementById('wordPageEditor');
    editor.innerHTML = '<h1>새 문서 제목</h1><p>여기에 내용을 입력하세요...</p>';
    saveWordState();
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

/* ==================== GLOBAL FILE OPEN (.xlsx, .csv, .docx, .pptx, etc.) ==================== */
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
      showToast(`'${file.name}' 문서를 불러왔습니다.`);
    } else if (['pptx'].includes(ext)) {
      switchOfficeMode('ppt');
      await importPptxFile(file);
    } else if (['json'].includes(ext)) {
      const jsonData = JSON.parse(await file.text());
      if (Array.isArray(jsonData) && jsonData[0] && jsonData[0].elements) {
        switchOfficeMode('ppt');
        pptDeck = jsonData;
        currentSlideIndex = 0;
        renderPptWorkspace();
        showToast('PPT 슬라이드 데이터를 불러왔습니다.');
      }
    } else {
      showToast('지원하지 않는 파일 형식입니다.');
    }
  } catch (err) {
    console.error(err);
    showToast('파일을 여는 중 오류가 발생했습니다: ' + err.message);
  }
}

/* ============================================================================
 * 1. EXCEL ENGINE (Luckysheet + SheetJS .xlsx / .csv)
 * ============================================================================ */
function initExcelEngine() {
  loadExcelSample('budget', true);
}

function createLuckysheetInstance(sheetsData) {
  if (window.luckysheet && typeof luckysheet.destroy === 'function' && luckysheetInitialized) {
    try {
      luckysheet.destroy();
    } catch (e) {}
  }

  luckysheet.create({
    container: 'luckysheet-container',
    lang: 'ko',
    showinfobar: false,
    showsheetbar: true,
    showstatisticBar: true,
    enableAddRow: true,
    enableAddBackTop: true,
    data: sheetsData
  });
  luckysheetInitialized = true;
}

function aoaToCelldata(aoa, headerBg = '#107c41') {
  const celldata = [];
  for (let r = 0; r < aoa.length; r++) {
    for (let c = 0; c < aoa[r].length; c++) {
      const raw = aoa[r][c];
      if (raw === null || raw === undefined || raw === '') continue;

      const cellObj = {
        r,
        c,
        v: {}
      };

      if (typeof raw === 'string' && raw.startsWith('=')) {
        cellObj.v.f = raw;
        cellObj.v.v = '';
      } else if (typeof raw === 'number') {
        cellObj.v.v = raw;
        cellObj.v.m = raw.toLocaleString('ko-KR');
        cellObj.v.ct = { fa: '#,##0', t: 'n' };
      } else {
        cellObj.v.v = String(raw);
        cellObj.v.m = String(raw);
      }

      if (r === 0) {
        cellObj.v.bg = headerBg;
        cellObj.v.fc = '#ffffff';
        cellObj.v.bl = 1;
        cellObj.v.ht = 0;
      }

      celldata.push(cellObj);
    }
  }
  return celldata;
}

function loadExcelSample(type, silent = false) {
  let aoa = [];
  let sheetName = 'Sheet1';

  if (type === 'budget') {
    sheetName = '월간예산_가계부';
    aoa = [
      ['항목 구분', '세부 내역', '예산 (원)', '실제 지출 (원)', '차액 (예산-실출)', '비고'],
      ['고정지출', '통신비 및 인터넷', 85000, 82000, '=C2-D2', '자동이체'],
      ['고정지출', '구독료 (MS 오피스 해지!)', 11900, 0, '=C3-D3', 'NEWT 웹오피스로 전환 완료'],
      ['생활비', '식비 / 장보기', 600000, 540000, '=C4-D4', '카드 결제'],
      ['생활비', '교통 / 주유비', 200000, 185000, '=C5-D5', ''],
      ['저축/투자', 'ETF 정기 적립', 1000000, 1000000, '=C6-D6', '완료'],
      ['합계', '총계', '=SUM(C2:C6)', '=SUM(D2:D6)', '=SUM(E2:E6)', '자동 수식 합계']
    ];
  } else if (type === 'invoice') {
    sheetName = '거래명세서';
    aoa = [
      ['품목코드', '품목명 / 규격', '수량', '단가 (원)', '공급가액', '부가세 (10%)', '합계금액'],
      ['ITM-001', '고성능 사무용 노트북 16GB', 2, 1250000, '=C2*D2', '=E2*0.1', '=E2+F2'],
      ['ITM-002', '27인치 4K 모니터', 2, 420000, '=C3*D3', '=E3*0.1', '=E3+F3'],
      ['ITM-003', '무선 키보드/마우스 세트', 4, 65000, '=C4*D4', '=E4*0.1', '=E4+F4'],
      ['합계', '총 청구 금액', '=SUM(C2:C4)', '', '=SUM(E2:E4)', '=SUM(F2:F4)', '=SUM(G2:G4)']
    ];
  } else if (type === 'calc') {
    sheetName = '수익률_시뮬레이션';
    aoa = [
      ['연차', '기초 자산 (원)', '연간 추가납입', '목표 수익률(%)', '연말 평가금액'],
      [1, 10000000, 6000000, 12, '=(B2+C2)*1.12'],
      [2, '=E2', 6000000, 12, '=(B3+C3)*1.12'],
      [3, '=E3', 6000000, 12, '=(B4+C4)*1.12'],
      [4, '=E4', 6000000, 12, '=(B5+C5)*1.12'],
      [5, '=E5', 6000000, 12, '=(B6+C6)*1.12']
    ];
  }

  createLuckysheetInstance([
    {
      name: sheetName,
      color: '#107c41',
      status: 1,
      order: 0,
      celldata: aoaToCelldata(aoa),
      row: 60,
      column: 26,
      config: {
        columnlen: { 0: 110, 1: 210, 2: 130, 3: 130, 4: 140, 5: 180, 6: 140 }
      }
    }
  ]);

  if (!silent) showToast(`'${sheetName}' 샘플 시트를 불러왔습니다.`);
}

async function importExcelFile(file) {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array', cellFormula: true });

  const luckySheets = wb.SheetNames.map((name, idx) => {
    const ws = wb.Sheets[name];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
    const celldata = [];

    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = ws[addr];
        if (!cell) continue;
        const vObj = {
          v: cell.v !== undefined ? cell.v : '',
          m: cell.w !== undefined ? String(cell.w) : String(cell.v !== undefined ? cell.v : '')
        };
        if (cell.f) {
          vObj.f = '=' + cell.f;
        }
        celldata.push({ r: R, c: C, v: vObj });
      }
    }

    return {
      name,
      status: idx === 0 ? 1 : 0,
      order: idx,
      celldata,
      row: Math.max(60, aoa.length + 20),
      column: Math.max(26, (aoa[0]?.length || 10) + 10)
    };
  });

  createLuckysheetInstance(luckySheets);
  showToast(`엑셀 파일 '${file.name}' (${luckySheets.length}개 시트) 열기 완료!`);
}

function convertLuckysheetToWorkbook() {
  const sheets = luckysheet.getAllSheets();
  const wb = XLSX.utils.book_new();

  sheets.forEach((sheet, idx) => {
    const data = sheet.data || [];
    let maxR = 0;
    let maxC = 0;

    for (let r = 0; r < data.length; r++) {
      if (!data[r]) continue;
      for (let c = 0; c < data[r].length; c++) {
        const cell = data[r][c];
        if (cell && (cell.v !== undefined && cell.v !== null && cell.v !== '' || cell.f)) {
          if (r > maxR) maxR = r;
          if (c > maxC) maxC = c;
        }
      }
    }

    const aoa = [];
    for (let r = 0; r <= maxR; r++) {
      const rowArr = [];
      for (let c = 0; c <= maxC; c++) {
        const cell = data[r] ? data[r][c] : null;
        if (!cell) {
          rowArr.push('');
        } else {
          rowArr.push(cell.v !== undefined && cell.v !== null ? cell.v : (cell.m || ''));
        }
      }
      aoa.push(rowArr);
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Preserve formulas if present
    for (let r = 0; r <= maxR; r++) {
      for (let c = 0; c <= maxC; c++) {
        const cell = data[r] ? data[r][c] : null;
        if (cell && cell.f) {
          const addr = XLSX.utils.encode_cell({ r, c });
          if (!ws[addr]) ws[addr] = { t: 'n', v: Number(cell.v) || 0 };
          ws[addr].f = cell.f.replace(/^=/, '');
        }
      }
    }

    XLSX.utils.book_append_sheet(wb, ws, sheet.name || `Sheet${idx + 1}`);
  });

  return wb;
}

function exportExcelXLSX() {
  const wb = convertLuckysheetToWorkbook();
  const fileName = `${getCleanDocTitle()}.xlsx`;
  XLSX.writeFile(wb, fileName);
  showToast(`'${fileName}' 엑셀 파일 다운로드 완료!`);
}

function exportExcelCSV() {
  const wb = convertLuckysheetToWorkbook();
  const firstSheet = wb.Sheets[wb.SheetNames[0]];
  const csv = XLSX.utils.sheet_to_csv(firstSheet);
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
  triggerBlobDownload(blob, `${getCleanDocTitle()}.csv`);
  showToast('CSV 파일 다운로드 완료!');
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
  const sel = window.getSelection();
  if (!sel.rangeCount) return;
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
  for (let c = 0; c < cols; c++) {
    html += `<th>항목 ${c + 1}</th>`;
  }
  html += '</tr></thead><tbody>';
  for (let r = 0; r < rows - 1; r++) {
    html += '<tr>';
    for (let c = 0; c < cols; c++) {
      html += '<td>내용 입력</td>';
    }
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
      const imgHtml = `<p><img src="${reader.result}" style="max-width:100%; height:auto; border-radius:4px;" /></p>`;
      execWordCmd('insertHTML', imgHtml);
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
      <p style="color:#475569;"><strong>작성일자:</strong> ${today} &nbsp;|&nbsp; <strong>작성자:</strong> 기획팀 &nbsp;|&nbsp; <strong>문서등급:</strong> 대외비</p>
      <hr style="border:none; border-top:2px solid #1e3a8a; margin:14px 0 20px;" />

      <h2>1. 추진 배경 및 목적</h2>
      <p>본 문서는 상용 오피스 구독 비용을 절감하고 어느 PC에서나 웹 브라우저만으로 엑셀·워드·파워포인트 문서를 즉시 열람·편집·저장할 수 있는 클라우드 독립형 오피스 환경을 구축하기 위해 작성되었습니다.</p>

      <h2>2. 주요 실행 계획 및 예산 비교</h2>
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
            <td><strong>연간 비용</strong></td>
            <td>매년 정기 구독료 발생</td>
            <td><strong>0원 (100% 영구 무료)</strong></td>
            <td>100% 비용 절감</td>
          </tr>
          <tr>
            <td><strong>설치 여부</strong></td>
            <td>노트북마다 무거운 설치 필요</td>
            <td><strong>설치 없이 웹 주소 접속 즉시 사용</strong></td>
            <td>관리 편의성 극대화</td>
          </tr>
          <tr>
            <td><strong>호환 포맷</strong></td>
            <td>.xlsx, .docx, .pptx</td>
            <td><strong>.xlsx, .docx/.doc, .pptx 완벽 호환</strong></td>
            <td>업무 연속성 유지</td>
          </tr>
        </tbody>
      </table>

      <h2>3. 향후 추진 일정</h2>
      <ul>
        <li><strong>1단계:</strong> 기존 노트북의 엑셀(.xlsx), 워드(.docx), 피피티(.pptx) 파일 열기 및 호환성 점검</li>
        <li><strong>2단계:</strong> 불필요한 MS 오피스 정기 구독 해지 완료</li>
        <li><strong>3단계:</strong> 외부 노트북 및 태블릿에서 북마크 하나로 즉시 업무 활용</li>
      </ul>
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
            <td style="width:30%;">대회의실 / 화상회의</td>
          </tr>
          <tr>
            <th>참석자</th>
            <td colspan="3">홍길동 팀장, 김철수 선임, 이영희 매니저</td>
          </tr>
          <tr>
            <th>회의 안건</th>
            <td colspan="3"><strong>2026년 신규 사업 마일스톤 점검 및 리소스 배분</strong></td>
          </tr>
        </tbody>
      </table>

      <h2>1. 주요 논의 사항</h2>
      <ul>
        <li>핵심 과제 A: 일정 차질 없이 진행 중 (진척률 85%)</li>
        <li>핵심 과제 B: 외부 협력사 검토 의견 반영하여 다음 주 월요일 최종본 배포 예정</li>
      </ul>

      <h2>2. 결정 사항 및 Action Item</h2>
      <table>
        <thead>
          <tr>
            <th>담당자</th>
            <th>실행 과제 (Action Item)</th>
            <th>완료 기한</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>홍길동</td>
            <td>최종 예산안 엑셀 시트 확정 및 공유</td>
            <td>이번 주 금요일</td>
          </tr>
          <tr>
            <td>이영희</td>
            <td>고객사 제안용 PPT 슬라이드 업데이트</td>
            <td>다음 주 화요일</td>
          </tr>
        </tbody>
      </table>
    `;
  } else if (type === 'contract') {
    editor.innerHTML = `
      <h1 style="text-align:center;">업무 협약 및 표준 합의서</h1>
      <p style="text-align:center; color:#64748b;">문서번호: 제 2026-001 호</p>
      <p><br></p>
      <p><strong>갑 (발주자):</strong> ________________________ (인)</p>
      <p><strong>을 (수행자):</strong> ________________________ (인)</p>
      <p><br></p>
      <p>상기 당사자는 상호 신뢰를 바탕으로 다음과 같이 업무 협약 및 계약을 체결한다.</p>

      <h2>제 1 조 (목적)</h2>
      <p>본 합의서는 '갑'과 '을' 간의 프로젝트 수행 범위, 납품 조건 및 상호 권리·의무를 명확히 규정함을 목적으로 한다.</p>

      <h2>제 2 조 (계약 기간 및 금액)</h2>
      <ul>
        <li><strong>계약 기간:</strong> ${today} 부터 2026년 12월 31일 까지</li>
        <li><strong>계약 금액:</strong> 일금 __________________ 원정 (₩ ____________ / 부가세 별도)</li>
      </ul>

      <h2>제 3 조 (특약 사항)</h2>
      <p>양 당사자는 본 계약의 증명으로 본 문서 2부를 작성하여 각각 서명 날인 후 1부씩 보관한다.</p>
      <p><br></p>
      <p style="text-align:center;"><strong>${today}</strong></p>
    `;
  }

  saveWordState();
  updateWordCounters();
  if (!silent) showToast('문서 양식을 적용했습니다.');
}

async function importDocxFile(file) {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer });
  const editor = document.getElementById('wordPageEditor');
  editor.innerHTML = result.value || '<p>(빈 문서)</p>';
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
        h1 { font-size: 20pt; color: #0f172a; }
        h2 { font-size: 14pt; color: #1e3a8a; border-bottom: 1px solid #cbd5e1; padding-bottom: 4px; }
        table { width: 100%; border-collapse: collapse; margin: 12px 0; }
        th, td { border: 1px solid #64748b; padding: 6px 8px; }
        th { background-color: #f1f5f9; font-weight: bold; }
      </style>
    </head>
    <body>
      ${content}
    </body>
    </html>
  `;

  const blob = new Blob(['\uFEFF', fullWordHtml], {
    type: 'application/msword;charset=utf-8'
  });
  triggerBlobDownload(blob, `${title}.doc`);
  showToast(`'${title}.doc' MS 워드 호환 파일 다운로드 완료!`);
}

function printWordDocument() {
  window.print();
}

/* ============================================================================
 * 3. POWERPOINT ENGINE (16:9 Visual Slide Editor + PptxGenJS .pptx Export & Import)
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
    notes: '발표 시작 인사 및 핵심 주제 소개',
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
        text: '부제목 또는 발표자 정보를 여기에 작성합니다 (더블클릭하여 편집)',
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
        notes: '표지 슬라이드입니다.',
        elements: [
          {
            id: 'e1',
            type: 'box',
            x: 70,
            y: 120,
            w: 150,
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
            text: 'MS 오피스 구독 없이 모든 노트북에서 즉시 실행되는 차세대 웹 오피스',
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
        notes: '핵심 장점 3가지를 설명합니다.',
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
            text: '⚡ 무설치 즉시 사용\n\n새 노트북에서도 설치 없이 웹 주소만 열면 바로 작업합니다.',
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
        notes: '일정 및 비용 비교 표 슬라이드입니다.',
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
              ['엑셀 (Spreadsheet)', '수식 계산(SUM/IF), 다중 시트, 셀 서식', '표준 .xlsx / .csv'],
              ['워드 (Document)', 'A4 페이지 편집, 표/이미지 삽입, 양식 제공', 'MS 호환 .doc / PDF 인쇄'],
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
        notes: '주간 보고 표지',
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
          },
          {
            id: 'we2',
            type: 'text',
            x: 70,
            y: 285,
            w: 700,
            h: 50,
            text: '금주 핵심 성과 요약 및 이슈 점검',
            fontSize: 20,
            color: '#475569',
            bg: 'transparent',
            bold: false,
            align: 'left'
          }
        ]
      },
      {
        id: 'w2',
        bg: '#ffffff',
        notes: '금주 실적 상세',
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
            text: '• 핵심 과제 1: 시스템 안정화 및 성능 개선 완료 (100%)\n\n• 핵심 과제 2: 신규 고객사 제안서 작성 및 발송 완료\n\n• 핵심 과제 3: 내부 운영 비용 절감 프로세스 적용',
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
  if (!silent) showToast('PPT 템플릿 세트를 불러왔습니다.');
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
  if (!slide) return;

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

    // Mouse down for selection & drag
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
        if (document.activeElement === editable) {
          e.stopPropagation();
        }
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
    const dx = moveEvt.clientX - startX;
    const dy = moveEvt.clientY - startY;
    el.x = Math.max(0, Math.min(920, origX + dx));
    el.y = Math.max(0, Math.min(500, origY + dy));
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
  const slide = pptDeck[currentSlideIndex];
  const el = slide?.elements.find((item) => item.id === selectedElementId);

  if (!el) {
    body.innerHTML = `<p class="inspector-empty">슬라이드의 텍스트나 도형을 클릭하면 크기·색상·정렬을 조절할 수 있습니다.<br><br>• <strong>이동:</strong> 마우스로 드래그<br>• <strong>글자 수정:</strong> 텍스트 클릭 후 타이핑</p>`;
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
        <input type="color" value="${el.color || '#0f172a'}" onchange="updateSelectedPptProp('color', this.value)" style="height:34px;" />
      </div>
      <div class="prop-row">
        <label>배경 색상</label>
        <input type="color" value="${el.bg && el.bg !== 'transparent' ? el.bg : '#ffffff'}" onchange="updateSelectedPptProp('bg', this.value)" style="height:34px;" />
      </div>
    </div>

    <div style="display:flex; gap:8px; margin-top:8px;">
      <button class="tool-btn" style="flex:1;" onclick="updateSelectedPptProp('bold', !${Boolean(el.bold)})">
        <b>B</b> 굵게 토글
      </button>
      <button class="tool-btn danger-hover" style="flex:1;" onclick="deleteSelectedPptElement()">
        🗑️ 요소 삭제
      </button>
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
  const newSlide = {
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
        text: '• 핵심 내용을 여기에 작성하세요.\n\n• 상단 툴바에서 표, 도형, 이미지를 추가할 수 있습니다.',
        fontSize: 20,
        color: '#334155',
        bg: '#f8fafc',
        bold: false,
        align: 'left'
      }
    ]
  };
  pptDeck.splice(currentSlideIndex + 1, 0, newSlide);
  currentSlideIndex++;
  selectedElementId = null;
  savePptState();
  renderPptWorkspace();
  showToast('새 슬라이드가 추가되었습니다.');
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
  showToast('슬라이드가 복제되었습니다.');
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
  showToast('슬라이드가 삭제되었습니다.');
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
      text: '새 텍스트 상자 (클릭하여 입력)',
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

/* ==================== FULLSCREEN SLIDESHOW (F5) ==================== */
function startSlideShow() {
  presentationIndex = currentSlideIndex;
  const modal = document.getElementById('pptPresentationModal');
  modal.classList.remove('hidden');
  if (modal.requestFullscreen) {
    modal.requestFullscreen().catch(() => {});
  }
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
  const indicator = document.getElementById('presentationPageIndicator');
  indicator.textContent = `${presentationIndex + 1} / ${pptDeck.length}`;

  const slide = pptDeck[presentationIndex];
  const scale = Math.min(window.innerWidth / 960, window.innerHeight / 540) * 0.94;

  const stageClone = document.createElement('div');
  stageClone.className = 'ppt-slide-stage';
  stageClone.style.background = slide.bg || '#ffffff';
  stageClone.style.transform = `scale(${scale})`;
  stageClone.style.transformOrigin = 'center center';

  slide.elements.forEach((el) => {
    const item = document.createElement('div');
    item.className = 'ppt-item';
    item.style.cursor = 'default';
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

/* ==================== EXPORT & IMPORT REAL .PPTX ==================== */
function exportPptxFile() {
  if (typeof PptxGenJS === 'undefined') {
    showToast('PPTX 생성 엔진을 불러오는 중입니다.');
    return;
  }

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9'; // 10 x 5.625 inches

  // Stage is 960 x 540 px => 1 inch = 96 px
  const pxToIn = (px) => Number((px / 96).toFixed(3));

  pptDeck.forEach((s) => {
    const slide = pptx.addSlide();
    if (s.bg && s.bg !== 'transparent') {
      slide.background = { color: s.bg.replace('#', '') };
    }
    if (s.notes) {
      slide.addNotes(s.notes);
    }

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
        slide.addImage({
          data: el.src,
          x,
          y,
          w,
          h
        });
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
    showToast(`'${fileName}' 파워포인트(.pptx) 저장 완료!`);
  });
}

async function importPptxFile(file) {
  const zip = await JSZip.loadAsync(file);
  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)\.xml/)[1], 10);
      const nb = parseInt(b.match(/slide(\d+)\.xml/)[1], 10);
      return na - nb;
    });

  if (slideFiles.length === 0) {
    showToast('PPTX 내부에서 슬라이드를 찾지 못했습니다.');
    return;
  }

  const importedDeck = [];

  for (let i = 0; i < slideFiles.length; i++) {
    const xmlStr = await zip.files[slideFiles[i]].async('string');
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');

    // Extract text paragraphs per shape (<p:sp>)
    const shapes = Array.from(xmlDoc.getElementsByTagName('p:sp'));
    const elements = [];
    let autoY = 60;

    shapes.forEach((sp, spIdx) => {
      const paragraphs = Array.from(sp.getElementsByTagName('a:p'));
      const lines = paragraphs
        .map((p) =>
          Array.from(p.getElementsByTagName('a:t'))
            .map((t) => t.textContent)
            .join('')
        )
        .filter((line) => line.trim().length > 0);

      if (lines.length === 0) return;

      const fullText = lines.join('\n');
      const isTitle = spIdx === 0;

      elements.push({
        id: `imp_${i}_${spIdx}`,
        type: isTitle ? 'text' : 'box',
        x: 60,
        y: autoY,
        w: 840,
        h: isTitle ? 70 : Math.min(300, Math.max(80, lines.length * 42)),
        text: fullText,
        fontSize: isTitle ? 32 : 20,
        color: '#0f172a',
        bg: isTitle ? 'transparent' : '#f8fafc',
        bold: isTitle,
        align: 'left'
      });

      autoY += isTitle ? 90 : Math.min(310, Math.max(95, lines.length * 45));
    });

    if (elements.length === 0) {
      elements.push({
        id: `imp_empty_${i}`,
        type: 'text',
        x: 70,
        y: 200,
        w: 800,
        h: 70,
        text: `슬라이드 ${i + 1}`,
        fontSize: 30,
        color: '#0f172a',
        bg: 'transparent',
        bold: true,
        align: 'left'
      });
    }

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
  showToast(`PPTX 파일 '${file.name}' (${importedDeck.length}장) 불러오기 완료!`);
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
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
