const state = {
  rawRows: [],
  rows: [],
  analysisCacheValid: false,
  analysisCacheKey: "",
  analysisRowsCache: [],
  headers: [],
  fileName: "",
  parsedRows: [],
  filteredRows: [],
  scoreBounds: {
    minGpa: 0,
    maxGpa: 4.5,
    minToeic: 0,
    maxToeic: 990,
    minMockToeic: 0,
    maxMockToeic: 990,
  },
  studentIdColumn: "",
  departmentColumn: "",
  gradeColumn: "",
  semesterColumn: "",
  gpaColumn: "",
  toeicColumnA: "",
  toeicDateColumnA: "",
  toeicColumnB: "",
  toeicDateColumnB: "",
  mockToeicColumnA: "",
  mockToeicDateColumnA: "",
  mockToeicColumnB: "",
  mockToeicDateColumnB: "",
};

const els = {
  fileInput: document.querySelector("#fileInput"),
  uploadPanel: document.querySelector("#uploadPanel"),
  mappingPanel: document.querySelector("#mappingPanel"),
  mappingHelp: document.querySelector("#mappingHelp"),
  studentIdColumn: document.querySelector("#studentIdColumn"),
  departmentColumn: document.querySelector("#departmentColumn"),
  gradeColumn: document.querySelector("#gradeColumn"),
  semesterColumn: document.querySelector("#semesterColumn"),
  gpaColumn: document.querySelector("#gpaColumn"),
  toeicColumnA: document.querySelector("#toeicColumnA"),
  toeicDateColumnA: document.querySelector("#toeicDateColumnA"),
  toeicColumnB: document.querySelector("#toeicColumnB"),
  toeicDateColumnB: document.querySelector("#toeicDateColumnB"),
  mockToeicColumnA: document.querySelector("#mockToeicColumnA"),
  mockToeicDateColumnA: document.querySelector("#mockToeicDateColumnA"),
  mockToeicColumnB: document.querySelector("#mockToeicColumnB"),
  mockToeicDateColumnB: document.querySelector("#mockToeicDateColumnB"),
  applyMapping: document.querySelector("#applyMapping"),
  dashboard: document.querySelector("#dashboard"),
  dataSummary: document.querySelector("#dataSummary"),
  departmentFilter: document.querySelector("#departmentFilter"),
  gradeFilter: document.querySelector("#gradeFilter"),
  semesterFilter: document.querySelector("#semesterFilter"),
  periodHelp: document.querySelector("#periodHelp"),
  gpaThreshold: document.querySelector("#gpaThreshold"),
  gpaThresholdNumber: document.querySelector("#gpaThresholdNumber"),
  toeicThreshold: document.querySelector("#toeicThreshold"),
  toeicThresholdNumber: document.querySelector("#toeicThresholdNumber"),
  mockToeicThreshold: document.querySelector("#mockToeicThreshold"),
  mockToeicThresholdNumber: document.querySelector("#mockToeicThresholdNumber"),
  mockToeicHelp: document.querySelector("#mockToeicHelp"),
  targetRate: document.querySelector("#targetRate"),
  targetRateNumber: document.querySelector("#targetRateNumber"),
  targetModeHelp: document.querySelector("#targetModeHelp"),
  recommendButton: document.querySelector("#recommendButton"),
  exportButton: document.querySelector("#exportButton"),
  totalCount: document.querySelector("#totalCount"),
  totalCountDetail: document.querySelector("#totalCountDetail"),
  gpaPass: document.querySelector("#gpaPass"),
  gpaPassRate: document.querySelector("#gpaPassRate"),
  toeicPass: document.querySelector("#toeicPass"),
  toeicPassRate: document.querySelector("#toeicPassRate"),
  bothPass: document.querySelector("#bothPass"),
  bothPassRate: document.querySelector("#bothPassRate"),
  heatmap: document.querySelector("#heatmap"),
  scatterCanvas: document.querySelector("#scatterCanvas"),
  scatterGpaValue: document.querySelector("#scatterGpaValue"),
  scatterToeicValue: document.querySelector("#scatterToeicValue"),
  sensitivityCanvas: document.querySelector("#sensitivityCanvas"),
  gpaHistogram: document.querySelector("#gpaHistogram"),
  toeicHistogram: document.querySelector("#toeicHistogram"),
  gpaStats: document.querySelector("#gpaStats"),
  toeicStats: document.querySelector("#toeicStats"),
  recommendationText: document.querySelector("#recommendationText"),
  candidateList: document.querySelector("#candidateList"),
};

const numberFormat = new Intl.NumberFormat("ko-KR");
const percentFormat = new Intl.NumberFormat("ko-KR", {
  maximumFractionDigits: 1,
});
let targetRateTimer = null;
const DEFAULT_DATASET = {
  url: "./data/2026_1_students.csv",
  displayName: "2026학년도 1학기",
  sourceName: "2026_1_students.csv",
};

const gpaBins = [
  [0, 2],
  [2, 2.5],
  [2.5, 2.8],
  [2.8, 3],
  [3, 3.2],
  [3.2, 3.4],
  [3.4, 3.6],
  [3.6, 3.8],
  [3.8, 4.0],
  [4.0, 4.5],
];

const toeicBins = [
  [0, 500],
  [500, 600],
  [600, 650],
  [650, 700],
  [700, 750],
  [750, 800],
  [800, 850],
  [850, 900],
  [900, 950],
  [950, 990],
];

els.fileInput.addEventListener("click", () => {
  els.fileInput.value = "";
});
els.fileInput.addEventListener("change", handleFile);
els.applyMapping.addEventListener("click", applyMapping);
els.recommendButton.addEventListener("click", applyCurrentTarget);
els.exportButton.addEventListener("click", exportEligibleCsv);
els.departmentFilter.addEventListener("change", handleAnalysisContextChange);
els.gradeFilter.addEventListener("change", handleAnalysisContextChange);
els.semesterFilter.addEventListener("change", handleAnalysisContextChange);
document.querySelectorAll('input[name="toeicPeriod"]').forEach((input) => input.addEventListener("change", handleAnalysisContextChange));
document.querySelectorAll('input[name="targetMode"]').forEach((input) => input.addEventListener("change", handleTargetModeChange));

syncCriterionPair(els.gpaThreshold, els.gpaThresholdNumber, "gpa");
syncCriterionPair(els.toeicThreshold, els.toeicThresholdNumber, "toeic");
syncCriterionPair(els.mockToeicThreshold, els.mockToeicThresholdNumber, "mockToeic");
syncPair(els.targetRate, els.targetRateNumber, scheduleTargetAutoApply);

loadDefaultDataset();

els.toeicThreshold.step = 1;
els.toeicThresholdNumber.step = 1;
els.mockToeicThreshold.step = 1;
els.mockToeicThresholdNumber.step = 1;

function syncPair(range, number, afterChange) {
  range.addEventListener("input", () => {
    number.value = range.value;
    afterChange();
  });
  number.addEventListener("input", () => {
    range.value = number.value;
    afterChange();
  });
}

function syncCriterionPair(range, number, changedCriterion) {
  range.addEventListener("input", () => {
    number.value = range.value;
    handleCriterionChange(changedCriterion);
  });
  number.addEventListener("input", () => {
    range.value = number.value;
    handleCriterionChange(changedCriterion);
  });
}

function scheduleTargetAutoApply() {
  window.clearTimeout(targetRateTimer);
  targetRateTimer = window.setTimeout(() => {
    applyTargetRateThresholds(isAutoTargetMode());
  }, 120);
}

function handleCriterionChange(changedCriterion) {
  if (isAutoTargetMode()) {
    if (changedCriterion === "mockToeic") {
      applyTargetRateThresholds(true);
      return;
    }
    adjustComplementaryThreshold(changedCriterion);
    return;
  }

  render();
  applyTargetRateThresholds(false);
}

function handleAnalysisContextChange() {
  if (isAutoTargetMode()) {
    applyTargetRateThresholds(true);
    return;
  }

  render();
  applyTargetRateThresholds(false);
}

function handleTargetModeChange() {
  const isAuto = isAutoTargetMode();
  document.getElementById("targetRateSection").classList.toggle("hidden", !isAuto);
  if (isAuto) applyTargetRateThresholds(true);
  else render();
}

function isAutoTargetMode() {
  return document.querySelector('input[name="targetMode"]:checked')?.value !== "manual";
}

async function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  if (!window.XLSX) {
    alert("엑셀 파서가 아직 로드되지 않았습니다. 인터넷 연결을 확인한 뒤 새로고침해 주세요.");
    return;
  }

  state.fileName = file.name;
  els.dashboard.classList.add("hidden");
  els.mappingPanel.classList.remove("hidden");
  els.mappingHelp.textContent = `${file.name} 파일을 읽는 중입니다.`;
  const rows = await readSpreadsheetRows(file);

  if (!rows.length) {
    alert("분석할 행을 찾지 못했습니다.");
    return;
  }

  loadRowsForMapping(rows, file.name, { autoAnalyze: false });
}

async function loadDefaultDataset() {
  if (!window.XLSX) {
    els.mappingHelp.textContent = "엑셀 파서가 로드되지 않아 기본 분석 파일을 불러오지 못했습니다.";
    els.mappingPanel.classList.remove("hidden");
    return;
  }

  try {
    const response = await fetch(DEFAULT_DATASET.url, { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const blob = await response.blob();
    const file = new File([blob], DEFAULT_DATASET.sourceName, { type: "text/csv" });
    const rows = await readSpreadsheetRows(file);

    if (!rows.length) throw new Error("empty dataset");
    loadRowsForMapping(rows, DEFAULT_DATASET.displayName, { autoAnalyze: true });
  } catch (error) {
    console.warn("Default dataset load failed:", error);
    els.mappingHelp.textContent = "기본 분석 파일을 자동으로 불러오지 못했습니다. 엑셀/CSV 파일을 업로드해 주세요.";
    els.mappingPanel.classList.remove("hidden");
  }
}

function loadRowsForMapping(rows, displayName, { autoAnalyze = false } = {}) {
  state.fileName = displayName;
  state.rawRows = rows;
  state.headers = Object.keys(rows[0]);
  state.studentIdColumn = findLikelyColumn(state.headers, ["학번", "학생id", "studentid", "id", "studentno"], true);
  state.departmentColumn = findLikelyColumn(state.headers, ["학과", "전공", "소속", "department", "major"], true);
  state.gradeColumn = findLikelyColumn(state.headers, ["학년", "gradelevel", "year", "학기"], true);
  state.semesterColumn = findLikelyColumn(state.headers, ["인정학기", "등록학기", "이수학기", "semester", "term"], true);
  state.gpaColumn = findLikelyColumn(state.headers, ["평점", "평균평점", "학점", "gpa", "grade"]);
  state.toeicColumnA = findLikelyColumn(state.headers, ["토익a", "toeica", "토익_a", "a토익", "최근2년", "2년토익", "toeic_a", "recent"], true);
  state.toeicDateColumnA = findLikelyColumn(state.headers, ["a취득일", "취득일a", "a_취득", "최근취득", "atoeicdate"], true);
  state.toeicColumnB = findLikelyColumn(state.headers, ["토익b", "toeicb", "토익_b", "b토익", "전체토익", "toeic_b", "전기간"], true);
  state.toeicDateColumnB = findLikelyColumn(state.headers, ["b취득일", "취득일b", "b_취득", "전체취득", "btoeicdate"], true);
  state.mockToeicColumnA = findLikelyColumn(state.headers, ["모의토익a", "모의toeica", "모의토익_a", "mocktoeica", "mock_toeic_a", "mocktoeic_a"], true);
  state.mockToeicDateColumnA = findLikelyColumn(state.headers, ["모의a취득일", "모의a취득일자", "모의취득일a", "mocktoeicadate", "mock_toeic_a_date"], true);
  state.mockToeicColumnB = findLikelyColumn(state.headers, ["모의토익b", "모의toeicb", "모의토익_b", "mocktoeicb", "mock_toeic_b", "mocktoeic_b"], true);
  state.mockToeicDateColumnB = findLikelyColumn(state.headers, ["모의b취득일", "모의b취득일자", "모의취득일b", "mocktoeicbdate", "mock_toeic_b_date"], true);

  // B열이 없으면 일반 토익 열을 B로 fallback
  if (!state.toeicColumnB) {
    state.toeicColumnB = findLikelyColumn(state.headers, ["토익", "toeic", "english", "영어"]);
  }

  fillSelect(els.studentIdColumn, state.headers, state.studentIdColumn, true);
  fillSelect(els.departmentColumn, state.headers, state.departmentColumn, true);
  fillSelect(els.gradeColumn, state.headers, state.gradeColumn, true);
  fillSelect(els.semesterColumn, state.headers, state.semesterColumn, true);
  fillSelect(els.gpaColumn, state.headers, state.gpaColumn, false);
  fillSelect(els.toeicColumnA, state.headers, state.toeicColumnA, true);
  fillSelect(els.toeicDateColumnA, state.headers, state.toeicDateColumnA, true);
  fillSelect(els.toeicColumnB, state.headers, state.toeicColumnB, true);
  fillSelect(els.toeicDateColumnB, state.headers, state.toeicDateColumnB, true);
  fillSelect(els.mockToeicColumnA, state.headers, state.mockToeicColumnA, true);
  fillSelect(els.mockToeicDateColumnA, state.headers, state.mockToeicDateColumnA, true);
  fillSelect(els.mockToeicColumnB, state.headers, state.mockToeicColumnB, true);
  fillSelect(els.mockToeicDateColumnB, state.headers, state.mockToeicDateColumnB, true);
  els.mappingHelp.textContent = `${displayName} 분석 파일에서 ${numberFormat.format(rows.length)}개 행을 읽었습니다. 자동 매핑이 다르면 열을 바꿔 주세요.`;
  els.mappingPanel.classList.remove("hidden");
  if (!autoAnalyze) {
    els.uploadPanel.classList.add("hidden");
    els.dashboard.classList.add("hidden");
    els.mappingPanel.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  if (autoAnalyze) applyMapping();
}

async function readSpreadsheetRows(file) {
  const data = await file.arrayBuffer();
  const isCsv = /\.csv$/i.test(file.name) || file.type.includes("csv");
  const workbook = isCsv
    ? XLSX.read(decodeCsvText(data), { type: "string", cellDates: true })
    : XLSX.read(data, { type: "array", cellDates: true });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
}

function decodeCsvText(data) {
  const candidates = ["utf-8", "euc-kr"];
  const decoded = candidates.map((encoding) => ({
    encoding,
    text: new TextDecoder(encoding).decode(data),
  }));

  return decoded.sort((a, b) => textEncodingScore(a.text) - textEncodingScore(b.text))[0].text;
}

function textEncodingScore(text) {
  const replacementCharacters = (text.match(/\uFFFD/g) || []).length;
  const mojibakeMarkers = (text.match(/[ÃÂíêìëð]/g) || []).length;
  const koreanCharacters = (text.match(/[가-힣]/g) || []).length;
  return replacementCharacters * 20 + mojibakeMarkers * 3 - koreanCharacters;
}

function findLikelyColumn(headers, needles, optional = false) {
  const normalized = headers.map((header) => ({
    original: header,
    value: String(header).toLowerCase().replace(/\s/g, ""),
  }));
  return (
    normalized.find((header) => needles.some((needle) => header.value.includes(needle.toLowerCase())))?.original ||
    (optional ? "" : headers[0])
  );
}

function fillSelect(select, headers, selected, optional) {
  select.innerHTML = "";
  if (optional) {
    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    emptyOption.textContent = "없음";
    emptyOption.selected = !selected;
    select.append(emptyOption);
  }
  headers.forEach((header) => {
    const option = document.createElement("option");
    option.value = header;
    option.textContent = header;
    option.selected = header === selected;
    select.append(option);
  });
}

function applyMapping() {
  state.studentIdColumn = els.studentIdColumn.value;
  state.departmentColumn = els.departmentColumn.value;
  state.gradeColumn = els.gradeColumn.value;
  state.semesterColumn = els.semesterColumn.value;
  state.gpaColumn = els.gpaColumn.value;
  state.toeicColumnA = els.toeicColumnA.value;
  state.toeicDateColumnA = els.toeicDateColumnA.value;
  state.toeicColumnB = els.toeicColumnB.value;
  state.toeicDateColumnB = els.toeicDateColumnB.value;
  state.mockToeicColumnA = els.mockToeicColumnA.value;
  state.mockToeicDateColumnA = els.mockToeicDateColumnA.value;
  state.mockToeicColumnB = els.mockToeicColumnB.value;
  state.mockToeicDateColumnB = els.mockToeicDateColumnB.value;
  state.analysisCacheValid = false;
  state.analysisCacheKey = "";
  state.analysisRowsCache = [];
  state.parsedRows = state.rawRows
    .map((row, index) => ({
      index: index + 1,
      raw: row,
      studentId: getStringValue(row, state.studentIdColumn) || `row-${index + 1}`,
      department: getStringValue(row, state.departmentColumn) || "미분류",
      gradeLevel: normalizeGrade(getStringValue(row, state.gradeColumn)) || "미분류",
      semester: normalizeSemester(getStringValue(row, state.semesterColumn)) || "미분류",
      gpa: toNumber(row[state.gpaColumn]),
      toeicA: state.toeicColumnA ? toNumber(row[state.toeicColumnA]) : Number.NaN,
      toeicDateA: parseDateValue(row[state.toeicDateColumnA]),
      toeicB: state.toeicColumnB ? toNumber(row[state.toeicColumnB]) : Number.NaN,
      toeicDateB: parseDateValue(row[state.toeicDateColumnB]),
      mockToeicA: state.mockToeicColumnA ? toNumber(row[state.mockToeicColumnA]) : Number.NaN,
      mockToeicDateA: parseDateValue(row[state.mockToeicDateColumnA]),
      mockToeicB: state.mockToeicColumnB ? toNumber(row[state.mockToeicColumnB]) : Number.NaN,
      mockToeicDateB: parseDateValue(row[state.mockToeicDateColumnB]),
    }))
    .filter((row) => Number.isFinite(row.gpa)); // 토익 없는 학생도 분석 모집단에 포함

  if (!state.parsedRows.length) {
    alert("선택한 열에서 숫자 데이터를 찾지 못했습니다.");
    return;
  }

  const maxGpa = Math.max(...state.parsedRows.map((row) => row.gpa));
  if (maxGpa > 4.5 && maxGpa <= 100) {
    const shouldNormalize = confirm("평균평점이 100점 만점처럼 보입니다. 4.5 만점 기준으로 환산할까요?");
    if (shouldNormalize) {
      state.parsedRows = state.parsedRows.map((row) => ({ ...row, gpa: (row.gpa / 100) * 4.5 }));
    }
  }

  populateFilters();
  state.rows = getAnalysisRows();
  updateThresholdBounds(state.rows, false);
  els.gpaThreshold.value = state.scoreBounds.minGpa.toFixed(2);
  els.gpaThresholdNumber.value = state.scoreBounds.minGpa.toFixed(2);
  els.toeicThreshold.value = state.scoreBounds.minToeic;
  els.toeicThresholdNumber.value = state.scoreBounds.minToeic;
  els.mockToeicThreshold.value = state.scoreBounds.minMockToeic;
  els.mockToeicThresholdNumber.value = state.scoreBounds.minMockToeic;

  els.uploadPanel.classList.add("hidden");
  els.mappingPanel.classList.add("hidden");
  els.dashboard.classList.remove("hidden");
  document.getElementById("targetRateSection").classList.toggle("hidden", !isAutoTargetMode());
  if (isAutoTargetMode()) applyTargetRateThresholds(true);
  else render();
}

function getStringValue(row, column) {
  if (!column) return "";
  const value = row[column];
  return value === undefined || value === null ? "" : String(value).trim();
}

function normalizeGrade(value) {
  if (!value) return "";
  const number = String(value).match(/\d+/)?.[0];
  return number ? `${number}학년` : value;
}

function normalizeSemester(value) {
  if (!value) return "";
  const number = String(value).match(/\d+/)?.[0];
  if (!number) return value;
  const semester = clamp(Number(number), 1, 8);
  return `${semester}학기`;
}

function parseDateValue(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && value > 20000 && value < 80000) {
    return new Date(Math.round((value - 25569) * 86400 * 1000));
  }
  const text = String(value).trim().replace(/\./g, "-").replace(/\//g, "-");
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? null : date;
}

function populateFilters() {
  fillFilter(els.departmentFilter, ["전체 학과", ...uniqueSorted(state.parsedRows.map((row) => row.department))]);
  fillFilter(els.gradeFilter, ["전체 학년", ...uniqueSorted(state.parsedRows.map((row) => row.gradeLevel), compareGrade)]);
  fillSemesterFilter();

  const hasA = state.parsedRows.some((row) => Number.isFinite(row.toeicA));
  const hasB = state.parsedRows.some((row) => Number.isFinite(row.toeicB));
  const hasMockA = state.parsedRows.some((row) => Number.isFinite(row.mockToeicA));
  const hasMockB = state.parsedRows.some((row) => Number.isFinite(row.mockToeicB));

  document.querySelectorAll('input[name="toeicPeriod"]').forEach((input) => {
    if (input.value === "recent") {
      input.disabled = !(hasA || hasMockA);
      if (!(hasA || hasMockA) && input.checked) {
        document.querySelector('input[name="toeicPeriod"][value="all"]').checked = true;
      }
    }
    if (input.value === "all") {
      input.disabled = !(hasB || hasMockB);
    }
  });

  els.periodHelp.textContent =
    (hasA || hasMockA) && (hasB || hasMockB) ? "A열: 최근 2년 최고점 · B열: 전체 기간 최고점"
    : !(hasA || hasMockA)                    ? "최근 2년 성적(A열) 없음 — 전체 기간만 분석"
    :                                          "전체 기간 성적(B열) 없음 — 최근 2년만 분석";
}

function fillSemesterFilter() {
  els.semesterFilter.innerHTML = "";
  const semesters = uniqueSorted(state.parsedRows.map((row) => row.semester), compareSemester);
  const defaultSemesters = semesters.length ? semesters : Array.from({ length: 8 }, (_, index) => `${index + 1}학기`);
  defaultSemesters.forEach((semester) => {
    const label = document.createElement("label");
    label.className = "semester-option";
    label.innerHTML = `
      <input type="checkbox" value="${semester}" checked />
      <span>${semester}</span>
    `;
    els.semesterFilter.append(label);
  });
}

function fillFilter(select, options) {
  select.innerHTML = "";
  options.forEach((value, index) => {
    const option = document.createElement("option");
    option.value = index === 0 ? "all" : value;
    option.textContent = value;
    select.append(option);
  });
}

function uniqueSorted(values, sorter = undefined) {
  return [...new Set(values.filter(Boolean))].sort(sorter);
}

function compareGrade(a, b) {
  return (Number(a.match(/\d+/)?.[0]) || 99) - (Number(b.match(/\d+/)?.[0]) || 99) || a.localeCompare(b, "ko");
}

function compareSemester(a, b) {
  return (Number(a.match(/\d+/)?.[0]) || 99) - (Number(b.match(/\d+/)?.[0]) || 99) || a.localeCompare(b, "ko");
}

function updateThresholdBounds(rows, clampCurrent = true) {
  if (!rows.length) return;

  const gpaValues = rows.map((row) => row.gpa);
  const toeicValues = rows.map((row) => row.toeic).filter((v) => Number.isFinite(v));
  const mockToeicValues = rows.map((row) => row.mockToeic).filter((v) => Number.isFinite(v));
  const minGpa = Number(Math.min(...gpaValues).toFixed(2));
  const maxGpa = Number(Math.max(...gpaValues).toFixed(2));
  const minToeic = toeicValues.length ? Math.min(...toeicValues) : 0;
  const maxToeic = toeicValues.length ? Math.max(...toeicValues) : 990;
  const minMockToeic = mockToeicValues.length ? Math.min(...mockToeicValues) : 0;
  const maxMockToeic = mockToeicValues.length ? Math.max(...mockToeicValues) : 990;

  state.scoreBounds = {
    minGpa,
    maxGpa,
    minToeic,
    maxToeic,
    minMockToeic,
    maxMockToeic,
  };

  els.gpaThreshold.min = minGpa;
  els.gpaThreshold.max = maxGpa;
  els.gpaThresholdNumber.min = minGpa;
  els.gpaThresholdNumber.max = maxGpa;
  els.toeicThreshold.min = minToeic;
  els.toeicThreshold.max = maxToeic;
  els.toeicThreshold.step = 1;
  els.toeicThresholdNumber.min = minToeic;
  els.toeicThresholdNumber.max = maxToeic;
  els.toeicThresholdNumber.step = 1;
  els.mockToeicThreshold.min = minMockToeic;
  els.mockToeicThreshold.max = maxMockToeic;
  els.mockToeicThreshold.step = 1;
  els.mockToeicThresholdNumber.min = minMockToeic;
  els.mockToeicThresholdNumber.max = maxMockToeic;
  els.mockToeicThresholdNumber.step = 1;
  els.mockToeicThreshold.disabled = !mockToeicValues.length;
  els.mockToeicThresholdNumber.disabled = !mockToeicValues.length;
  els.mockToeicHelp.textContent = mockToeicValues.length
    ? "토익 또는 모의토익 중 하나만 충족해도 토익성적 충족으로 계산"
    : "모의토익 자료가 없으면 기존 토익 기준만 적용";

  if (!clampCurrent) return;

  const nextGpa = clamp(Number(els.gpaThreshold.value), minGpa, maxGpa);
  const nextToeic = clamp(Number(els.toeicThreshold.value), minToeic, maxToeic);
  const nextMockToeic = clamp(Number(els.mockToeicThreshold.value), minMockToeic, maxMockToeic);
  els.gpaThreshold.value = nextGpa.toFixed(2);
  els.gpaThresholdNumber.value = nextGpa.toFixed(2);
  els.toeicThreshold.value = nextToeic;
  els.toeicThresholdNumber.value = nextToeic;
  els.mockToeicThreshold.value = nextMockToeic;
  els.mockToeicThresholdNumber.value = nextMockToeic;
}

function toNumber(value) {
  if (typeof value === "number") return value;
  const cleaned = String(value).replace(/,/g, "").replace(/[^\d.-]/g, "");
  return cleaned === "" ? Number.NaN : Number(cleaned);
}

function render() {
  if (!state.parsedRows.length) return;
  state.rows = getAnalysisRows();
  updateThresholdBounds(state.rows);

  const gpaThreshold = Number(els.gpaThreshold.value);
  const toeicThreshold = Number(els.toeicThreshold.value);
  const mockToeicThreshold = Number(els.mockToeicThreshold.value);
  const mockAvailable = hasMockToeicData();
  const total = state.rows.length;
  if (!total) {
    renderEmptyState();
    return;
  }
  let gpaPassCount = 0;
  let toeicPassCount = 0;
  let bothPassCount = 0;
  const gpaValues = [];
  const toeicScores = [];

  state.rows.forEach((row) => {
    const gpaPass = row.gpa >= gpaThreshold;
    const toeicPass = passesLanguageCriterion(row, toeicThreshold, mockToeicThreshold, mockAvailable);
    if (gpaPass) gpaPassCount += 1;
    if (toeicPass) toeicPassCount += 1;
    if (gpaPass && toeicPass) bothPassCount += 1;
    gpaValues.push(row.gpa);
    if (Number.isFinite(row.toeic)) toeicScores.push(row.toeic);
  });

  renderSummaryChips(total);
  els.totalCount.textContent = "100%";
  els.totalCountDetail.textContent = `${numberFormat.format(total)}명`;
  els.gpaPass.textContent = rateText(gpaPassCount, total);
  els.toeicPass.textContent = rateText(toeicPassCount, total);
  els.bothPass.textContent = rateText(bothPassCount, total);
  els.gpaPassRate.textContent = `${numberFormat.format(gpaPassCount)}명`;
  els.toeicPassRate.textContent = `${numberFormat.format(toeicPassCount)}명`;
  els.bothPassRate.textContent = `${numberFormat.format(bothPassCount)}명`;

  renderHeatmap(gpaThreshold, toeicThreshold);
  renderScatter(gpaThreshold, toeicThreshold);
  const noScoreCount = state.rows.length - toeicScores.length;
  renderHistogram(els.gpaHistogram, gpaValues, gpaBins, gpaThreshold, 0);
  renderHistogram(els.toeicHistogram, toeicScores, toeicBins, toeicThreshold, noScoreCount);
  renderSensitivity(gpaThreshold, toeicThreshold);
  els.gpaStats.textContent = statText(gpaValues, 2);
  els.toeicStats.textContent = toeicScores.length
    ? statText(toeicScores, 0) + ` · 미응시 ${numberFormat.format(noScoreCount)}명`
    : `미응시 ${numberFormat.format(noScoreCount)}명`;
}

function getAnalysisRows() {
  const department = els.departmentFilter.value;
  const grade = els.gradeFilter.value;
  const period = document.querySelector('input[name="toeicPeriod"]:checked')?.value || "all";
  const selectedSemesters = getSelectedSemesters();
  const semesterKey = [...selectedSemesters].sort(compareSemester).join("|");
  const cacheKey = `${department}__${grade}__${period}__${semesterKey}`;

  if (state.analysisCacheValid && state.analysisCacheKey === cacheKey) {
    return state.analysisRowsCache;
  }

  const rows = state.parsedRows
    .filter((row) => {
      if (department !== "all" && row.department !== department) return false;
      if (grade !== "all" && row.gradeLevel !== grade) return false;
      if (selectedSemesters.size && !selectedSemesters.has(row.semester)) return false;
      return true;
    })
    .map((row) => ({
      ...row,
      toeic: period === "recent" ? row.toeicA : row.toeicB,
      mockToeic: period === "recent" ? row.mockToeicA : row.mockToeicB,
      // toeic이 NaN인 학생 = 미응시, 모집단에 포함하되 충족 기준 미달 처리됨
    }));

  const grouped = new Map();
  rows.forEach((row) => {
    const previous = grouped.get(row.studentId);
    const rowHasLanguageScore = Number.isFinite(row.toeic) || Number.isFinite(row.mockToeic);
    const previousHasLanguageScore = previous && (Number.isFinite(previous.toeic) || Number.isFinite(previous.mockToeic));
    const rowBestLanguageScore = bestLanguageScore(row);
    const previousBestLanguageScore = previous ? bestLanguageScore(previous) : -Infinity;
    if (
      !previous ||
      (rowHasLanguageScore && !previousHasLanguageScore) ||
      (rowHasLanguageScore === previousHasLanguageScore && rowBestLanguageScore > previousBestLanguageScore) ||
      (rowHasLanguageScore === previousHasLanguageScore && rowBestLanguageScore === previousBestLanguageScore && row.gpa > previous.gpa)
    ) {
      grouped.set(row.studentId, row);
    }
  });
  const analysisRows = [...grouped.values()];
  state.analysisCacheValid = true;
  state.analysisCacheKey = cacheKey;
  state.analysisRowsCache = analysisRows;
  return analysisRows;
}

function hasMockToeicData() {
  return state.rows.some((row) => Number.isFinite(row.mockToeic));
}

function passesLanguageCriterion(row, toeicThreshold, mockToeicThreshold, mockAvailable = hasMockToeicData()) {
  return row.toeic >= toeicThreshold || (mockAvailable && row.mockToeic >= mockToeicThreshold);
}

function bestLanguageScore(row) {
  return Math.max(
    Number.isFinite(row.toeic) ? row.toeic : -Infinity,
    Number.isFinite(row.mockToeic) ? row.mockToeic : -Infinity,
  );
}

function getSelectedSemesters() {
  return new Set(
    [...els.semesterFilter.querySelectorAll('input[type="checkbox"]:checked')]
      .map((input) => input.value)
      .filter(Boolean),
  );
}

function createLanguageSweep(rows, mockToeicThreshold) {
  const sortedRows = [...rows].sort((a, b) => b.gpa - a.gpa);
  const maxToeicScore = Math.max(
    990,
    ...rows.map((row) => (
      Number.isFinite(row.toeic)
        ? Math.ceil(row.toeic)
        : 0
    )),
  );
  const toeicFreq = new Uint32Array(maxToeicScore + 2);
  const suffixCounts = new Uint32Array(maxToeicScore + 2);
  let cursor = 0;
  let autoPassCount = 0;
  let suffixDirty = true;

  function markRow(row) {
    if (Number.isFinite(row.mockToeic) && row.mockToeic >= mockToeicThreshold) {
      autoPassCount += 1;
      return;
    }
    if (!Number.isFinite(row.toeic)) return;
    const score = clamp(Math.ceil(row.toeic), 0, maxToeicScore);
    toeicFreq[score] += 1;
    suffixDirty = true;
  }

  function ensureSuffixCounts() {
    if (!suffixDirty) return;
    let running = 0;
    for (let score = maxToeicScore; score >= 0; score -= 1) {
      running += toeicFreq[score];
      suffixCounts[score] = running;
    }
    suffixDirty = false;
  }

  return {
    advanceToGpa(threshold) {
      while (cursor < sortedRows.length && sortedRows[cursor].gpa >= threshold) {
        markRow(sortedRows[cursor]);
        cursor += 1;
      }
    },
    countAtToeic(threshold) {
      ensureSuffixCounts();
      const score = clamp(Math.ceil(threshold), 0, maxToeicScore + 1);
      return autoPassCount + (score > maxToeicScore ? 0 : suffixCounts[score]);
    },
  };
}

function lowerBound(sortedValues, target) {
  let left = 0;
  let right = sortedValues.length;
  while (left < right) {
    const mid = Math.floor((left + right) / 2);
    if (sortedValues[mid] < target) left = mid + 1;
    else right = mid;
  }
  return left;
}

function percentileRankFromSorted(sortedValues, value) {
  if (!sortedValues.length) return 0;
  return lowerBound(sortedValues, value) / sortedValues.length;
}

function renderSummaryChips(total) {
  const department = els.departmentFilter.value === "all" ? "전체 학과" : els.departmentFilter.value;
  const grade = els.gradeFilter.value === "all" ? "전체 학년" : els.gradeFilter.value;
  const semesters = [...getSelectedSemesters()].sort(compareSemester);
  const semesterText = semesters.length === els.semesterFilter.querySelectorAll('input[type="checkbox"]').length
    ? "전체 인정학기"
    : semesters.join(", ");
  const period = document.querySelector('input[name="toeicPeriod"]:checked')?.value === "recent" ? "최근 2년" : "전체 기간";

  const ic = (path) =>
    `<svg xmlns="http://www.w3.org/2000/svg" width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${path}</svg>`;

  const svgDept   = ic(`<rect x="3" y="9" width="18" height="13" rx="1"/><path d="M8 9V5a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v4"/>`);
  const svgGrade  = ic(`<path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/>`);
  const svgSemester = ic(`<path d="M4 19.5V6a2 2 0 0 1 2-2h12v16H6a2 2 0 0 1-2-.5z"/><path d="M8 8h6M8 12h4"/>`);
  const svgPeriod = ic(`<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>`);
  const svgCount  = ic(`<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/>`);
  const svgFile = ic(`<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M8 13h8M8 17h5"/>`);

  els.dataSummary.innerHTML = [
    `<span class="summary-chip chip-file">${svgFile}${state.fileName || "분석 파일"}</span>`,
    `<span class="summary-chip">${svgDept}${department}</span>`,
    `<span class="summary-chip">${svgGrade}${grade}</span>`,
    `<span class="summary-chip">${svgSemester}${semesterText}</span>`,
    `<span class="summary-chip">${svgPeriod}${period} 토익</span>`,
    `<span class="summary-chip chip-count">${svgCount}전체 ${numberFormat.format(total)}명</span>`,
  ].join("");
}

function renderEmptyState() {
  renderSummaryChips(0);
  els.totalCount.textContent = "0%";
  els.totalCountDetail.textContent = "0명";
  [els.gpaPass, els.toeicPass, els.bothPass].forEach((element) => {
    element.textContent = "0%";
  });
  [els.gpaPassRate, els.toeicPassRate, els.bothPassRate].forEach((element) => {
    element.textContent = "0명";
  });
  els.heatmap.innerHTML = "";
  els.gpaHistogram.innerHTML = "";
  els.toeicHistogram.innerHTML = "";
  els.gpaStats.textContent = "-";
  els.toeicStats.textContent = "-";
  renderBlankCanvas(els.scatterCanvas, "선택한 조건에 해당하는 학생이 없습니다.");
  renderBlankCanvas(els.sensitivityCanvas, "선택한 조건에 해당하는 학생이 없습니다.");
}

function setMetric(element, value) {
  element.textContent = numberFormat.format(value);
}

function rateText(count, total) {
  return `${percentFormat.format((count / total) * 100)}%`;
}

function renderHeatmap(gpaThreshold, toeicThreshold) {
  const counts = toeicBins.map(() => gpaBins.map(() => 0));

  state.rows.forEach((row) => {
    const x = findBin(row.gpa, gpaBins);
    const y = findBin(row.toeic, toeicBins);
    if (x >= 0 && y >= 0) counts[y][x] += 1;
  });

  const max = Math.max(1, ...counts.flat());
  els.heatmap.innerHTML = "";
  els.heatmap.append(makeDiv("heatmap-corner", "토익\\평점"));
  gpaBins.forEach((bin) => els.heatmap.append(makeDiv("heatmap-label", labelBin(bin, 1))));

  counts
    .map((row, index) => ({ row, bin: toeicBins[index] }))
    .reverse()
    .forEach(({ row, bin }) => {
      els.heatmap.append(makeDiv("heatmap-label", labelBin(bin, 0)));
      row.forEach((count, gpaIndex) => {
        const percent = (count / state.rows.length) * 100;
        const intensity = count / max;
        const cell = makeDiv("heatmap-cell", `${numberFormat.format(count)}명\n${percentFormat.format(percent)}%`);
        const eligible = gpaBins[gpaIndex][1] > gpaThreshold && bin[1] > toeicThreshold;
        cell.classList.toggle("eligible", eligible);
        cell.style.background = eligible
          ? `rgba(16, 185, 129, ${0.10 + intensity * 0.65})`
          : `rgba(244, 63, 94, ${0.05 + intensity * 0.38})`;
        cell.title = `분포지수: ${percentFormat.format(percent)}%`;
        els.heatmap.append(cell);
      });
    });
}

function makeDiv(className, text) {
  const div = document.createElement("div");
  div.className = className;
  div.textContent = text;
  return div;
}

function labelBin([start, end], digits) {
  return `${start.toFixed(digits)}-${end.toFixed(digits)}`;
}

function findBin(value, bins) {
  return bins.findIndex(([start, end], index) => value >= start && (value < end || index === bins.length - 1));
}

function renderScatter(gpaThreshold, toeicThreshold) {
  els.scatterGpaValue.textContent = gpaThreshold.toFixed(2);
  els.scatterToeicValue.textContent = numberFormat.format(toeicThreshold);

  const canvas = els.scatterCanvas;
  const ctx = canvas.getContext("2d");
  const width = canvas.width;
  const height = canvas.height;
  const pad = 46;
  const plotW = width - pad * 2;
  const plotH = height - pad * 2;

  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#08111f";
  ctx.fillRect(0, 0, width, height);
  ctx.strokeStyle = "rgba(148, 163, 184, 0.10)";
  ctx.lineWidth = 1;

  for (let i = 0; i <= 5; i += 1) {
    const x = pad + (plotW / 5) * i;
    const y = pad + (plotH / 5) * i;
    line(ctx, x, pad, x, height - pad);
    line(ctx, pad, y, width - pad, y);
  }

  const toX = (gpa) => pad + (clamp(gpa, 0, 4.5) / 4.5) * plotW;
  const toY = (toeic) => height - pad - (clamp(toeic, 0, 990) / 990) * plotH;
  const sampleStep = Math.max(1, Math.floor(state.rows.length / 7000));

  state.rows.forEach((row, index) => {
    if (index % sampleStep !== 0) return;
    if (!Number.isFinite(row.toeic)) return; // 미응시 제외
    const eligible = row.gpa >= gpaThreshold && row.toeic >= toeicThreshold;
    ctx.fillStyle = eligible ? "rgba(16, 185, 129, 0.72)" : "rgba(244, 63, 94, 0.22)";
    ctx.beginPath();
    ctx.arc(toX(row.gpa), toY(row.toeic), eligible ? 2.4 : 1.8, 0, Math.PI * 2);
    ctx.fill();
  });

  ctx.strokeStyle = "rgba(244, 63, 94, 0.85)";
  ctx.lineWidth = 2;
  const thresholdX = toX(gpaThreshold);
  const thresholdY = toY(toeicThreshold);
  line(ctx, thresholdX, pad, thresholdX, height - pad);
  line(ctx, pad, thresholdY, width - pad, thresholdY);

  drawThresholdBadge(ctx, `평점 ${gpaThreshold.toFixed(2)}`, clamp(thresholdX + 8, pad + 4, width - pad - 92), pad + 10, "#10b981");
  drawThresholdBadge(ctx, `토익 ${numberFormat.format(toeicThreshold)}`, pad + 10, clamp(thresholdY - 30, pad + 4, height - pad - 30), "#f43f5e");

  ctx.fillStyle = "#94a3b8";
  ctx.font = "700 14px sans-serif";
  ctx.fillText("평균평점", width - 104, height - 14);
  ctx.save();
  ctx.translate(16, 120);
  ctx.rotate(-Math.PI / 2);
  ctx.fillText("토익성적", 0, 0);
  ctx.restore();
}

function drawThresholdBadge(ctx, text, x, y, color) {
  ctx.save();
  ctx.font = "800 13px sans-serif";
  const width = Math.ceil(ctx.measureText(text).width) + 18;
  const height = 24;
  ctx.fillStyle = "rgba(8, 17, 31, 0.92)";
  ctx.strokeStyle = color;
  ctx.lineWidth = 1.5;
  roundRect(ctx, x, y, width, height, 5);
  ctx.fill();
  ctx.stroke();
  ctx.fillStyle = "#f0f4ff";
  ctx.fillText(text, x + 9, y + 16);
  ctx.restore();
}

function roundRect(ctx, x, y, width, height, radius) {
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.arcTo(x + width, y, x + width, y + height, radius);
  ctx.arcTo(x + width, y + height, x, y + height, radius);
  ctx.arcTo(x, y + height, x, y, radius);
  ctx.arcTo(x, y, x + width, y, radius);
  ctx.closePath();
}

function renderBlankCanvas(canvas, message) {
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#08111f";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#475569";
  ctx.font = "700 18px sans-serif";
  ctx.textAlign = "center";
  ctx.fillText(message, canvas.width / 2, canvas.height / 2);
  ctx.textAlign = "left";
}

function renderSensitivity(gpaThreshold, toeicThreshold) {
  const canvas = els.sensitivityCanvas;
  const ctx = canvas.getContext("2d");
  const W = canvas.width;
  const H = canvas.height;
  const total = state.rows.length || 1;

  ctx.clearRect(0, 0, W, H);
  ctx.fillStyle = "#08111f";
  ctx.fillRect(0, 0, W, H);

  if (!total) { renderBlankCanvas(canvas, "데이터가 없습니다."); return; }

  const padTop = 26;
  const padBottom = 64;
  const padLeft = 72;
  const padRight = 132;
  const plotW = W - padLeft - padRight;
  const plotH = H - padTop - padBottom;

  const gMin = state.scoreBounds.minGpa;
  const gMax = state.scoreBounds.maxGpa;
  const tMin = state.scoreBounds.minToeic;
  const tMax = state.scoreBounds.maxToeic;
  const mockToeicThreshold = Number(els.mockToeicThreshold.value);
  const mockAvailable = hasMockToeicData();
  const finiteToeicRows = state.rows.filter((row) => Number.isFinite(row.toeic));

  if (!finiteToeicRows.length) {
    renderBlankCanvas(canvas, "선택한 토익 기간의 성적이 없습니다.");
    return;
  }

  if (gMin === gMax || tMin === tMax) {
    renderBlankCanvas(canvas, "민감도 분석에는 서로 다른 평점과 토익 점수가 필요합니다.");
    return;
  }

  // Grid resolution
  const nG = Math.min(48, Math.max(8, Math.ceil((gMax - gMin) / 0.1) + 1));
  const nT = Math.min(40, Math.max(6, Math.ceil((tMax - tMin) / 25) + 1));
  const gStep = (gMax - gMin) / (nG - 1);
  const tStep = (tMax - tMin) / (nT - 1);
  const gThresholds = Array.from({ length: nG }, (_, index) => gMin + index * gStep);
  const tThresholds = Array.from({ length: nT }, (_, index) => tMin + index * tStep);

  // Precompute rate grid with a cumulative GPA sweep so large datasets remain interactive.
  const grid = [];
  let gridMax = 0;
  const sweep = createLanguageSweep(state.rows, mockToeicThreshold);
  for (let gi = nG - 1; gi >= 0; gi -= 1) {
    const g = gThresholds[gi];
    sweep.advanceToGpa(g);
    grid[gi] = [];
    for (let ti = 0; ti < nT; ti += 1) {
      const count = sweep.countAtToeic(tThresholds[ti]);
      const rate = count / total;
      grid[gi][ti] = rate;
      if (rate > gridMax) gridMax = rate;
    }
  }
  if (gridMax === 0) { renderBlankCanvas(canvas, "조건을 충족하는 학생이 없습니다."); return; }

  const toX = (gi) => padLeft + (gi / (nG - 1)) * plotW;
  const toY = (ti) => padTop + plotH - (ti / (nT - 1)) * plotH;

  // Draw heatmap cells
  const cW = plotW / (nG - 1);
  const cH = plotH / (nT - 1);
  for (let gi = 0; gi < nG - 1; gi++) {
    for (let ti = 0; ti < nT - 1; ti++) {
      const avg = (grid[gi][ti] + grid[gi + 1][ti] + grid[gi][ti + 1] + grid[gi + 1][ti + 1]) / 4;
      ctx.fillStyle = sensitivityColor(avg / gridMax);
      ctx.fillRect(toX(gi), toY(ti + 1), cW + 0.6, cH + 0.6);
    }
  }

  // Iso-rate contour labels
  const isoTargets = [0.05, 0.10, 0.15, 0.20, 0.25, 0.30, 0.40, 0.50];
  ctx.font = "800 11px sans-serif";
  ctx.textAlign = "center";
  const labeled = new Set();
  isoTargets.forEach((iso) => {
    if (iso >= gridMax) return;
    for (let gi = 1; gi < nG - 1; gi++) {
      for (let ti = 1; ti < nT - 1; ti++) {
        const r = grid[gi][ti];
        const rr = grid[gi + 1]?.[ti] ?? r;
        const ru = grid[gi]?.[ti + 1] ?? r;
        if ((r - iso) * (rr - iso) < 0 || (r - iso) * (ru - iso) < 0) {
          const key = `${Math.round(gi / 3)},${Math.round(ti / 3)}`;
          if (!labeled.has(key)) {
            labeled.add(key);
            const cx2 = toX(gi);
            const cy2 = toY(ti);
            ctx.fillStyle = "rgba(8, 17, 31, 0.88)";
            ctx.fillRect(cx2 - 17, cy2 - 11, 34, 16);
            ctx.fillStyle = "rgba(240, 244, 255, 0.9)";
            ctx.fillText(`${Math.round(iso * 100)}%`, cx2, cy2);
          }
        }
      }
    }
  });
  ctx.textAlign = "left";

  // Current crosshair
  const cgFrac = (gpaThreshold - gMin) / (gMax - gMin);
  const ctFrac = (toeicThreshold - tMin) / (tMax - tMin);
  const cx = padLeft + clamp(cgFrac, 0, 1) * plotW;
  const cy = padTop + (1 - clamp(ctFrac, 0, 1)) * plotH;

  ctx.setLineDash([6, 4]);
  ctx.strokeStyle = "rgba(255,255,255,0.85)";
  ctx.lineWidth = 1.5;
  ctx.beginPath(); ctx.moveTo(cx, padTop); ctx.lineTo(cx, padTop + plotH); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(padLeft, cy); ctx.lineTo(padLeft + plotW, cy); ctx.stroke();
  ctx.setLineDash([]);

  // Crosshair dot
  ctx.beginPath(); ctx.arc(cx, cy, 8, 0, Math.PI * 2);
  ctx.fillStyle = "rgba(255,255,255,0.9)"; ctx.fill();
  ctx.beginPath(); ctx.arc(cx, cy, 5, 0, Math.PI * 2);
  ctx.fillStyle = "#10b981"; ctx.fill();

  // Current rate bubble
  const currentCount = state.rows.filter((row) => row.gpa >= gpaThreshold && passesLanguageCriterion(row, toeicThreshold, mockToeicThreshold, mockAvailable)).length;
  const curRate = currentCount / total;
  const bubbleText = `${percentFormat.format(curRate * 100)}%`;
  ctx.font = "800 16px sans-serif";
  const bw = ctx.measureText(bubbleText).width + 16;
  const bx = cx + 12 + bw < padLeft + plotW ? cx + 12 : cx - bw - 12;
  const by = cy - 26 > padTop + 6 ? cy - 26 : cy + 14;
  ctx.fillStyle = "rgba(16, 185, 129, 0.92)";
  ctx.beginPath();
  ctx.roundRect(bx - 7, by - 16, bw + 4, 26, 6);
  ctx.fill();
  ctx.fillStyle = "#012118";
  ctx.fillText(bubbleText, bx, by);

  // X axis — GPA
  ctx.fillStyle = "#94a3b8";
  ctx.font = "700 13px sans-serif";
  ctx.textAlign = "center";
  const gTickN = Math.min(8, nG - 1);
  for (let i = 0; i <= gTickN; i++) {
    const gi = Math.round((i / gTickN) * (nG - 1));
    const gpa = gMin + gi * gStep;
    ctx.fillText(gpa.toFixed(1), toX(gi), padTop + plotH + 20);
  }
  ctx.fillStyle = "#94a3b8";
  ctx.font = "800 14px sans-serif";
  ctx.fillText("평균평점 기준 →", padLeft + plotW / 2, padTop + plotH + 46);

  // Y axis — TOEIC
  ctx.textAlign = "right";
  ctx.fillStyle = "#94a3b8";
  ctx.font = "700 13px sans-serif";
  const tTickN = Math.min(7, nT - 1);
  for (let i = 0; i <= tTickN; i++) {
    const ti = Math.round((i / tTickN) * (nT - 1));
    const toeic = tMin + ti * tStep;
    ctx.fillText(Math.round(toeic), padLeft - 10, toY(ti) + 5);
  }
  ctx.save();
  ctx.translate(18, padTop + plotH / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.fillStyle = "#94a3b8";
  ctx.font = "800 14px sans-serif";
  ctx.textAlign = "center";
  ctx.fillText("토익 기준 →", 0, 0);
  ctx.restore();
  ctx.textAlign = "left";

  // Plot border
  ctx.strokeStyle = "rgba(148, 163, 184, 0.15)";
  ctx.lineWidth = 1;
  ctx.strokeRect(padLeft, padTop, plotW, plotH);

  // Legend gradient bar
  const lgX = W - padRight + 20;
  const lgY = padTop;
  const lgH = plotH;
  const lgW = 16;
  for (let i = 0; i < lgH; i++) {
    ctx.fillStyle = sensitivityColor(1 - i / lgH);
    ctx.fillRect(lgX, lgY + i, lgW, 1.5);
  }
  ctx.strokeStyle = "rgba(148, 163, 184, 0.15)";
  ctx.lineWidth = 1;
  ctx.strokeRect(lgX, lgY, lgW, lgH);

  ctx.fillStyle = "#94a3b8";
  ctx.font = "700 12px sans-serif";
  ctx.textAlign = "left";
  for (let i = 0; i <= 5; i++) {
    const rate = (1 - i / 5) * gridMax;
    const ly = lgY + (i / 5) * lgH;
    const label = i === 0 ? `최대 ${Math.round(rate * 100)}%` : `${Math.round(rate * 100)}%`;
    ctx.fillText(label, lgX + lgW + 5, ly + 4);
  }
  ctx.fillStyle = "#94a3b8";
  ctx.font = "800 12px sans-serif";
  ctx.fillText("동시충족", lgX - 2, lgY - 7);
}

function sensitivityColor(t) {
  // t=0: near-black, t=1: bright emerald (dark mode)
  const stops = [
    [0,    [8,   17,  31]],
    [0.2,  [4,   55,  43]],
    [0.42, [6,   95,  70]],
    [0.68, [16,  185, 129]],
    [1.0,  [52,  211, 153]],
  ];
  for (let i = 1; i < stops.length; i++) {
    const [t0, c0] = stops[i - 1];
    const [t1, c1] = stops[i];
    if (t <= t1) {
      const s = (t - t0) / (t1 - t0);
      return `rgb(${Math.round(c0[0] + s * (c1[0] - c0[0]))},${Math.round(c0[1] + s * (c1[1] - c0[1]))},${Math.round(c0[2] + s * (c1[2] - c0[2]))})`;
    }
  }
  return "rgb(99,72,210)";
}

function line(ctx, x1, y1, x2, y2) {
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
}

function renderHistogram(container, values, bins, threshold, noScoreCount = 0) {
  const counts = bins.map(([start, end], index) =>
    values.filter((value) => value >= start && (value < end || index === bins.length - 1)).length,
  );
  const max = Math.max(1, ...counts, noScoreCount);
  container.innerHTML = "";

  // 미응시 바 (상단에 별도 표시)
  if (noScoreCount > 0) {
    const row = document.createElement("div");
    row.className = "bar-row";

    const label = document.createElement("span");
    label.textContent = "미응시";
    label.style.color = "#9aa8ba";

    const track = document.createElement("div");
    track.className = "bar-track";

    const fill = document.createElement("div");
    fill.className = "bar-fill";
    fill.style.width = `${(noScoreCount / max) * 100}%`;
    fill.style.background = "repeating-linear-gradient(135deg, #1e3a2e 0px, #1e3a2e 3px, #0f2419 3px, #0f2419 7px)";
    track.append(fill);

    const count = document.createElement("span");
    count.textContent = numberFormat.format(noScoreCount);
    count.style.color = "#9aa8ba";

    row.append(label, track, count);
    container.append(row);
  }

  bins.forEach((bin, index) => {
    const row = document.createElement("div");
    row.className = "bar-row";

    const label = document.createElement("span");
    label.textContent = labelBin(bin, bin[1] <= 5 ? 1 : 0);

    const track = document.createElement("div");
    track.className = "bar-track";

    const fill = document.createElement("div");
    fill.className = "bar-fill";
    fill.style.width = `${(counts[index] / max) * 100}%`;
    fill.style.background = bin[1] > threshold
      ? "linear-gradient(90deg, #059669, #10b981, #34d399)"
      : "rgba(16, 185, 129, 0.18)";
    track.append(fill);

    const count = document.createElement("span");
    count.textContent = numberFormat.format(counts[index]);

    row.append(label, track, count);
    container.append(row);
  });
}

function statText(values, digits) {
  return `중앙값 ${quantile(values, 0.5).toFixed(digits)} · 상위 25% ${quantile(values, 0.75).toFixed(digits)} · 상위 10% ${quantile(values, 0.9).toFixed(digits)}`;
}

function applyCurrentTarget() {
  applyTargetRateThresholds(true);
}

function adjustComplementaryThreshold(changedCriterion) {
  state.rows = getAnalysisRows();
  if (!state.rows.length) {
    render();
    return;
  }
  updateThresholdBounds(state.rows);

  const target = Number(els.targetRate.value) / 100;
  if (changedCriterion === "mockToeic") {
    applyTargetRateThresholds(true);
    return;
  }
  const candidate =
    changedCriterion === "toeic"
      ? findBestGpaForFixedToeic(Number(els.toeicThreshold.value), target)
      : findBestToeicForFixedGpa(Number(els.gpaThreshold.value), target);

  if (candidate) {
    els.gpaThreshold.value = candidate.gpa.toFixed(2);
    els.gpaThresholdNumber.value = candidate.gpa.toFixed(2);
    els.toeicThreshold.value = candidate.toeic;
    els.toeicThresholdNumber.value = candidate.toeic;
  }

  render();
  updateRecommendationCards(
    candidate ? [candidate] : [],
    candidate
      ? `목표 ${percentFormat.format(target * 100)}% 유지를 위해 ${changedCriterion === "toeic" ? "평균평점" : "토익"} 기준을 자동 보정했습니다.`
      : `목표 ${percentFormat.format(target * 100)}%에 맞는 보정 기준을 찾지 못했습니다.`,
  );
}

function applyTargetRateThresholds(autoApply) {
  state.rows = getAnalysisRows();
  if (!state.rows.length) return;
  updateThresholdBounds(state.rows);

  const target = Number(els.targetRate.value) / 100;
  const best = findTargetCandidates(target).slice(0, 3);
  const top = best[0];

  if (autoApply && top) {
    els.gpaThreshold.value = top.gpa.toFixed(2);
    els.gpaThresholdNumber.value = top.gpa.toFixed(2);
    els.toeicThreshold.value = top.toeic;
    els.toeicThresholdNumber.value = top.toeic;
    render();
  }

  updateRecommendationCards(
    best,
    top && autoApply
      ? `목표 ${percentFormat.format(target * 100)}% 기준을 적용했습니다: 평점 ${top.gpa.toFixed(2)} / 토익 ${top.toeic}점`
      : top
        ? `수동 적용 중입니다. 목표 ${percentFormat.format(target * 100)}%에 가까운 후보: 평점 ${top.gpa.toFixed(2)} / 토익 ${top.toeic}점`
      : `목표 ${percentFormat.format(target * 100)}%에 가까운 기준 후보를 찾지 못했습니다.`,
  );
}

function updateRecommendationCards(candidates, text) {
  if (!els.recommendationText || !els.candidateList) return;
  els.recommendationText.textContent = text;
  els.candidateList.innerHTML = "";
  candidates.forEach((candidate, index) => {
    const card = document.createElement("article");
    card.className = "candidate";
    card.innerHTML = `
      <strong>${index + 1}. 평점 ${candidate.gpa.toFixed(2)} / 토익 ${candidate.toeic}</strong>
      <span>동시 충족 ${numberFormat.format(candidate.count)}명 · ${percentFormat.format(candidate.rate * 100)}%</span>
      <span>목표와 차이 ${percentFormat.format(candidate.distance * 100)}%p</span>
    `;
    card.addEventListener("click", () => {
      els.gpaThreshold.value = candidate.gpa.toFixed(2);
      els.gpaThresholdNumber.value = candidate.gpa.toFixed(2);
      els.toeicThreshold.value = candidate.toeic;
      els.toeicThresholdNumber.value = candidate.toeic;
      render();
    });
    els.candidateList.append(card);
  });
}

function findBestGpaForFixedToeic(toeic, target) {
  const mockToeicThreshold = Number(els.mockToeicThreshold.value);
  const sweep = createLanguageSweep(state.rows, mockToeicThreshold);
  let best = null;

  for (let gpa = state.scoreBounds.maxGpa; gpa >= state.scoreBounds.minGpa - 0.001; gpa -= 0.01) {
    const roundedGpa = Number(gpa.toFixed(2));
    sweep.advanceToGpa(roundedGpa);
    const count = sweep.countAtToeic(toeic);
    const rate = count / state.rows.length;
    const candidate = {
      gpa: roundedGpa,
      toeic,
      count,
      rate,
      distance: Math.abs(rate - target),
      balance: 0,
    };

    if (!best || candidate.distance < best.distance || (candidate.distance === best.distance && candidate.gpa > best.gpa)) {
      best = candidate;
    }
  }

  return best;
}

function findBestToeicForFixedGpa(gpa, target) {
  const mockToeicThreshold = Number(els.mockToeicThreshold.value);
  const sweep = createLanguageSweep(state.rows, mockToeicThreshold);
  sweep.advanceToGpa(gpa);
  let best = null;

  for (let toeic = state.scoreBounds.minToeic; toeic <= state.scoreBounds.maxToeic; toeic += 5) {
    const count = sweep.countAtToeic(toeic);
    const rate = count / state.rows.length;
    const candidate = {
      gpa,
      toeic,
      count,
      rate,
      distance: Math.abs(rate - target),
      balance: 0,
    };

    if (!best || candidate.distance < best.distance || (candidate.distance === best.distance && candidate.toeic > best.toeic)) {
      best = candidate;
    }
  }

  return best;
}

function findTargetCandidates(target) {
  const candidates = [];
  const gpaValuesSorted = [...state.rows.map((row) => row.gpa)].sort((a, b) => a - b);
  const toeicValuesSorted = state.rows
    .map((row) => row.toeic)
    .filter((value) => Number.isFinite(value))
    .sort((a, b) => a - b);
  if (!toeicValuesSorted.length) return candidates;
  const mockToeicThreshold = Number(els.mockToeicThreshold.value);
  const sweep = createLanguageSweep(state.rows, mockToeicThreshold);

  for (let gpa = state.scoreBounds.maxGpa; gpa >= state.scoreBounds.minGpa - 0.001; gpa -= 0.05) {
    const roundedGpa = Number(gpa.toFixed(2));
    sweep.advanceToGpa(roundedGpa);
    for (let toeic = state.scoreBounds.minToeic; toeic <= state.scoreBounds.maxToeic; toeic += 25) {
      const count = sweep.countAtToeic(toeic);
      const rate = count / state.rows.length;
      const distance = Math.abs(rate - target);
      const academicWeight = percentileRankFromSorted(gpaValuesSorted, roundedGpa);
      const languageWeight = percentileRankFromSorted(toeicValuesSorted, toeic);
      candidates.push({
        gpa: roundedGpa,
        toeic,
        count,
        rate,
        distance,
        balance: 1 - Math.abs(academicWeight - languageWeight),
      });
    }
  }

  return candidates.sort(
    (a, b) => a.distance - b.distance || b.balance - a.balance || b.gpa + b.toeic / 1000 - (a.gpa + a.toeic / 1000),
  );
}

function exportEligibleCsv() {
  state.rows = getAnalysisRows();
  const gpaThreshold = Number(els.gpaThreshold.value);
  const toeicThreshold = Number(els.toeicThreshold.value);
  const mockToeicThreshold = Number(els.mockToeicThreshold.value);
  const mockAvailable = hasMockToeicData();
  const eligible = state.rows
    .filter((row) => row.gpa >= gpaThreshold && passesLanguageCriterion(row, toeicThreshold, mockToeicThreshold, mockAvailable))
    .map((row) => ({
      ...row.raw,
      분석_학생ID: row.studentId,
      분석_학과: row.department,
      분석_학년: row.gradeLevel,
      분석_인정학기: row.semester,
      분석_평균평점: row.gpa,
      분석_토익A_최근2년: Number.isFinite(row.toeicA) ? row.toeicA : "",
      분석_토익B_전체기간: Number.isFinite(row.toeicB) ? row.toeicB : "",
      분석_모의토익A_최근2년: Number.isFinite(row.mockToeicA) ? row.mockToeicA : "",
      분석_모의토익B_전체기간: Number.isFinite(row.mockToeicB) ? row.mockToeicB : "",
      분석_적용성적: row.toeic,
      분석_적용모의토익: Number.isFinite(row.mockToeic) ? row.mockToeic : "",
      분석_어학충족구분: row.toeic >= toeicThreshold
        ? "토익"
        : mockAvailable && row.mockToeic >= mockToeicThreshold
          ? "모의토익"
          : "",
      분석_토익기간: document.querySelector('input[name="toeicPeriod"]:checked')?.value === "recent" ? "최근 2년(A)" : "전체 기간(B)",
    }));

  if (!eligible.length) {
    alert("현재 기준을 동시에 충족하는 학생이 없습니다.");
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(eligible);
  const csv = XLSX.utils.sheet_to_csv(worksheet);
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `eligible_gpa_${gpaThreshold}_toeic_${toeicThreshold}_mock_${mockToeicThreshold}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function quantile(values, q) {
  const sorted = [...values].sort((a, b) => a - b);
  const position = (sorted.length - 1) * q;
  const base = Math.floor(position);
  const rest = position - base;
  return sorted[base + 1] === undefined ? sorted[base] : sorted[base] + rest * (sorted[base + 1] - sorted[base]);
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function roundTo(value, step) {
  return Math.round(value / step) * step;
}
