const APP_VERSION = "2.5.0";
const DAY_CUTOFF_SECONDS = 4 * 3600;
const GUIDE_STORAGE_KEY = "saby-guide-collapsed";

const attendanceInput = document.getElementById("attendanceInput");
const staffInput = document.getElementById("staffInput");
const payrollInput = document.getElementById("payrollInput");
const premiumInput = document.getElementById("premiumInput");
const allFilesInput = document.getElementById("allFilesInput");
const statusEl = document.getElementById("status");
const dateSelect = document.getElementById("dateSelect");
const restaurantSelect = document.getElementById("restaurantSelect");
const calcBtn = document.getElementById("calcBtn");
const csvBtn = document.getElementById("csvBtn");
const xlsxBtn = document.getElementById("xlsxBtn");
const summaryEl = document.getElementById("summary");
const tableBody = document.querySelector("#resultTable tbody");
const appVersionEl = document.getElementById("appVersion");
const guidePanelEl = document.getElementById("guidePanel");
const guideToggleBtn = document.getElementById("guideToggleBtn");

let baseRecords = [];
let mappedRecords = [];
let staffRestaurantMap = new Map();
let payrollByPerson = new Map();
let premiumByPerson = new Map();
let premiumLoadInfo = { loaded: false, files: 0, matched: 0 };
let staffConflicts = 0;
let staffConflictKeys = new Set();
let mappingStats = { matched: 0, total: 0 };
let lastResultRows = [];

appVersionEl.textContent = APP_VERSION;

function setGuideCollapsed(collapsed) {
  if (!guidePanelEl || !guideToggleBtn) return;
  guidePanelEl.classList.toggle("collapsed", collapsed);
  guideToggleBtn.textContent = collapsed ? "Показать" : "Скрыть";
  try {
    localStorage.setItem(GUIDE_STORAGE_KEY, collapsed ? "1" : "0");
  } catch (_) {}
}

if (guidePanelEl && guideToggleBtn) {
  let collapsed = false;
  try {
    collapsed = localStorage.getItem(GUIDE_STORAGE_KEY) === "1";
  } catch (_) {}
  setGuideCollapsed(collapsed);
  guideToggleBtn.addEventListener("click", () => {
    const isCollapsed = guidePanelEl.classList.contains("collapsed");
    setGuideCollapsed(!isCollapsed);
  });
}

if (window.pdfjsLib) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = "./vendor/pdf.worker.min.js";
}

function excelDateToSerialDay(value) {
  const days = Number(value);
  if (!Number.isFinite(days)) return NaN;
  return Math.floor(days);
}

function serialDayToISO(serialDay) {
  const utcValue = (serialDay - 25569) * 86400;
  const date = new Date(utcValue * 1000);
  if (Number.isNaN(date.getTime())) return "";
  return date.toISOString().slice(0, 10);
}

function parseExcelTimeToSeconds(value) {
  if (value === null || value === undefined || value === "") return NaN;
  const numeric = Number(value);
  if (Number.isFinite(numeric)) return Math.round(numeric * 24 * 3600);
  const text = String(value).trim();
  const m = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return NaN;
  return Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3] || 0);
}

function normalize(text) {
  return String(text || "").toLowerCase().replace(/ё/g, "е").replace(/\s+/g, " ").trim();
}

function normalizeFio(text) {
  return normalize(text).replace(/[^a-zа-я0-9 ]/gi, "");
}

function classifyRole(roleText) {
  const role = normalize(roleText);
  if (/повар|шеф/.test(role)) return "Кухня";
  if (/официант|менеджер зала|мойщ|мойк/.test(role)) return "Зал";
  if (/логист|курьер|водител/.test(role)) return "Доставка";
  if (/барменедж|барбэк|барбек|бармен/.test(role)) return "Бар";
  return null;
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function getSelectedValues(selectEl) {
  return Array.from(selectEl.selectedOptions).map((o) => o.value);
}

function getCheckedGroups() {
  return Array.from(document.querySelectorAll(".groupCheck:checked")).map((el) => el.value);
}

function fillMultiSelect(selectEl, values, selectedValues = []) {
  const selectedSet = new Set(selectedValues.length ? selectedValues : values);
  selectEl.innerHTML = "";
  values.forEach((v) => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    opt.selected = selectedSet.has(v);
    selectEl.appendChild(opt);
  });
}

function findHeaderIndex(header, candidates) {
  for (const name of candidates) {
    const idx = header.findIndex((h) => normalize(String(h)) === normalize(name));
    if (idx !== -1) return idx;
  }
  return -1;
}

function toNum(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const n = Number(String(v).replace(/\s+/g, "").replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

function monthLabel(isoDate) {
  return isoDate.slice(0, 7);
}

function formatMoney(value) {
  return Number(value || 0).toFixed(2);
}

function parseStaffWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
  if (!rows.length) throw new Error("Файл сотрудников пустой.");

  const header = rows[0].map((h) => String(h).trim());
  const fioIdx = header.indexOf("ФИО");
  const restaurantIdx = header.indexOf("Название подразделения");
  if (fioIdx === -1 || restaurantIdx === -1) {
    throw new Error("В файле сотрудников нужны колонки: ФИО и Название подразделения.");
  }

  const map = new Map();
  let conflicts = 0;
  const conflictKeys = new Set();

  for (let i = 1; i < rows.length; i += 1) {
    const fio = String(rows[i][fioIdx] || "").trim();
    const restaurant = String(rows[i][restaurantIdx] || "").trim();
    if (!fio || !restaurant) continue;
    const key = normalizeFio(fio);
    if (!key) continue;
    if (!map.has(key)) map.set(key, restaurant);
    else if (map.get(key) !== restaurant) {
      conflicts += 1;
      conflictKeys.add(key);
    }
  }

  return { map, conflicts, conflictKeys };
}

function parseAttendanceWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
  if (!rows.length) return [];

  const header = rows[0].map((h) => String(h).trim());
  const idx = {
    date: findHeaderIndex(header, ["Дата"]),
    time: findHeaderIndex(header, ["Время"]),
    source: findHeaderIndex(header, ["Источник"]),
    direction: findHeaderIndex(header, ["Направление"]),
    surname: findHeaderIndex(header, ["Фамилия"]),
    name: findHeaderIndex(header, ["Имя"]),
    middle: findHeaderIndex(header, ["Отчество"]),
    fio: findHeaderIndex(header, ["ФИО"]),
    role: findHeaderIndex(header, ["Должность"]),
    address: findHeaderIndex(header, ["Адрес"])
  };

  const required = ["date", "time", "source", "direction", "role", "address"];
  const missing = required.filter((k) => idx[k] === -1);
  if (missing.length) throw new Error(`Не найдены нужные колонки в файле проходной: ${missing.join(", ")}`);
  if (idx.fio === -1 && (idx.surname === -1 || idx.name === -1)) {
    throw new Error("Не найдены колонки ФИО или Фамилия+Имя в файле проходной.");
  }

  const parsed = [];
  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    if (String(row[idx.source] || "").trim() !== "Проходная") continue;
    const dateSerialDay = excelDateToSerialDay(row[idx.date]);
    if (!Number.isFinite(dateSerialDay)) continue;
    const timeSec = parseExcelTimeToSeconds(row[idx.time]);
    if (!Number.isFinite(timeSec)) continue;

    const operationalSerialDay = timeSec < DAY_CUTOFF_SECONDS ? dateSerialDay - 1 : dateSerialDay;
    const dateIso = serialDayToISO(operationalSerialDay);
    if (!dateIso) continue;

    const roleRaw = String(row[idx.role] || "").trim();
    const group = classifyRole(roleRaw);
    if (!group) continue;

    const person = idx.fio !== -1
      ? String(row[idx.fio] || "").trim()
      : [row[idx.surname], row[idx.name], row[idx.middle]].filter(Boolean).join(" ").trim();
    if (!person) continue;

    const direction = String(row[idx.direction] || "").trim();
    if (direction !== "Вход" && direction !== "Выход") continue;

    parsed.push({
      dateIso,
      absSec: dateSerialDay * 86400 + timeSec,
      person,
      personKey: normalizeFio(person),
      group,
      role: roleRaw || "Не указана",
      direction,
      restaurantFromGate: String(row[idx.address] || "").trim() || "Не указан"
    });
  }

  return parsed;
}

function parsePayrollWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
  if (!rows.length) throw new Error("Файл зарплаты пустой.");

  const header = rows[0].map((h) => String(h).trim());
  const nameIdx = findHeaderIndex(header, ["Название", "ФИО", "Сотрудник"]);
  const accrualIdx = findHeaderIndex(header, ["Начисления", "Оклад"]);
  const ndflIdx = findHeaderIndex(header, ["НДФЛ"]);
  if (nameIdx === -1 || accrualIdx === -1 || ndflIdx === -1) {
    throw new Error("В файле зарплаты нужны колонки: Название/ФИО, Начисления, НДФЛ.");
  }

  const map = new Map();
  rows.slice(1).forEach((row) => {
    const name = String(row[nameIdx] || "").trim();
    if (!name) return;
    const salary = toNum(row[accrualIdx]);
    const ndfl = toNum(row[ndflIdx]);
    if (!(salary || ndfl)) return;
    if (name.split(/\s+/).length < 2) return;

    const key = normalizeFio(name);
    map.set(key, { salary, ndfl });
  });

  return map;
}

function parsePremiumCSV(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.trim());
  if (!lines.length) return new Map();

  const delim = lines[0].includes(";") ? ";" : ",";
  const header = lines[0].split(delim).map((v) => normalize(v));
  const fioIdx = header.findIndex((h) => h === normalize("ФИО"));
  const premiumIdx = header.findIndex((h) => h === normalize("Премии") || h === normalize("Премия"));
  if (fioIdx === -1 || premiumIdx === -1) {
    throw new Error("В файле премий нужны колонки: ФИО и Премии.");
  }

  const map = new Map();
  for (let i = 1; i < lines.length; i += 1) {
    const parts = lines[i].split(delim).map((s) => s.trim().replace(/^"|"$/g, ""));
    const fio = parts[fioIdx] || "";
    if (!fio) continue;
    const key = normalizeFio(fio);
    const value = toNum(parts[premiumIdx]);
    map.set(key, (map.get(key) || 0) + value);
  }

  return map;
}

function extractFioFromPremiumPage(text) {
  const compact = String(text || "").replace(/\s+/g, " ").trim();
  const fioMatch = compact.match(/Табельный номер\s*(.*?)\s*\(\s*фамилия,\s*имя,\s*отчество\s*\)/i);
  if (fioMatch && fioMatch[1]) return fioMatch[1].replace(/\s+\d+$/, "").trim();

  const alt = compact.match(/о поощрении работника\s*Табельный номер\s*(.*?)\s*\(/i);
  if (alt && alt[1]) return alt[1].replace(/\s+\d+$/, "").trim();
  return "";
}

function extractAmountFromPremiumPage(text) {
  const money = text.match(/\d[\d ]*\.\d{2}/g);
  if (!money || !money.length) return 0;
  let value = toNum(money[money.length - 1]);
  if (/В сумме\s*Минус/i.test(text) || /\bВзыскание\b/i.test(text)) value = -Math.abs(value);
  return value;
}

async function parsePremiumPDFs(files) {
  if (!window.pdfjsLib) throw new Error("Библиотека PDF не загружена.");

  const map = new Map();
  let matched = 0;
  let processedPages = 0;
  for (const file of files) {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const doc = await window.pdfjsLib.getDocument({ data: bytes, disableWorker: true }).promise;

    for (let pageNum = 1; pageNum <= doc.numPages; pageNum += 1) {
      const page = await doc.getPage(pageNum);
      const content = await page.getTextContent();
      const text = content.items.map((i) => i.str).join(" ");
      processedPages += 1;

      const fio = extractFioFromPremiumPage(text);
      if (!fio) continue;
      const amount = extractAmountFromPremiumPage(text);
      if (!amount) continue;

      const key = normalizeFio(fio);
      map.set(key, (map.get(key) || 0) + amount);
      matched += 1;
    }
  }

  return { map, matched, processedPages };
}

function rebuildMappedRecords() {
  mappingStats = { matched: 0, total: baseRecords.length };
  mappedRecords = baseRecords.map((r) => {
    const mappedRestaurant = staffRestaurantMap.get(r.personKey);
    if (mappedRestaurant) mappingStats.matched += 1;
    return { ...r, restaurant: mappedRestaurant || "Не определен в списке сотрудников", hasConflict: staffConflictKeys.has(r.personKey) };
  });

  const prevDates = getSelectedValues(dateSelect);
  const prevRestaurants = getSelectedValues(restaurantSelect);
  const months = [...new Set(mappedRecords.map((r) => monthLabel(r.dateIso)))].sort();
  const restaurants = [...new Set(mappedRecords.map((r) => r.restaurant))].sort((a, b) => a.localeCompare(b, "ru"));

  fillMultiSelect(dateSelect, months, prevDates);
  fillMultiSelect(restaurantSelect, restaurants, prevRestaurants);
}

function calcWorkedSeconds(events) {
  const sorted = [...events].sort((a, b) => a.absSec - b.absSec);
  let total = 0;
  let inWork = false;
  let startSec = 0;
  sorted.forEach((e) => {
    if (e.direction === "Вход") {
      inWork = true;
      startSec = e.absSec;
      return;
    }
    if (e.direction === "Выход" && inWork && e.absSec >= startSec) {
      total += e.absSec - startSec;
      inWork = false;
    }
  });
  if (total === 0 && sorted.length >= 2) {
    const fallback = sorted[sorted.length - 1].absSec - sorted[0].absSec;
    if (fallback > 0) total = fallback;
  }
  return total;
}

function workedSecondsToShift(workedSeconds) {
  if (workedSeconds <= 0) return 0;
  return workedSeconds > 7 * 3600 ? 1 : 0.5;
}

function pickMainRole(roleCounts) {
  let bestRole = "Не указана";
  let bestCount = -1;
  roleCounts.forEach((count, role) => {
    if (count > bestCount) {
      bestRole = role;
      bestCount = count;
    }
  });
  return bestRole;
}

function calculate(records) {
  const selectedMonths = getSelectedValues(dateSelect);
  const selectedRestaurants = getSelectedValues(restaurantSelect);
  const selectedGroups = new Set(getCheckedGroups());

  const filtered = records.filter(
    (r) => selectedMonths.includes(monthLabel(r.dateIso)) && selectedRestaurants.includes(r.restaurant) && selectedGroups.has(r.group)
  );

  const personDay = new Map();
  filtered.forEach((r) => {
    const key = `${r.dateIso}||${r.restaurant}||${r.group}||${r.person}`;
    if (!personDay.has(key)) {
      personDay.set(key, {
        dateIso: r.dateIso,
        month: monthLabel(r.dateIso),
        restaurant: r.restaurant,
        group: r.group,
        person: r.person,
        personKey: r.personKey,
        events: [],
        roles: new Map()
      });
    }
    const item = personDay.get(key);
    item.events.push({ direction: r.direction, absSec: r.absSec });
    item.roles.set(r.role, (item.roles.get(r.role) || 0) + 1);
  });

  const personMonth = new Map();
  Array.from(personDay.values()).forEach((day) => {
    const shift = workedSecondsToShift(calcWorkedSeconds(day.events));
    if (!shift) return;

    const monthKey = `${day.month}||${day.restaurant}||${day.group}||${day.personKey}`;
    if (!personMonth.has(monthKey)) {
      personMonth.set(monthKey, {
        month: day.month,
        restaurant: day.restaurant,
        group: day.group,
        person: day.person,
        personKey: day.personKey,
        roleCounts: new Map(),
        shifts: 0
      });
    }

    const row = personMonth.get(monthKey);
    row.shifts += shift;
    day.roles.forEach((count, role) => {
      row.roleCounts.set(role, (row.roleCounts.get(role) || 0) + count);
    });
  });

  return Array.from(personMonth.values())
    .map((r) => {
      const salary = payrollByPerson.get(r.personKey) || { salary: 0, ndfl: 0 };
      const premium = premiumByPerson.get(r.personKey) || 0;
      return {
        month: r.month,
        restaurant: r.restaurant,
        group: r.group,
        person: r.person,
        role: pickMainRole(r.roleCounts),
        shifts: r.shifts,
        salary: salary.salary,
        ndfl: salary.ndfl,
        premium
      };
    })
    .sort((a, b) => {
      if (a.month !== b.month) return a.month.localeCompare(b.month);
      if (a.restaurant !== b.restaurant) return a.restaurant.localeCompare(b.restaurant, "ru");
      if (a.group !== b.group) return a.group.localeCompare(b.group, "ru");
      return a.person.localeCompare(b.person, "ru");
    });
}

function renderTable(rows) {
  tableBody.innerHTML = "";
  if (!rows.length) {
    summaryEl.textContent = "По выбранным фильтрам данных нет.";
    csvBtn.disabled = true;
    xlsxBtn.disabled = true;
    return;
  }

  let totalShifts = 0;
  let totalSalary = 0;
  let totalNdfl = 0;
  let totalPremium = 0;

  rows.forEach((r) => {
    totalShifts += r.shifts;
    totalSalary += r.salary;
    totalNdfl += r.ndfl;
    totalPremium += r.premium;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(r.month)}</td>
      <td>${escapeHtml(r.restaurant)}</td>
      <td>${escapeHtml(r.group)}</td>
      <td>${escapeHtml(r.person)}</td>
      <td>${escapeHtml(r.role)}</td>
      <td>${r.shifts}</td>
      <td>${formatMoney(r.salary)}</td>
      <td>${formatMoney(r.ndfl)}</td>
      <td>${formatMoney(r.premium)}</td>
    `;
    tableBody.appendChild(tr);
  });

  summaryEl.textContent = `Строк: ${rows.length}. Смен: ${totalShifts}. Начисления: ${formatMoney(totalSalary)}. НДФЛ: ${formatMoney(totalNdfl)}. Премии: ${formatMoney(totalPremium)}.`;
  csvBtn.disabled = false;
  xlsxBtn.disabled = false;
}

function toCSV(rows) {
  const head = ["Месяц", "Ресторан", "Категория персонала", "ФИО", "Должность", "Количество смен", "Начисления", "НДФЛ", "Премии"];
  const lines = [head.join(";")];
  rows.forEach((r) => {
    lines.push([
      r.month,
      `"${String(r.restaurant).replaceAll('"', '""')}"`,
      `"${String(r.group).replaceAll('"', '""')}"`,
      `"${String(r.person).replaceAll('"', '""')}"`,
      `"${String(r.role).replaceAll('"', '""')}"`,
      r.shifts,
      formatMoney(r.salary),
      formatMoney(r.ndfl),
      formatMoney(r.premium)
    ].join(";"));
  });
  return lines.join("\n");
}

function exportExcel(rows) {
  const wb = XLSX.utils.book_new();
  const aoa = [["Месяц", "Ресторан", "Категория персонала", "ФИО", "Должность", "Количество смен", "Начисления", "НДФЛ", "Премии"]];
  rows.forEach((r) => {
    aoa.push([r.month, r.restaurant, r.group, r.person, r.role, r.shifts, r.salary, r.ndfl, r.premium]);
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Отчет");
  XLSX.writeFile(wb, `зарплата_по_сотрудникам_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function refreshStatus() {
  if (!baseRecords.length) {
    statusEl.textContent = "Загрузите файл проходной.";
    return;
  }

  const staffLoaded = staffRestaurantMap.size > 0;
  const payrollLoaded = payrollByPerson.size > 0;
  const premiumLoaded = premiumByPerson.size > 0;

  const staffPart = staffLoaded
    ? ` Сотрудники: сопоставлено ${mappingStats.matched} из ${mappingStats.total}.${staffConflicts ? ` Конфликтов ФИО: ${staffConflicts}.` : ""}`
    : " Список сотрудников не загружен.";

  const payrollPart = payrollLoaded ? ` Начисления/НДФЛ загружены: ${payrollByPerson.size} ФИО.` : " Файл начислений не загружен.";
  let premiumPart = " Файл премий не загружен (будет 0).";
  if (premiumLoadInfo.loaded) {
    premiumPart = ` Премии: файлов ${premiumLoadInfo.files}, приказов найдено ${premiumLoadInfo.matched}, ФИО с суммами ${premiumByPerson.size}.`;
  } else if (premiumLoaded) {
    premiumPart = ` Премии загружены: ${premiumByPerson.size} ФИО.`;
  }

  statusEl.textContent = `Записей проходной: ${baseRecords.length}.${staffPart}${payrollPart}${premiumPart}`;
}

async function applyAttendanceFromBuffer(buf) {
  baseRecords = parseAttendanceWorkbook(buf);
  rebuildMappedRecords();
  lastResultRows = [];
  tableBody.innerHTML = "";
  summaryEl.textContent = "Выберите фильтры и нажмите «Рассчитать».";
  csvBtn.disabled = true;
  xlsxBtn.disabled = true;
  refreshStatus();
}

async function applyStaffFromBuffer(buf) {
  const staff = parseStaffWorkbook(buf);
  staffRestaurantMap = staff.map;
  staffConflicts = staff.conflicts;
  staffConflictKeys = staff.conflictKeys;
  if (baseRecords.length) rebuildMappedRecords();
  refreshStatus();
}

async function applyPayrollFromBuffer(buf) {
  payrollByPerson = parsePayrollWorkbook(buf);
  refreshStatus();
}

async function applyPremiumFromFiles(files) {
  const csvFiles = files.filter((f) => f.name.toLowerCase().endsWith(".csv"));
  const pdfFiles = files.filter((f) => f.name.toLowerCase().endsWith(".pdf"));
  const map = new Map();

  for (const csvFile of csvFiles) {
    const csvMap = parsePremiumCSV(await csvFile.text());
    csvMap.forEach((value, key) => map.set(key, (map.get(key) || 0) + value));
  }

  if (pdfFiles.length) {
    statusEl.textContent = "Обрабатываю PDF с премиями, это может занять время...";
    const pdfResult = await parsePremiumPDFs(pdfFiles);
    pdfResult.map.forEach((value, key) => map.set(key, (map.get(key) || 0) + value));
    premiumLoadInfo = { loaded: true, files: pdfFiles.length, matched: pdfResult.matched };
  } else {
    premiumLoadInfo = { loaded: csvFiles.length > 0, files: csvFiles.length, matched: map.size };
  }

  premiumByPerson = map;
  refreshStatus();
}

function detectWorkbookKindFromRows(rows) {
  if (!rows.length) return "unknown";
  const header = rows[0].map((h) => normalize(String(h)));

  const has = (name) => header.includes(normalize(name));
  const hasAll = (names) => names.every(has);

  if (hasAll(["Дата", "Время", "Источник", "Направление"])) return "attendance";
  if (hasAll(["ФИО", "Название подразделения"])) return "staff";
  if (hasAll(["Название", "Начисления", "НДФЛ"])) return "payroll";
  if (hasAll(["ФИО", "Премии"]) || hasAll(["ФИО", "Премия"])) return "premium_csv_like";
  return "unknown";
}

attendanceInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    await applyAttendanceFromBuffer(await file.arrayBuffer());
  } catch (err) {
    statusEl.textContent = `Ошибка файла проходной: ${err.message}`;
  }
});

staffInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    await applyStaffFromBuffer(await file.arrayBuffer());
  } catch (err) {
    statusEl.textContent = `Ошибка файла сотрудников: ${err.message}`;
  }
});

payrollInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    await applyPayrollFromBuffer(await file.arrayBuffer());
  } catch (err) {
    statusEl.textContent = `Ошибка файла начислений: ${err.message}`;
  }
});

premiumInput.addEventListener("change", async (e) => {
  const files = Array.from(e.target.files || []);
  if (!files.length) return;
  try {
    await applyPremiumFromFiles(files);
  } catch (err) {
    premiumLoadInfo = { loaded: false, files: 0, matched: 0 };
    statusEl.textContent = `Ошибка файла премий: ${err.message}`;
  }
});

allFilesInput?.addEventListener("change", async (e) => {
  const files = Array.from(e.target.files || []);
  if (!files.length) return;

  const premiumFiles = [];
  const unknownFiles = [];
  let loaded = { attendance: 0, staff: 0, payroll: 0, premium: 0 };

  statusEl.textContent = `Автоопределение ${files.length} файлов...`;

  for (const file of files) {
    const lower = file.name.toLowerCase();
    try {
      if (lower.endsWith(".pdf")) {
        premiumFiles.push(file);
        loaded.premium += 1;
        continue;
      }
      if (lower.endsWith(".csv")) {
        premiumFiles.push(file);
        loaded.premium += 1;
        continue;
      }
      if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
        const kind = detectWorkbookKindFromRows(rows);

        if (kind === "attendance") {
          await applyAttendanceFromBuffer(buf);
          loaded.attendance += 1;
        } else if (kind === "staff") {
          await applyStaffFromBuffer(buf);
          loaded.staff += 1;
        } else if (kind === "payroll") {
          await applyPayrollFromBuffer(buf);
          loaded.payroll += 1;
        } else if (kind === "premium_csv_like") {
          premiumFiles.push(file);
          loaded.premium += 1;
        } else {
          unknownFiles.push(file.name);
        }
        continue;
      }
      unknownFiles.push(file.name);
    } catch (err) {
      unknownFiles.push(`${file.name} (${err.message})`);
    }
  }

  if (premiumFiles.length) {
    try {
      await applyPremiumFromFiles(premiumFiles);
    } catch (err) {
      statusEl.textContent = `Ошибка обработки премий при автоимпорте: ${err.message}`;
      return;
    }
  }

  refreshStatus();
  const unknownPart = unknownFiles.length ? ` Не распознаны: ${unknownFiles.join(", ")}.` : "";
  statusEl.textContent = `Автозагрузка завершена. Проходная: ${loaded.attendance}, сотрудники: ${loaded.staff}, начисления: ${loaded.payroll}, премии-файлы: ${loaded.premium}.${unknownPart}`;
});

calcBtn.addEventListener("click", () => {
  if (!mappedRecords.length) {
    summaryEl.textContent = "Сначала загрузите файл проходной.";
    return;
  }
  if (!getCheckedGroups().length) {
    summaryEl.textContent = "Выберите хотя бы одну группу должностей.";
    return;
  }
  lastResultRows = calculate(mappedRecords);
  renderTable(lastResultRows);
});

csvBtn.addEventListener("click", () => {
  if (!lastResultRows.length) return;
  const csv = toCSV(lastResultRows);
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "зарплата_по_сотрудникам.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
});

xlsxBtn.addEventListener("click", () => {
  if (!lastResultRows.length) return;
  exportExcel(lastResultRows);
});
