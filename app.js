const fileInput = document.getElementById("fileInput");
const localFileSelect = document.getElementById("localFileSelect");
const calculateLocalBtn = document.getElementById("calculateLocalBtn");
const resultBody = document.getElementById("resultBody");
const totalRow = document.getElementById("totalRow");
const messages = document.getElementById("messages");

const REQUIRED_COLUMNS = {
  article: "Артикул",
  name: "Номенклатура",
  unit: "Ед. изм.",
  inStock: "В наличии",
  reserved: "В резерве",
  available: "Доступно",
  toSupply: "К обеспечению",
  deficit: "Дефицит",
  safetyStock: "Страховой запас",
};

initLocalFilesSelect();

fileInput.addEventListener("change", async (event) => {
  const files = Array.from(event.target.files || []);
  if (!files.length) {
    return;
  }

  if (typeof XLSX === "undefined") {
    setMessage(
      "Ошибка: библиотека чтения Excel не загрузилась. Обновите страницу и попробуйте снова."
    );
    return;
  }

  setMessage(`Обработка файлов: ${files.map((f) => f.name).join(", ")}`);
  try {
    const allRows = await parseFiles(files);
    renderTable(allRows);
    setMessage(`Готово. Загружено строк: ${allRows.length}`);
  } catch (error) {
    setMessage(`Ошибка: ${error.message}`);
  }
});

calculateLocalBtn.addEventListener("click", async () => {
  if (typeof XLSX === "undefined") {
    setMessage(
      "Ошибка: библиотека чтения Excel не загрузилась. Обновите страницу и попробуйте снова."
    );
    return;
  }

  const selected = Array.from(localFileSelect.selectedOptions).map((opt) => opt.value);
  if (!selected.length) {
    setMessage("Выберите хотя бы один файл в списке локальных файлов.");
    return;
  }

  setMessage(`Загрузка локальных файлов: ${selected.join(", ")}`);

  try {
    const allRows = [];
    for (const fileName of selected) {
      const response = await fetch(encodeURI(fileName));
      if (!response.ok) {
        throw new Error(`Не удалось загрузить локальный файл: ${fileName}`);
      }
      const buffer = await response.arrayBuffer();
      const rows = parseWorkbookBuffer(buffer, fileName);
      allRows.push(...rows);
    }
    renderTable(allRows);
    setMessage(`Готово. Загружено строк: ${allRows.length}`);
  } catch (error) {
    setMessage(
      `Ошибка: ${error.message}. Для локального списка используйте запуск через http://localhost (python3 -m http.server).`
    );
  }
});

async function parseFile(file) {
  const buffer = await file.arrayBuffer();
  return parseWorkbookBuffer(buffer, file.name);
}

function parseWorkbookBuffer(buffer, sourceName) {
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  const headerIndex = findHeaderIndex(raw);
  if (headerIndex === -1) {
    throw new Error(`В файле ${sourceName} не найдена строка заголовков.`);
  }

  const header = raw[headerIndex].map((cell) => normalizeText(cell));
  const columnIndexes = getColumnIndexes(header);

  validateRequiredColumns(sourceName, columnIndexes);

  const dataRows = raw.slice(headerIndex + 1);
  const parsed = [];

  for (const row of dataRows) {
    const article = getCell(row, columnIndexes.article);
    const name = getCell(row, columnIndexes.name);
    const unit = normalizeText(getCell(row, columnIndexes.unit)).toLowerCase();

    if (!isProductRow(article, name, unit)) {
      continue;
    }

    const factor = getUnitFactor(unit, name);
    parsed.push({
      source: sourceName,
      article: safeText(article),
      name: safeText(name),
      unit: safeText(unit || "не определено"),
      inStock: toSquareMeters(getCell(row, columnIndexes.inStock), factor),
      reserved: toSquareMeters(getCell(row, columnIndexes.reserved), factor),
      available: toSquareMeters(getCell(row, columnIndexes.available), factor),
      toSupply: toSquareMeters(getCell(row, columnIndexes.toSupply), factor),
      deficit: toSquareMeters(getCell(row, columnIndexes.deficit), factor),
      safetyStock: toSquareMeters(getCell(row, columnIndexes.safetyStock), factor),
    });
  }

  return parsed;
}

async function parseFiles(files) {
  const allRows = [];
  for (const file of files) {
    const rows = await parseFile(file);
    allRows.push(...rows);
  }
  return allRows;
}

async function initLocalFilesSelect() {
  try {
    const response = await fetch("./files-manifest.json");
    if (!response.ok) {
      throw new Error("Не удалось загрузить files-manifest.json");
    }

    const payload = await response.json();
    const fileNames = Array.isArray(payload.files) ? payload.files : [];
    localFileSelect.innerHTML = "";

    if (!fileNames.length) {
      localFileSelect.innerHTML = '<option value="">Нет файлов в манифесте</option>';
      return;
    }

    for (const fileName of fileNames) {
      const option = document.createElement("option");
      option.value = fileName;
      option.textContent = fileName;
      localFileSelect.appendChild(option);
    }
  } catch (_error) {
    localFileSelect.innerHTML =
      '<option value="">Не удалось загрузить список (запустите через http://localhost)</option>';
  }
}

function findHeaderIndex(rows) {
  return rows.findIndex((row) => {
    const normalized = row.map((cell) => normalizeText(cell));
    return normalized.includes("Артикул") && normalized.includes("Номенклатура");
  });
}

function getColumnIndexes(header) {
  return {
    article: header.findIndex((h) => h === REQUIRED_COLUMNS.article),
    name: header.findIndex((h) => h === REQUIRED_COLUMNS.name),
    unit: header.findIndex((h) => h === REQUIRED_COLUMNS.unit),
    inStock: header.findIndex((h) => h === REQUIRED_COLUMNS.inStock),
    reserved: header.findIndex((h) => h === REQUIRED_COLUMNS.reserved),
    available: header.findLastIndex
      ? header.findLastIndex((h) => h === REQUIRED_COLUMNS.available)
      : findLastIndexFallback(header, REQUIRED_COLUMNS.available),
    toSupply: header.findIndex((h) => h === REQUIRED_COLUMNS.toSupply),
    deficit: header.findIndex((h) => h === REQUIRED_COLUMNS.deficit),
    safetyStock: header.findIndex((h) => h === REQUIRED_COLUMNS.safetyStock),
  };
}

function findLastIndexFallback(items, value) {
  for (let i = items.length - 1; i >= 0; i -= 1) {
    if (items[i] === value) {
      return i;
    }
  }
  return -1;
}

function validateRequiredColumns(fileName, indexes) {
  const missing = Object.entries(indexes)
    .filter(([, index]) => index < 0)
    .map(([key]) => REQUIRED_COLUMNS[key]);

  if (missing.length > 0) {
    throw new Error(
      `Файл ${fileName}: отсутствуют обязательные колонки: ${missing.join(", ")}`
    );
  }
}

function isProductRow(article, name, unit) {
  const nameText = safeText(name);
  const articleText = safeText(article);
  if (!nameText) {
    return false;
  }

  const isLikelyServiceRow =
    nameText === "Партнёр" ||
    nameText === "Номенклатура" ||
    nameText.includes("Заказ") ||
    nameText.includes("Итого");

  if (isLikelyServiceRow) {
    return false;
  }

  const hasUnit = unit.includes("пог") || unit.includes("лист");
  const hasArticleOrMaterialName = /\d/.test(articleText) || /sylomer|sylodyn/i.test(nameText);
  return hasUnit && hasArticleOrMaterialName;
}

function getUnitFactor(unit, name) {
  if (unit.includes("пог")) {
    return 1.5;
  }
  if (unit.includes("лист")) {
    return 1.2;
  }

  const nameText = safeText(name).toLowerCase();
  if (nameText.includes("1200 х 1500") || nameText.includes("1200x1500")) {
    return 1.5;
  }
  if (nameText.includes("800 х 1500") || nameText.includes("800x1500")) {
    return 1.2;
  }
  return 1;
}

function toSquareMeters(value, factor) {
  const number = toNumber(value);
  return number * factor;
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  const text = String(value ?? "")
    .replace(/\s+/g, "")
    .replace(",", ".");

  const parsed = Number.parseFloat(text);
  return Number.isFinite(parsed) ? parsed : 0;
}

function getCell(row, index) {
  if (index < 0 || index >= row.length) {
    return "";
  }
  return row[index];
}

function safeText(value) {
  return String(value ?? "").trim();
}

function normalizeText(value) {
  return safeText(value).replace(/\s+/g, " ");
}

function formatNumber(value) {
  return value.toLocaleString("ru-RU", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function renderTable(rows) {
  if (!rows.length) {
    resultBody.innerHTML =
      '<tr><td colspan="8" class="placeholder">Не найдено строк с материалами.</td></tr>';
    updateTotalRow([0, 0, 0, 0, 0, 0]);
    return;
  }

  const sortedRows = [...rows].sort((a, b) => {
    const aHasDeficit = a.deficit > 0;
    const bHasDeficit = b.deficit > 0;
    if (aHasDeficit === bHasDeficit) {
      return 0;
    }
    return aHasDeficit ? -1 : 1;
  });

  resultBody.innerHTML = sortedRows
    .map(
      (row) => `
      <tr class="${row.deficit > 0 ? "deficit-row" : ""}">
        <td>${escapeHtml(row.article)}</td>
        <td>${escapeHtml(row.name)}</td>
        <td>${formatNumber(row.inStock)}</td>
        <td>${formatNumber(row.reserved)}</td>
        <td class="available-col">${formatNumber(row.available)}</td>
        <td>${formatNumber(row.toSupply)}</td>
        <td class="${row.deficit > 0 ? "deficit-positive" : ""}">${formatNumber(row.deficit)}</td>
        <td>${formatNumber(row.safetyStock)}</td>
      </tr>
    `
    )
    .join("");

  const totals = sortedRows.reduce(
    (acc, row) => {
      acc[0] += row.inStock;
      acc[1] += row.reserved;
      acc[2] += row.available;
      acc[3] += row.toSupply;
      acc[4] += row.deficit;
      acc[5] += row.safetyStock;
      return acc;
    },
    [0, 0, 0, 0, 0, 0]
  );

  updateTotalRow(totals);
}
function updateTotalRow(totals) {
  const cells = totalRow.querySelectorAll("td");
  for (let i = 0; i < totals.length; i += 1) {
    cells[i + 1].textContent = formatNumber(totals[i]);
  }

  const deficitCell = cells[5];
  deficitCell.classList.toggle("deficit-positive", totals[4] > 0);
}

function setMessage(text) {
  messages.textContent = text;
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
