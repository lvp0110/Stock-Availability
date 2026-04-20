const fileInput = document.getElementById("fileInput");
const localFileSelect = document.getElementById("localFileSelect");
const calculateLocalBtn = document.getElementById("calculateLocalBtn");
const resultBody = document.getElementById("resultBody");
const totalRow = document.getElementById("totalRow");
const messages = document.getElementById("messages");
const messagesCard = messages?.closest(".card");
const uploadCardToggle = document.getElementById("uploadCardToggle");
const uploadCardContent = document.getElementById("uploadCardContent");
const uploadCardIndicator = uploadCardToggle?.querySelector(".toggle-indicator");

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
initUploadCardToggle();
toggleMessagesCardVisibility("");

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
    setMessage(`Загружено строк: ${allRows.length}`);
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
  const parsed = [];
  let foundHeader = false;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const headerIndexes = findHeaderIndexes(raw);

    for (let i = 0; i < headerIndexes.length; i += 1) {
      const headerIndex = headerIndexes[i];
      const nextHeaderIndex = headerIndexes[i + 1] ?? raw.length;
      const header = raw[headerIndex].map((cell) => normalizeText(cell));
      const columnIndexes = getColumnIndexes(header);

      if (!hasAllRequiredColumns(columnIndexes)) {
        continue;
      }

      foundHeader = true;
      const dataRows = raw.slice(headerIndex + 1, nextHeaderIndex);
      parsed.push(...parseDataRows(dataRows, columnIndexes, sourceName, sheetName));
    }
  }

  if (!foundHeader) {
    throw new Error(
      `В файле ${sourceName} не найдены таблицы с обязательными колонками (${Object.values(
        REQUIRED_COLUMNS
      ).join(", ")}).`
    );
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

function initUploadCardToggle() {
  if (!uploadCardToggle || !uploadCardContent) {
    return;
  }

  updateUploadCardIndicator(false);

  uploadCardToggle.addEventListener("click", () => {
    const isCollapsed = uploadCardContent.classList.toggle("is-collapsed");
    uploadCardToggle.setAttribute("aria-expanded", String(!isCollapsed));
    updateUploadCardIndicator(!isCollapsed);
  });
}

function updateUploadCardIndicator(isExpanded) {
  if (!uploadCardIndicator) {
    return;
  }
  uploadCardIndicator.textContent = isExpanded ? "▾" : "▸";
}

function findHeaderIndexes(rows) {
  const indexes = [];
  rows.forEach((row, index) => {
    const normalized = row.map((cell) => normalizeText(cell));
    if (normalized.includes("Артикул") && normalized.includes("Номенклатура")) {
      indexes.push(index);
    }
  });
  return indexes;
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
    toOrder: header.findIndex((h) => h === "К заказу"),
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

function hasAllRequiredColumns(indexes) {
  return Object.keys(REQUIRED_COLUMNS).every((key) => indexes[key] >= 0);
}

function parseDataRows(dataRows, columnIndexes, sourceName, sheetName) {
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
      source: `${sourceName} / ${sheetName}`,
      article: safeText(article),
      name: safeText(name),
      unit: safeText(unit || "не определено"),
      inStock: toSquareMeters(getCell(row, columnIndexes.inStock), factor),
      reserved: toSquareMeters(getCell(row, columnIndexes.reserved), factor),
      available: toSquareMeters(getCell(row, columnIndexes.available), factor),
      toSupply: toSquareMeters(getCell(row, columnIndexes.toSupply), factor),
      toOrder: toSquareMeters(getCell(row, columnIndexes.toOrder), factor),
      deficit: toSquareMeters(getCell(row, columnIndexes.deficit), factor),
      safetyStock: toSquareMeters(getCell(row, columnIndexes.safetyStock), factor),
    });
  }
  return parsed;
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

function getToOrderValue(row) {
  return getToOrderDetails(row).value;
}

function getToOrderDetails(row) {
  const roundUpToInteger = (value) => Math.ceil(value);

  if (row.deficit <= 0) {
    return {
      value: roundUpToInteger(row.toOrder),
      formula: "Округление вверх: ceil(К заказу из файла)",
    };
  }

  const thickness = detectMaterialThicknessMm(row);
  if (thickness === 12,5) {
    return {
      value: roundUpToInteger((row.deficit / 76) * 1.2),
      formula: "Округление вверх: ceil((Дефицит / 76) * 1.2)",
    };
  }
  if (thickness === 25) {
    return {
      value: roundUpToInteger((row.deficit / 38) * 1.2),
      formula: "Округление вверх: ceil((Дефицит / 38) * 1.2)",
    };
  }

  return {
    value: roundUpToInteger(row.toOrder),
    formula: "Толщина не распознана: округление вверх исходного К заказу",
  };
}

function detectMaterialThicknessMm(row) {
  const text = `${safeText(row.name)} ${safeText(row.article)}`
    .toLowerCase()
    .replace(",", ".");
  const match = text.match(/(?:^|[^\d])(12\.5|25)(?:\s*мм|\b)/);
  if (!match) {
    return null;
  }
  const parsed = Number.parseFloat(match[1]);
  return Number.isFinite(parsed) ? parsed : null;
}

function renderTable(rows) {
  if (!rows.length) {
    resultBody.innerHTML =
      '<tr><td colspan="10" class="placeholder">Не найдено строк с материалами.</td></tr>';
    updateTotalRow(totalRow, [0, 0, 0, 0, 0, 0, 0]);
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
      (row) => {
        const toOrderDetails = getToOrderDetails(row);
        const materialColor = getMaterialColorHex(row);
        const colorCellStyle = materialColor
          ? ` style="--color-cell-bg:${materialColor};background-color:${materialColor}"`
          : "";

        return `
      <tr class="${row.deficit > 0 ? "deficit-row" : ""}">
        <td class="color-cell"${colorCellStyle}></td>
        <td>${escapeHtml(row.article)}</td>
        <td>${escapeHtml(row.name)}</td>
        <td>${formatNumber(row.inStock)}</td>
        <td>${formatNumber(row.reserved)}</td>
        <td class="available-col">${formatNumber(row.available)}</td>
        <td>${formatNumber(row.toSupply)}</td>
        <td class="${row.deficit > 0 ? "deficit-positive" : ""}">${formatNumber(row.deficit)}</td>
        <td>${formatNumber(row.safetyStock)}</td>
        <td title="${escapeHtml(toOrderDetails.formula)}">${formatNumber(toOrderDetails.value)}</td>
      </tr>
    `;
      }
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
      acc[6] += getToOrderValue(row);
      return acc;
    },
    [0, 0, 0, 0, 0, 0, 0]
  );

  updateTotalRow(totalRow, totals);
}
function updateTotalRow(targetRow, totals) {
  const cells = targetRow.querySelectorAll("td");
  for (let i = 0; i < totals.length; i += 1) {
    cells[i + 1].textContent = formatNumber(totals[i]);
  }

  const deficitCell = cells[5];
  deficitCell.classList.toggle("deficit-positive", totals[4] > 0);
}

function setMessage(text) {
  const nextText = String(text ?? "").trim();
  messages.textContent = nextText;
  toggleMessagesCardVisibility(nextText);
}

function toggleMessagesCardVisibility(text) {
  if (!messagesCard) {
    return;
  }
  messagesCard.hidden = !text;
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getMaterialColorHex(row) {
  const text = `${safeText(row.name)} ${safeText(row.article)}`.toLowerCase();

  const colorByKeyword = [
    // Match explicit color words from source description first.
    { re: /yellow|желт/i, hex: "#facc15" },
    { re: /brown|коричнев/i, hex: "#8b5a2b" },
    { re: /orange|оранж/i, hex: "#f97316" },
    { re: /red|красн/i, hex: "#dc2626" },
    { re: /blue|син/i, hex: "#2563eb" },
    { re: /pink|розов/i, hex: "#ec4899" },
    { re: /green|зелен/i, hex: "#16a34a" },
    { re: /grey|gray|сер/i, hex: "#6b7280" },
    { re: /black|черн/i, hex: "#111827" },
    // Fallback by material code if color words are missing.
    { re: /\b(sr ?11|ср ?11)\b/, hex: "#facc15" },
    { re: /\b(sr ?110|ср ?110)\b/, hex: "#8b5a2b" },
    { re: /\b(sr ?18|ср ?18)\b/, hex: "#f97316" },
    { re: /\b(sr ?220|ср ?220)\b/, hex: "#dc2626" },
    { re: /\b(sr ?28|ср ?28)\b/, hex: "#2563eb" },
    { re: /\b(sr ?42|ср ?42)\b/, hex: "#ec4899" },
    { re: /\b(sr ?55|ср ?55)\b/, hex: "#16a34a" },
    { re: /\b(sr ?330|ср ?330)\b/, hex: "#111827" },
    { re: /\b(sr ?450|ср ?450)\b/, hex: "#6b7280" },
    { re: /\b(sd\b|sylodyn)\b|purple|фиолет/i, hex: "#7c3aed" },
  ];

  const matched = colorByKeyword.find((item) => item.re.test(text));
  return matched?.hex ?? "";
}
