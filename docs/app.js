"use strict";

const fileInput = document.querySelector("#product-file");
const status = document.querySelector("#status");
const download = document.querySelector("#download");
const reportInput = document.querySelector("#report-file");
const reportStatus = document.querySelector("#report-status");
const reportDownload = document.querySelector("#report-download");
const dropStatus = document.querySelector("#drop-status");
const shortageList = document.querySelector("#shortage-list");
const shortageCount = document.querySelector("#shortage-count");
const shortageRows = document.querySelector("#shortage-rows");
const order = new Intl.Collator("zh-Hant-TW-u-co-stroke");
const xlsxType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const sourceHeaders = {
  AH: "[29]商品款式",
  X: "[19]商品款式",
  Z: "[21]商品款式",
  AK: "[32]商品款式",
};

async function readXml(zip, path) {
  const entry = zip.file(path);
  if (!entry) throw new Error("Excel 檔案缺少必要內容。");
  const contents = await entry.async("text");
  const xml = new DOMParser().parseFromString(contents, "application/xml");
  if (xml.querySelector("parsererror")) throw new Error("Excel 檔案內容無法讀取。");
  return xml;
}

function relationshipPath(relationship) {
  const base = "https://xlsx.invalid/xl/workbook.xml";
  const path = new URL(relationship.getAttribute("Target"), base);
  if (path.origin !== new URL(base).origin || relationship.getAttribute("TargetMode") === "External") {
    throw new Error("Excel 檔案的工作表路徑無法讀取。");
  }
  return decodeURIComponent(path.pathname.slice(1));
}

function excelText(node) {
  if (!node) return "";
  return Array.from(node.querySelectorAll(":scope > t, :scope > r > t"), (text) => text.textContent)
    .join("")
    .replace(/_x([\da-f]{4})_/gi, (_, code) => String.fromCharCode(parseInt(code, 16)));
}

async function openWorkbook(file) {
  if (!/\.xlsx$/i.test(file.name)) throw new Error("請選擇 .xlsx 檔案。");
  const zip = await JSZip.loadAsync(await file.arrayBuffer()).catch(() => {
    throw new Error("無法開啟檔案，請選擇有效且未加密的 .xlsx 檔案。");
  });
  const [workbook, relationshipXml] = await Promise.all([
    readXml(zip, "xl/workbook.xml"),
    readXml(zip, "xl/_rels/workbook.xml.rels"),
  ]);
  return { zip, workbook, relationshipXml };
}

async function openWorksheet(file, sheetName, source) {
  const { zip, workbook, relationshipXml } = source ?? await openWorkbook(file);
  const sheet = workbook.querySelector(`sheets > sheet[name="${sheetName}"]`);
  if (!sheet) throw new Error(`找不到「${sheetName}」工作表。`);
  const relationships = Array.from(relationshipXml.querySelectorAll("Relationship"));
  const sheetId = sheet.getAttributeNS(relationshipNamespace, "id");
  const sheetLink = relationships.find((link) => link.getAttribute("Id") === sheetId);
  if (!sheetLink) throw new Error(`無法讀取「${sheetName}」工作表。`);
  const sheetPath = relationshipPath(sheetLink);
  const stringsLink = relationships.find((link) => link.getAttribute("Type")?.endsWith("/sharedStrings"));
  // ponytail: parse one worksheet in memory. For exports that stall the UI, move parsing to a worker.
  const [data, sharedStrings] = await Promise.all([
    readXml(zip, sheetPath),
    stringsLink && readXml(zip, relationshipPath(stringsLink)),
  ]);
  const strings = sharedStrings ? Array.from(sharedStrings.querySelectorAll("sst > si"), excelText) : [];

  function textAt(row, column) {
    const cell = row?.querySelector(`c[r="${column}${row.getAttribute("r")}"]`);
    if (!cell || cell.getAttribute("t") === "e") return "";
    const value = cell.querySelector("v")?.textContent ?? "";
    switch (cell.getAttribute("t")) {
      case "s": return strings[value] ?? "";
      case "inlineStr": return excelText(cell.querySelector("is"));
      default: return value;
    }
  }

  return { zip, data, sheetPath, relationships, textAt };
}

function isProductHeader(header, textAt) {
  return header?.getAttribute("r") === "1" && Object.entries(sourceHeaders).every(([col, prefix]) =>
    textAt(header, col).trim().startsWith(prefix)
  );
}

function isReportHeader(header, textAt) {
  return header?.getAttribute("r") === "1" && textAt(header, "M").trim() === "商品款式" &&
    textAt(header, "N").trim() === "SKU" && textAt(header, "P").trim() === "數量";
}

async function identifyFile(file) {
  const source = await openWorkbook(file);
  const matches = [];
  for (const [type, sheet, validate] of [
    ["product", "Products", isProductHeader],
    ["report", "訂單明細", isReportHeader],
  ]) {
    if (!source.workbook.querySelector(`sheets > sheet[name="${sheet}"]`)) continue;
    const { data, textAt } = await openWorksheet(file, sheet, source);
    if (validate(data.querySelector("sheetData > row"), textAt)) matches.push(type);
  }
  if (matches.length !== 1) {
    throw new Error(matches.length ? "同時符合兩種報表格式，無法自動分類。" : "工作表及欄位不符合商品列表或訂單報表格式。");
  }
  return matches[0];
}

async function readProducts(file) {
  const { data, textAt } = await openWorksheet(file, "Products");
  const [header, ...rows] = data.querySelectorAll("sheetData > row");
  if (!isProductHeader(header, textAt)) {
    throw new Error("欄位與商品列表格式不符，請確認 AH、X、Z 與 AK 欄位。");
  }

  const result = [];
  for (const row of rows) {
    const stock = Number(textAt(row, "AK"));
    if (Number.isFinite(stock) && stock < 0) result.push([textAt(row, "AH"), textAt(row, "X"), textAt(row, "Z"), stock]);
  }
  return result.sort((a, b) => order.compare(a[1], b[1]) || order.compare(a[2], b[2]) || a[3] - b[3]);
}

function productKey(sku, variant) {
  return JSON.stringify([sku.trim(), variant.trim().replace(/\s*\/\s*/g, "/")]);
}

async function markOrderReport(file, shortages) {
  const { zip, data, sheetPath, relationships, textAt } = await openWorksheet(file, "訂單明細");
  const [header, ...rows] = data.querySelectorAll("sheetData > row");
  if (!isReportHeader(header, textAt)) {
    throw new Error("訂單明細的 M、N、P 欄須為商品款式、SKU、數量。");
  }

  const remaining = new Map();
  for (const [sku, color, size, stock] of shortages) {
    const variant = [color.trim(), size.trim()].filter(Boolean).join("/");
    const key = productKey(sku, variant);
    remaining.set(key, (remaining.get(key) ?? 0) - stock);
  }

  const required = new Map(remaining);
  const marked = [];
  let markedQuantity = 0;
  for (const row of rows.reverse()) {
    const sku = textAt(row, "N").trim();
    if (!sku) continue;
    const key = productKey(sku, textAt(row, "M"));
    const needed = remaining.get(key);
    if (!(needed > 0)) continue;
    const quantity = Number(textAt(row, "P"));
    if (Number.isFinite(quantity) && quantity > 0) {
      marked.push(row);
      markedQuantity += quantity;
      remaining.set(key, needed - quantity);
    }
  }

  if (marked.length) {
    const stylesLink = relationships.find((link) => link.getAttribute("Type")?.endsWith("/styles"));
    if (!stylesLink) throw new Error("訂單報表缺少儲存格格式，無法標綠。");
    const stylesPath = relationshipPath(stylesLink);
    const styles = await readXml(zip, stylesPath);
    const fills = styles.querySelector("styleSheet > fills");
    const formats = styles.querySelector("styleSheet > cellXfs");
    if (!fills || !formats?.firstElementChild) throw new Error("訂單報表的儲存格格式無法讀取。");

    const fillId = fills.children.length;
    const styleNamespace = styles.documentElement.namespaceURI;
    const fill = styles.createElementNS(styleNamespace, "fill");
    const pattern = styles.createElementNS(styleNamespace, "patternFill");
    const foreground = styles.createElementNS(styleNamespace, "fgColor");
    pattern.setAttribute("patternType", "solid");
    foreground.setAttribute("rgb", "FF4EA72E");
    pattern.append(foreground);
    fill.append(pattern);
    fills.append(fill);
    fills.setAttribute("count", fills.children.length);

    const originalFormats = [...formats.children];
    const coloredStyles = new Map();
    const columnOf = (cell) => cell.getAttribute("r").replace(/\d+$/, "");
    const columns = Array.from(header.querySelectorAll(":scope > c"), columnOf);
    for (const row of marked) {
      const cells = new Map(Array.from(row.querySelectorAll(":scope > c"), (cell) => [columnOf(cell), cell]));
      for (const column of columns) {
        if (!cells.has(column)) {
          const blank = data.createElementNS(row.namespaceURI, "c");
          blank.setAttribute("r", column + row.getAttribute("r"));
          cells.set(column, blank);
        }
        const cell = cells.get(column);
        const sourceStyle = Number(cell.getAttribute("s") ?? row.getAttribute("s") ?? 0);
        if (!coloredStyles.has(sourceStyle)) {
          const format = originalFormats[sourceStyle]?.cloneNode(true);
          if (!format) throw new Error("訂單報表含有無法讀取的儲存格格式。");
          format.setAttribute("fillId", fillId);
          format.setAttribute("applyFill", "1");
          coloredStyles.set(sourceStyle, formats.children.length);
          formats.append(format);
        }
        cell.setAttribute("s", coloredStyles.get(sourceStyle));
      }
      const extension = row.querySelector(":scope > extLst");
      for (const [, cell] of [...cells].sort(([a], [b]) => a.length - b.length || a.localeCompare(b, "en"))) {
        row.insertBefore(cell, extension);
      }
    }
    formats.setAttribute("count", formats.children.length);
    const serializer = new XMLSerializer();
    zip.file(sheetPath, serializer.serializeToString(data));
    zip.file(stylesPath, serializer.serializeToString(styles));
  }

  const unresolvedItems = [...remaining].filter(([, needed]) => needed > 0).map(([key, needed]) => {
    const [sku, variant] = JSON.parse(key);
    const demand = required.get(key);
    return { sku, variant, demand, ordered: demand - needed, missing: needed };
  });
  const unresolved = unresolvedItems.length;
  const extra = [...remaining.values()].reduce((sum, needed) => sum + Math.max(0, -needed), 0);
  const blob = await zip.generateAsync({ type: "blob", mimeType: xlsxType, compression: "DEFLATE" });
  return { blob, marked: marked.length, markedQuantity, unresolved, extra, unresolvedItems };
}

function escapeExcelText(value) {
  return value.replace(/_x[\da-f]{4}_|[\u0000-\u0008\u000b\u000c\u000d\u000e-\u001f\ufffe\uffff]|[&<>]/gi, (text) => {
    if (text.length > 1) return `_x005f_${text.slice(1)}`;
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;" }[text] ?? `_x${text.charCodeAt(0).toString(16).padStart(4, "0")}_`;
  });
}

async function createWorkbook(rows) {
  const response = await fetch("./template.xlsx");
  if (!response.ok) throw new Error("無法讀取匯出格式，請重新載入網頁後再試。");
  const zip = await JSZip.loadAsync(await response.arrayBuffer());
  const sheetPath = "xl/worksheets/sheet1.xml";
  const sheet = await zip.file(sheetPath).async("text");
  const records = rows.map(([sku, color, size, stock], index) => {
    const row = index + 2;
    const cells = [sku, color, size].map((text, col) => {
      const reference = "ABC"[col] + row;
      return `<c r="${reference}" s="1" t="inlineStr"><is><t xml:space="preserve">${escapeExcelText(text)}</t></is></c>`;
    });
    return `<row r="${row}" spans="1:4" x14ac:dyDescent="0.2">${cells.join("")}<c r="D${row}" s="2"><v>${stock}</v></c></row>`;
  });
  zip.file(sheetPath, sheet.replace('ref="A1:D1"', `ref="A1:D${rows.length + 1}"`)
    .replace("</sheetData>", () => `${records.join("")}</sheetData>`));
  return zip.generateAsync({ type: "blob", mimeType: xlsxType, compression: "DEFLATE" });
}

function clearDownload(link) {
  if (link.hasAttribute("href")) URL.revokeObjectURL(link.href);
  link.removeAttribute("href");
  link.hidden = true;
}

function showOrderShortages(items) {
  const rows = items.map(({ sku, variant, demand, ordered, missing }) => {
    const row = document.createElement("tr");
    for (const value of [sku, variant, demand, ordered, missing]) {
      const cell = document.createElement("td");
      cell.textContent = typeof value === "number" ? value.toLocaleString("zh-TW") : value;
      row.append(cell);
    }
    return row;
  });
  shortageRows.replaceChildren(...rows);
  shortageCount.textContent = `${rows.length.toLocaleString("zh-TW")} 個品項`;
  shortageList.hidden = !rows.length;
}

let selection = 0;
let reportSelection = 0;
let shortages = null;

async function updateOrderReport() {
  const current = ++reportSelection;
  clearDownload(reportDownload);
  showOrderShortages([]);
  reportStatus.removeAttribute("data-error");
  if (shortages === null) {
    reportStatus.textContent = reportInput.files[0] ? "已選擇訂單報表，等待商品列表。" : "請先選擇商品列表，或先選擇訂單報表。";
    return;
  }
  const file = reportInput.files[0];
  if (!file) {
    reportStatus.textContent = "選擇訂單報表後，系統會依負庫存從表尾往上累計數量並標綠。";
    return;
  }
  reportStatus.textContent = "正在標記訂單報表…";
  try {
    const { blob, markedQuantity, unresolvedItems } = await markOrderReport(file, shortages);
    if (current !== reportSelection) return;
    reportStatus.textContent = `已標綠 ${markedQuantity.toLocaleString("zh-TW")} 件。`;
    showOrderShortages(unresolvedItems);
    reportDownload.href = URL.createObjectURL(blob);
    reportDownload.download = `${file.name.slice(0, -5)}_負庫存標記.xlsx`;
    reportDownload.hidden = false;
  } catch (error) {
    if (current !== reportSelection) return;
    reportStatus.setAttribute("data-error", "");
    reportStatus.textContent = error.message || "處理失敗，請重新選擇訂單報表。";
  }
}

async function updateProducts() {
  const current = ++selection;
  clearDownload(download);
  shortages = null;
  void updateOrderReport();
  status.removeAttribute("data-error");
  const file = fileInput.files[0];
  if (!file) {
    status.textContent = "選擇商品列表後即可下載結果。";
    return;
  }

  status.textContent = "正在讀取商品列表…";
  try {
    const rows = await readProducts(file);
    if (current !== selection) return;
    const result = await createWorkbook(rows);
    if (current !== selection) return;
    const totalShortage = rows.reduce((sum, row) => sum - row[3], 0);
    status.textContent = rows.length ? `找到 ${rows.length.toLocaleString("zh-TW")} 筆負庫存，共缺貨 ${totalShortage.toLocaleString("zh-TW")} 件。` : "沒有負庫存，可下載只有欄名的 Excel。";
    download.href = URL.createObjectURL(result);
    download.download = `${file.name.slice(0, -5)}_負庫存.xlsx`;
    download.hidden = false;
    shortages = rows;
    void updateOrderReport();
  } catch (error) {
    if (current !== selection) return;
    status.setAttribute("data-error", "");
    status.textContent = error.message || "轉檔失敗，請重新選擇商品列表。";
  }
}

let dropSelection = 0;
let dragDepth = 0;

function setDropStatus(message, error = false) {
  dropStatus.textContent = message;
  dropStatus.hidden = !message;
  dropStatus.toggleAttribute("data-error", error);
}

function cancelPendingDrop() {
  ++dropSelection;
  setDropStatus("");
}

fileInput.addEventListener("change", () => {
  cancelPendingDrop();
  void updateProducts();
});
reportInput.addEventListener("change", () => {
  cancelPendingDrop();
  void updateOrderReport();
});

function hasFiles(event) {
  return Array.from(event.dataTransfer?.types ?? []).includes("Files");
}

function assignFile(input, file) {
  const transfer = new DataTransfer();
  transfer.items.add(file);
  input.files = transfer.files;
}

async function acceptDroppedFiles(files) {
  const current = ++dropSelection;
  setDropStatus("正在辨識檔案格式…");
  try {
    if (files.length > 2) throw new Error("一次最多拖入一份商品列表及一份訂單報表。");
    const identified = await Promise.all(files.map(async (file) => {
      try {
        return { file, type: await identifyFile(file) };
      } catch (error) {
        throw new Error(file.name + "：" + error.message);
      }
    }));
    if (current !== dropSelection) return;
    const product = identified.find(item => item.type === "product");
    const report = identified.find(item => item.type === "report");
    if (identified.length === 2 && (!product || !report)) {
      throw new Error("拖入了兩份相同類型的檔案，請一次選擇一份。");
    }
    // Assign both inputs first so product processing automatically picks up the report.
    if (product) assignFile(fileInput, product.file);
    if (report) assignFile(reportInput, report.file);
    setDropStatus("");
    if (product) void updateProducts();
    else if (report) void updateOrderReport();
  } catch (error) {
    if (current !== dropSelection) return;
    setDropStatus(error.message || "無法辨識檔案。", true);
  }
}

window.addEventListener("dragenter", event => {
  if (!hasFiles(event)) return;
  event.preventDefault();
  ++dragDepth;
  document.body.classList.add("is-dragging");
});
window.addEventListener("dragover", event => {
  if (!hasFiles(event)) return;
  event.preventDefault();
  event.dataTransfer.dropEffect = "copy";
  document.body.classList.add("is-dragging");
});
window.addEventListener("dragleave", () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if (!dragDepth) document.body.classList.remove("is-dragging");
});
window.addEventListener("drop", event => {
  if (!hasFiles(event) && !event.dataTransfer?.files.length) return;
  event.preventDefault();
  dragDepth = 0;
  document.body.classList.remove("is-dragging");
  const files = Array.from(event.dataTransfer.files);
  if (files.length) void acceptDroppedFiles(files);
});
window.addEventListener("blur", () => {
  dragDepth = 0;
  document.body.classList.remove("is-dragging");
});
