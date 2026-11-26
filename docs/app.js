"use strict";

const fileInput = document.querySelector("#product-file");
const status = document.querySelector("#status");
const download = document.querySelector("#download");
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
  return Array.from(node.querySelectorAll(":scope > t, :scope > r > t"), (text) => text.textContent)
    .join("")
    .replace(/_x([\da-f]{4})_/gi, (_, code) => String.fromCharCode(parseInt(code, 16)));
}

async function readProducts(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer()).catch(() => {
    throw new Error("無法開啟檔案，請選擇有效且未加密的 .xlsx 商品列表。");
  });
  const [workbook, relationshipXml] = await Promise.all([
    readXml(zip, "xl/workbook.xml"),
    readXml(zip, "xl/_rels/workbook.xml.rels"),
  ]);
  const sheet = workbook.querySelector('sheets > sheet[name="Products"]');
  if (!sheet) throw new Error("找不到 Products 工作表，請選擇商品列表。");
  const relationships = Array.from(relationshipXml.querySelectorAll("Relationship"));
  const sheetId = sheet.getAttributeNS(relationshipNamespace, "id");
  const sheetLink = relationships.find((link) => link.getAttribute("Id") === sheetId);
  if (!sheetLink) throw new Error("無法讀取 Products 工作表。");
  const stringsLink = relationships.find((link) => link.getAttribute("Type").endsWith("/sharedStrings"));
  const [products, sharedStrings] = await Promise.all([
    readXml(zip, relationshipPath(sheetLink)),
    stringsLink && readXml(zip, relationshipPath(stringsLink)),
  ]);
  const strings = sharedStrings ? Array.from(sharedStrings.querySelectorAll("sst > si"), excelText) : [];

  function cellText(cell) {
    if (!cell || cell.getAttribute("t") === "e") return "";
    const value = cell.querySelector("v")?.textContent ?? "";
    switch (cell.getAttribute("t")) {
      case "s": return strings[value] ?? "";
      case "inlineStr": return excelText(cell.querySelector("is"));
      default: return value;
    }
  }

  // ponytail: parse one worksheet in memory. For exports that stall the UI, move parsing to a worker.
  const [header, ...rows] = products.querySelectorAll("sheetData > row");
  if (header?.getAttribute("r") !== "1" || Object.entries(sourceHeaders).some(([col, prefix]) =>
    !cellText(header.querySelector(`c[r="${col}1"]`)).trim().startsWith(prefix)
  )) {
    throw new Error("欄位與商品列表格式不符，請確認 AH、X、Z 與 AK 欄位。");
  }

  const result = [];
  for (const row of rows) {
    const number = row.getAttribute("r");
    const value = (col) => cellText(row.querySelector(`c[r="${col}${number}"]`));
    const stock = Number(value("AK"));
    if (Number.isFinite(stock) && stock < 0) result.push([value("AH"), value("X"), value("Z"), stock]);
  }
  return result.sort((a, b) => order.compare(a[1], b[1]) || order.compare(a[2], b[2]) || a[3] - b[3]);
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

let selection = 0;
fileInput.addEventListener("change", async () => {
  const current = ++selection;
  if (download.hasAttribute("href")) URL.revokeObjectURL(download.href);
  download.removeAttribute("href");
  download.hidden = true;
  status.removeAttribute("data-error");
  const file = fileInput.files[0];
  if (!file) {
    status.textContent = "選擇商品列表後即可下載結果。";
    return;
  }

  status.textContent = "正在讀取商品列表…";
  try {
    if (!/\.xlsx$/i.test(file.name)) throw new Error("請選擇 .xlsx 商品列表。");
    const rows = await readProducts(file);
    if (current !== selection) return;
    const result = await createWorkbook(rows);
    if (current !== selection) return;
    status.textContent = rows.length ? `找到 ${rows.length.toLocaleString("zh-TW")} 筆負庫存。` : "沒有負庫存，可下載只有欄名的 Excel。";
    download.href = URL.createObjectURL(result);
    download.download = `${file.name.slice(0, -5)}_負庫存.xlsx`;
    download.hidden = false;
  } catch (error) {
    if (current !== selection) return;
    status.setAttribute("data-error", "");
    status.textContent = error.message || "轉檔失敗，請重新選擇商品列表。";
  }
});
