// With the site open, run: playwright-cli run-code --filename=check.js
async (page) => {
  const fixture = await page.evaluate(async () => {
    const createUrl = URL.createObjectURL.bind(URL);
    URL.createObjectURL = (blob) => {
      window.testWorkbook = blob;
      return createUrl(blob);
    };
    const zip = new JSZip();
    const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    zip.file("xl/workbook.xml", `<workbook xmlns:r="${rel}"><sheets><sheet name="Products" r:id="rId1"/></sheets></workbook>`);
    zip.file("xl/_rels/workbook.xml.rels", `<Relationships><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
    const row = (number, values) => `<row r="${number}">${["AH", "X", "Z", "AK"].map((col, index) => {
      const ref = `${col}${number}`;
      const value = values[index];
      return typeof value === "number" ? `<c r="${ref}"><v>${value}</v></c>` :
        `<c r="${ref}" t="inlineStr"><is><t>${value}</t></is></c>`;
    }).join("")}</row>`;
    zip.file("xl/worksheets/sheet1.xml", `<worksheet><sheetData>${[
      ["[29]商品款式", "[19]商品款式", "[21]商品款式", "[32]商品款式"],
      ["005", "黑", "F", "-31"],
      ["004", "黑", "2XL", -1],
      ["003", "白", "F", 0],
      ["002", "白", "F", ""],
      ["001", "白", "F", 1],
      ["000", "白", "F", " -2 "],
    ].map((values, index) => row(index + 1, values)).join("")}</sheetData></worksheet>`);
    return zip.generateAsync({ type: "base64" });
  });

  const input = page.locator("#product-file");
  await input.setInputFiles({ name: "products.xlsx", mimeType: "application/octet-stream", buffer: Buffer.from(fixture, "base64") });
  const link = page.getByRole("link", { name: "下載 Excel" });
  await link.waitFor({ state: "visible" });
  const data = await page.evaluate(async () => {
    const zip = await JSZip.loadAsync(await window.testWorkbook.arrayBuffer());
    const sheet = await zip.file("xl/worksheets/sheet1.xml").async("text");
    const xml = new DOMParser().parseFromString(sheet, "application/xml");
    return Array.from(xml.querySelectorAll('row:not([r="1"])'), (row) =>
      Array.from(row.querySelectorAll("c"), (cell) => [cell.getAttribute("t"), cell.textContent])
    );
  });
  const expected = [
    [["inlineStr", "000"], ["inlineStr", "白"], ["inlineStr", "F"], [null, "-2"]],
    [["inlineStr", "004"], ["inlineStr", "黑"], ["inlineStr", "2XL"], [null, "-1"]],
    [["inlineStr", "005"], ["inlineStr", "黑"], ["inlineStr", "F"], [null, "-31"]],
  ];
  if (JSON.stringify(data) !== JSON.stringify(expected)) throw new Error("負庫存資料、欄位型別或排序不符。");

  const report = await page.evaluate(async () => {
    const zip = new JSZip();
    const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    zip.file("xl/workbook.xml", `<workbook xmlns:r="${rel}"><sheets><sheet name="訂單明細" r:id="rId1"/></sheets></workbook>`);
    zip.file("xl/_rels/workbook.xml.rels", `<Relationships><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/></Relationships>`);
    zip.file("xl/styles.xml", '<styleSheet><fills count="1"><fill><patternFill patternType="none"/></fill></fills><cellXfs count="2"><xf fillId="0"/><xf fillId="0" numFmtId="2"/></cellXfs></styleSheet>');
    const values = [
      ["商品款式", "SKU", "數量"],
      ["白 / F", "000", 1],
      ["紅 / F", "000", 1],
      ["白 / F", "000", 2],
      ["黑 / 2XL", "004", 1],
      ["白 / F", "000", 1],
    ];
    zip.file("xl/worksheets/sheet1.xml", `<worksheet><sheetData>${values.map((row, index) => {
      const r = index + 1;
      return `<row r="${r}">${index ? "" : '<c r="A1" t="inlineStr"><is><t>時間</t></is></c>'}${row.map((value, col) => {
        const ref = ["M", "N", "P"][col] + r;
        return typeof value === "number" ? `<c r="${ref}" s="1"><v>${value}</v></c>` :
          `<c r="${ref}" t="inlineStr"><is><t>${value}</t></is></c>`;
      }).join("")}</row>`;
    }).join("")}</sheetData></worksheet>`);
    return zip.generateAsync({ type: "base64" });
  });
  await page.locator("#report-file").setInputFiles({ name: "orders.xlsx", mimeType: "application/octet-stream", buffer: Buffer.from(report, "base64") });
  const reportLink = page.getByRole("link", { name: "下載標綠訂單報表" });
  await reportLink.waitFor({ state: "visible" });
  const reportSummary = await page.locator("#report-status").textContent();
  const highlight = await page.evaluate(async () => {
    const zip = await JSZip.loadAsync(await window.testWorkbook.arrayBuffer());
    const sheet = await zip.file("xl/worksheets/sheet1.xml").async("text");
    const styles = await zip.file("xl/styles.xml").async("text");
    const xml = new DOMParser().parseFromString(sheet, "application/xml");
    const rows = Array.from(xml.querySelectorAll('row:has(c[r^="A"][s="2"])'), (row) => row.getAttribute("r"));
    return { rows, keptNumberFormat: styles.includes('fillId="1" numFmtId="2"') };
  });
  if (highlight.rows.join() !== "4,5,6" || !highlight.keptNumberFormat || !reportSummary.includes("多 1 件")) {
    throw new Error("訂單款式配對、表尾累計、綠色格式或原始數字格式不符。");
  }

  await input.setInputFiles({ name: "invalid.xlsx", mimeType: "application/octet-stream", buffer: Buffer.from("invalid") });
  await page.locator("#status[data-error]").waitFor({ state: "visible" });
  await link.waitFor({ state: "hidden" });
  await reportLink.waitFor({ state: "hidden" });
  return "PASS";
}
