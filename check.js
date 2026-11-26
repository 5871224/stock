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
  await input.setInputFiles({ name: "invalid.xlsx", mimeType: "application/octet-stream", buffer: Buffer.from("invalid") });
  await page.locator("#status[data-error]").waitFor({ state: "visible" });
  await link.waitFor({ state: "hidden" });
  return "PASS";
}
