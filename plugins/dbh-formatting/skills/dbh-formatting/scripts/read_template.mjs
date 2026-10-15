import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = process.argv[2];
const outputPath = process.argv[3];
if (!inputPath || !outputPath) throw new Error("Usage: node read_template.mjs <input.xlsx> <output.json>");

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheets = [];
for (const sheet of workbook.worksheets.items) {
  const used = sheet.getUsedRange();
  const address = used?.address ?? "A1:Z50";
  const region = await workbook.inspect({
    kind: "region",
    sheetId: sheet.name,
    range: address,
    maxChars: 24000,
    tableMaxRows: 100,
    tableMaxCols: 40,
    tableMaxCellChars: 200,
  });
  const formulas = await workbook.inspect({
    kind: "formula",
    sheetId: sheet.name,
    range: address,
    maxChars: 12000,
    options: { maxResults: 300 },
  });
  sheets.push({ name: sheet.name, usedRange: address, region: region.ndjson, formulas: formulas.ndjson });
}
await fs.writeFile(outputPath, JSON.stringify({ schema: "dbh-formatting/excel-template@0.1", sheets }, null, 2), "utf8");
console.log(JSON.stringify({ sheets: sheets.map((s) => ({ name: s.name, usedRange: s.usedRange })) }, null, 2));

