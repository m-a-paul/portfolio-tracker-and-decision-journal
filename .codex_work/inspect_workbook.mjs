import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const input = await FileBlob.load('/Users/paulmiyake/Downloads/Stocks and Options.xlsx');
const workbook = await SpreadsheetFile.importXlsx(input);
console.log('workbook keys', Object.keys(workbook));
console.log('worksheets keys', Object.keys(workbook.worksheets ?? {}));
console.log('sheet names', workbook.worksheets?.items?.map(s => s.name) ?? workbook.worksheets?.map?.(s=>s.name));
for (const sheet of workbook.worksheets.items ?? []) {
  console.log('sheet', sheet.name, Object.keys(sheet));
  try {
    const out = await workbook.inspect({kind:'table', range:`${sheet.name}!A1:Z40`, include:'values,formulas', tableMaxRows:40, tableMaxCols:26});
    console.log(out.ndjson.slice(0,4000));
  } catch(e) { console.log('inspect error', e.message); }
}
