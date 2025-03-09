/* global Excel console */

export function getColumnNumber(column: string): number {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result *= 26;
    result += column[i].charCodeAt(0) - "A".charCodeAt(0) + 1;
  }
  return result;
}

// Convert column number to letters (e.g., 1 -> 'A', 26 -> 'Z', 27 -> 'AA')
export function getColumnLetter(columnNumber: number): string {
  let dividend = columnNumber;
  let columnName = "";
  let modulo: number;

  while (dividend > 0) {
    modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - 1) / 26);
  }

  return columnName;
}
export function getCellAddress(baseColumn: string, baseRow: number, rowOffset: number, colOffset: number): string {
  const columnNumber = getColumnNumber(baseColumn);
  const newColumnNumber = columnNumber + colOffset;
  const newColumn = getColumnLetter(newColumnNumber);
  const newRow = baseRow + rowOffset;
  return `${newColumn}${newRow}`;
}
export async function insertText(text: string) {
  // Write text to the top left cell.
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1");
      range.values = [[text]];
      range.format.autofitColumns();
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export function mapRowsToCells(adressRange: string, vals: Array<Array<any>>, worksheetName: string) {
    const startCell = adressRange.split(":")[0].split("!")[1]; // Get starting address
    const baseColumn = startCell.replace(/[0-9]/g, "");
    const baseRow = parseInt(startCell.replace(/[^0-9]/g, ""));

    const worksheetData = {
        name: worksheetName,
        cells: {},
    };
    for (let row = 0; row < vals.length; row++) {
       for (let col = 0; col < vals[row].length; col++) {
           const data = vals[row][col]
           const cellAddress = getCellAddress(baseColumn, baseRow, row, col)
           worksheetData.cells[cellAddress] = data  

       } 
    }
    return worksheetData
}
