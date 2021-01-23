function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('Sheet3');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}

function addRow(sheet: ExcelScript.Worksheet, data: (string | number | boolean)[]): void {

    const usedRange = sheet.getUsedRange();
    console.log(usedRange.getAddress());
    const startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);

    targetRange.setValues([data]);
    return;
}
