function main(workbook: ExcelScript.Workbook) {
    // Get active cell
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell 
    const tables = cell.getTables();
    
    // If there is no table on the selection, return
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }
    // Get table (since it is a )
    const currentTable = tables[0];

    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    const entireCol = cell.getEntireColumn();
    const intersect = entireCol.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get column
    const col = currentTable.getColumnByName(headerCellValue);
    // Clear filter
    col.getFilter().clear();
    
}
