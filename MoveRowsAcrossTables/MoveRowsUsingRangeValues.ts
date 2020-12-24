function main(workbook: ExcelScript.Workbook) {

  // Update the table names, column index to look-up on as needed
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';
  const IndexOfColumnToFilterOn = 1; // 0-index
  const ValueToFilterOn = 'Clothing';

  // Get the table objects
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TargetTableName}) and target table (${SourceTableName}) are present before running the script. `);
    return;
  }
  // Range object of table data
  const sourceRange = sourceTable.getRangeBetweenHeaderAndTotal();

  const dataRows: (number | string | boolean)[][] = sourceTable.getRangeBetweenHeaderAndTotal().getValues();

  let rowsToRemoveValues: (number | string | boolean)[][] = [];
  let rowAddressToRemove: string[] = [];

  // Get the data values to insert to target table 
  for (let i = 0; i < dataRows.length; i++) { 
    if (dataRows[i][IndexOfColumnToFilterOn] === ValueToFilterOn) {
      rowsToRemoveValues.push(dataRows[i]);
      // Get the intersection between table address and the entire row where we found the match. This provides the address of the range to remove. 
      let address = sourceRange.getIntersection(sourceRange.getCell(i,0).getEntireRow()).getAddress();
      rowAddressToRemove.push(address);
    }
  }
  // If no data rows to process, exit script.
  if (rowsToRemoveValues.length < 1) {
    console.log('No data rows selected from the source table that matched the filter criteria.');
    return;
  }
  console.log(`Adding ${rowsToRemoveValues.length} rows to target table.`);
  // Insert rows at the end of target table. Change the first argument to suit your target location (e.g., 0 for beginning, -1 for end).
  targetTable.addRows(-1, rowsToRemoveValues)
  // Get worksheet reference where the table rows to be deleted resides. 
  const sheet = sourceTable.getWorksheet();

  // Save all of the filter criteria 
  // Initialize an empty object to hold the filter criteria
  const tableFilters = {};
  // For each table column, collect the filter criteria
  sourceTable.getColumns().forEach((column) => {
    let colFilterCriteria = column.getFilter().getCriteria();
    if (colFilterCriteria) {
      // If we don't remove these two keys, the API fails for some reason. So, remove these..
      delete colFilterCriteria['@odata.type'];
      delete colFilterCriteria['subField'];
      tableFilters[column.getName()] = colFilterCriteria;
    }
  });
  // Remove all filters before removing rows
  sourceTable.getAutoFilter().clearCriteria();

  // !!Important!!  Reverse the address and remove from the bottom so that the right rows are removed. If not reversed, the resulting row upwards shift will mean that incorrect rows will be removed. 
  console.log(`Removing ${rowAddressToRemove.length} from the source table. `)
  rowAddressToRemove.reverse().forEach((address) => {
    console.log(`Deleting row: ${address}`)
    sheet.getRange(address).delete(ExcelScript.DeleteShiftDirection.up);
  });
  // Re-apply filters 
  // Log the criteria for testing purpose (not required)
  console.log(tableFilters);

  // Re-apply all column filters
  Object.keys(tableFilters).forEach((columnName) => {
    sourceTable.getColumnByName(columnName).getFilter().apply(tableFilters[columnName]);
  })  
  console.log("Finished.")
  return;
}
