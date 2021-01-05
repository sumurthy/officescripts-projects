function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getActiveWorksheet();

    // Sample data that'll be repeated in range
    const sampleData = ['2020', 'Bread', 'Donut', 500, 0.2];
    let data: (string | number | boolean)[][] = [];
    // Number of rows in the random data (x 6 columns)
    const sampleRows = 10000;
    // Dynamically generate some random data for testing purpose. 
    for (let i = 0; i < sampleRows; i++) {
      data.push([i, ...sampleData]);
    }
    
    const updated = updateRangeInChunks(sheet.getRange("B1"), data);
    if (!updated) {
        console.log(`Update did not take place or complete. Chech and run again.`)
    }
    return;
}
  
function updateRangeInChunks(
    startCell: ExcelScript.Range, 
    values: (string | boolean | number)[][], 
    cellsInChunk: number = 10000
    ): boolean {

    console.log(`Cells per chunk setting: ${cellsInChunk}`);
    if (!values) {
      console.log(`Invalid input values to update.`);
      return false;
    }
    if (values.length === 0 || values[0].length === 0) {
        console.log(`Empty data -- nothing to update.`);
        return true;
    }
    const totalCells = values.length * values[0].length;
  
    console.log(`Total cells to update in the target range: ${totalCells}`);
    if (totalCells <= cellsInChunk) {
      console.log(`No need to chunk -- updating directly`);
      updateTargetRange(startCell, values);
      return true;
    }
  
    const rowsPerChunk = Math.floor(cellsInChunk / values[0].length);
    console.log("Rows per chunk " + rowsPerChunk);
    let rowCount = 0;
    let totalRowsUpdated = 0;
    let chunkCount = 0;
  
    for (let i = 0; i < values.length; i++) {
      rowCount++;
      if (rowCount === rowsPerChunk) {
        chunkCount++;
        console.log(`Calling update next chunk function. Chunk#: ${chunkCount}`);
        updateNextChunk(startCell, values, rowsPerChunk, totalRowsUpdated);
        rowCount = 0;
        totalRowsUpdated += rowsPerChunk;
        console.log(`${(totalRowsUpdated/values.length)*100}% Done`);
  
      }
    }
    console.log(`Updating remaining rows -- last chunk: ${rowCount}`)
    if (rowCount > 0) {
      updateNextChunk(startCell, values, rowCount, totalRowsUpdated);
    }
    console.log(`Done with all updates.`);
    return true;
}
  
  /**
   * A Helper function that computes the target range and updates. 
   */
  
function updateNextChunk(
      startingCell: ExcelScript.Range, 
      data: (string | boolean | number)[][], 
      rowsPerChunk: number, 
      totalRowsUpdated: number
    ) {
  
    const newStartCell = startingCell.getOffsetRange(totalRowsUpdated, 0);
    const targetRange = newStartCell.getResizedRange(rowsPerChunk - 1, data[0].length - 1);
    console.log(`Updating chunk at range ${targetRange.getAddress()}`);
    const dataToUpdate = data.slice(totalRowsUpdated, totalRowsUpdated + rowsPerChunk);
    try {
      targetRange.setValues(dataToUpdate);
    } catch (e) {
      throw `Error while updating the chunk range: ${JSON.stringify(e)}`;
    }
    return;
}
  
  /**
   * A Helper function that computes the target range given the target range's starting cell and selected range and updates the values. 
   */
function updateTargetRange(
      targetCell: ExcelScript.Range, 
      values: (string | boolean | number)[][]
    ) {
    const targetRange = targetCell.getResizedRange(values.length - 1, values[0].length - 1);
    console.log(`Updating the range. ${targetRange.getAddress()}`);
    try {
      targetRange.setValues(values);
    } catch (e) {
      throw `Error while updating the whole range: ${JSON.stringify(e)}`;
    }
    return;
}
