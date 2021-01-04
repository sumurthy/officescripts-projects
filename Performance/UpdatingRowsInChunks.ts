function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const sampleData = ['2020', 'Bread', 'Donut', 500, 0.2];
  let data: (string | number | boolean)[][] = [];
  const sampleRows = 10000;
  for (let i=0; i < sampleRows; i++) {
    data.push([i, ...sampleData]);
  }

  updateRangeInChunks(sheet.getRange("B1"), data)
}

function updateRangeInChunks(startCell: ExcelScript.Range, values: (string | boolean | number)[][], cellsInChunk: number = 1000): boolean {
  console.log("Cells per chunk setting: " + cellsInChunk);
  if (!values.length) {
    console.log(`Invalid input values to update.`);
    return false;
  }
  if (values.length === 0) {
    console.log(`Empty data -- nothing to update.`);
    return true;
  }
  const totalCells = values.length * values[0].length;

  console.log(`Total cells to update in the target range: ${totalCells}`);
  if (totalCells <= cellsInChunk) {
    console.log("No need to chunk -- updating directly")
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
      console.log(`Calling update next chunk. Chunk#: ${chunkCount}`);
      updateNextChunk(startCell, values, rowsPerChunk, totalRowsUpdated);
      rowCount = 0;
      totalRowsUpdated += rowsPerChunk;
    } 
  }
  console.log(rowCount)
  if (rowCount > 0) {
    updateNextChunk(startCell, values, rowCount, totalRowsUpdated);
  }

  return true;
}

/**
 * A Helper function that computes the target range and updates. 
 */

function updateNextChunk(startingCell: ExcelScript.Range, data: (string | boolean | number)[][], rowsPerChunk: number, totalRowsUpdated: number): boolean {
  
  const newStartCell = startingCell.getOffsetRange(totalRowsUpdated, 0);
  const targetRange = newStartCell.getResizedRange(rowsPerChunk - 1, data[0].length - 1);
  console.log(`Updating next chunk at range ${targetRange.getAddress()}`);
  const dataToUpdate = data.slice(totalRowsUpdated, totalRowsUpdated + rowsPerChunk);
  try { 
    targetRange.setValues(dataToUpdate);
  } catch (e) {
    console.log(`Error while updating the chunk range: ${JSON.stringify(e)}`)
    return false;
  }
  return true;
}


/**
 * A Helper function that computes the target range given the target range's starting cell and selected range and updates the values. 
 */
function updateTargetRange(targetCell: ExcelScript.Range, values: (string | boolean | number)[][]): boolean {
  const targetRange = targetCell.getResizedRange(values.length - 1, values[0].length - 1);
  console.log(`Updating the range. ${targetRange.getAddress()}`);
  try { 
    targetRange.setValues(values);
  } catch (e) {
    console.log(`Error while updating the whole range: ${JSON.stringify(e)}`)
    return false;
  }  
  return true;
}
