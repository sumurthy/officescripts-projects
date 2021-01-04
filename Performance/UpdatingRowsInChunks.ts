function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();

  const data = [
    ['2020', 'Bread', 'Donut', 500, 0.2],
    ['2021', 'Bread', 'Donut', 500, 0.2],
    ['2022', 'Bread', 'Donut', 500, 0.2],
    ['2023', 'Bread', 'Donut', 500, 0.2],
    ['2024', 'Bread', 'Donut', 500, 0.2],
    ['2025', 'Bread', 'Donut', 500, 0.2],
    ['2026', 'Bread', 'Donut', 500, 0.2],
    ['2027', 'Bread', 'Donut', 500, 0.2],
    ['2028', 'Bread', 'Donut', 500, 0.2],
    ['2029', 'Bread', 'Donut', 500, 0.2],
    ['2030', 'Bread', 'Donut', 500, 0.2],
    ['2031', 'Bread', 'Donut', 500, 0.2],
    ['2032', 'Bread', 'Donut', 500, 0.2],
    ['2033', 'Bread', 'Donut', 500, 0.2],    
    ['2034', 'Bread', 'Donut', 500, 0.2],
    ['2035', 'Bread', 'Donut', 500, 0.2],
    ['2036', 'Bread', 'Donut', 500, 0.2],
    ['2037', 'Bread', 'Donut', 500, 0.2],
    ['2038', 'Bread', 'Donut', 500, 0.2],
    ['2039', 'Bread', 'Donut', 500, 0.2],
    ['2040', 'Bread', 'Donut', 500, 0.2],
    ['2041', 'Bread', 'Donut', 500, 0.2],
    ['2042', 'Bread', 'Donut', 500, 0.2],
    ['2043', 'Bread', 'Donut', 500, 0.2],
    ['2044', 'Bread', 'Donut', 500, 0.2],    
    ['2045', 'Bread', 'Donut', 500, 0.2],  
    ['2046', 'Bread', 'Donut', 500, 0.2],      
    ['2047', 'Bread', 'Donut', 500, 0.2],      
    ['2048', 'Bread', 'Donut', 500, 0.2],      
    ['2049', 'Bread', 'Donut', 500, 0.2],      
    ]

  updateRangeInChunks(sheet.getRange("B1"), data, 23)
}

function updateRangeInChunks(startCell: ExcelScript.Range, values: (string | boolean | number)[][], cellsInChunk: number = 10): boolean {
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
