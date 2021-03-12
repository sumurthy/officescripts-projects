function main(workbook: ExcelScript.Workbook)
{
  const table = workbook.getActiveWorksheet().getTables()[0];
  const header = table.getHeaderRowRange().getTexts();
  // Display header as 1D array
  console.log(header[0]); // Extract 1st row


  const sales = table.getColumnByName('Sales').getRange().getValues();
  console.log(sales);
  
  // Get column-0
  const salesAs1DArray = extractColumn(sales, 0);
  console.log(salesAs1DArray);

  // Add 100 to each value
  const revisedSales = salesAs1DArray.map(data => data as number + 100);
  // Add a column
  table.addColumn(-1, revisedSales);

  const salesBackTo2D = convertColumnTo2D(salesAs1DArray);
  console.log(salesBackTo2D);

}

/**
 * Extract a column from 2D array and return result.
 */
function extractColumn(data: (string | number | boolean)[][], index: number): (string | number | boolean)[] {

  const column = data.map((row) => {
    return row[index];
  })
  return column;
}

/**
 * Convert a flat array into 2D array that can be used as range column
 */
function convertColumnTo2D(data: (string | number | boolean)[]): (string | number | boolean)[][] {

  const columnAs2D = data.map((row) => {
    return [row];
  })
  return columnAs2D;
}
