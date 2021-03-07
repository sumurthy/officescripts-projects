function main(workbook: ExcelScript.Workbook) {

  const xlDate1value = workbook.getActiveWorksheet().getRange('A2').getValue();

  const jsDate1 = xlDateToJSDate(xlDate1value as number);
  console.log(jsDate1.toUTCString())
  
  // Search for Breaks in the current worksheet 
  const result = workbook.getActiveWorksheet().findAll("brakes", {matchCase: false});
  if (result) {
    console.log(result.getAddress());

    // Split the resulting adddress string into array 
    const cellAddressArray = result.getAddress().split(',');
    console.log(cellAddressArray);
  }
}


/**
 * Function to return the JS date from Excel date
 * 
 */
function xlDateToJSDate(serialDate: number): Date {
  var days = Math.floor(serialDate);
  var hours = Math.floor((serialDate % 1) * 24);
  var minutes = Math.floor((((serialDate % 1) * 24) - hours) * 60)
  const returnDate = new Date(Date.UTC(0, 0, serialDate, hours - 17, minutes));
  return returnDate;
}
