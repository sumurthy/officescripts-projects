function main(workbook: ExcelScript.Workbook): BasicObj[] {
    const sheet = workbook.getWorksheet('Sheet4');
    const table = sheet.getTables()[0];
    const dataRows = table.getRange().getValues() as string[][];
    // or
    // let dataRows = sheet.getUsedRange().getValues();
    const selectedRows = dataRows.filter((row, i) => {
      // Select header row and any data row with the status column equal to approach value
      return (row[3] === 'Approach' || i === 0)
    })
    const output: BasicObj[] = returnObjectFromValues(selectedRows);
    console.log(output);
    return output;
  }
  
  /**
   * This helper funciton converts table values into an object array.
   */
  function returnObjectFromValues(values: string[][]): BasicObj[] {
    let objArray = [];
    let objKeys: string[] = [];
    for (let i = 0; i < values.length; i++) {
      if (i === 0) {
        objKeys = values[i]
        continue;
      }
      let obj = {}
      for (let j = 0; j < values[i].length; j++) {
        obj[objKeys[j]] = values[i][j]
      }
      objArray.push(obj);
    }
    console.log(JSON.stringify(objArray));
    return objArray;
  }
  
  interface BasicObj {
    [key: string]: string
  }
  