function main(workbook: ExcelScript.Workbook, keys: string): string {

    // Needed for testing purpose. Override the keys
    // keys = `[{"event":"E123","date":43892,"location":"Montgomery","capacity":10},{"event":"E124","date":43892,"location":"Juneau","capacity":15},{"event":"E125","date":43897,"location":"Phoenix","capacity":15},{"event":"E126","date":43914,"location":"Boise","capacity":25},{"event":"E127","date":43918,"location":"Salt Lake City","capacity":20},{"event":"E128","date":43938,"location":"Fremont","capacity":3},{"event":"E129","date":43938,"location":"Vancouver","capacity":50}]`;
  
    let table = workbook.getWorksheet('Transactions').getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    range.clear(ExcelScript.ClearApplyTo.formats);
  
    let overallMatch = true;
  
    table.getColumnByName('Date').getRangeBetweenHeaderAndTotal().setNumberFormatLocal("yyyy-mm-dd;@");
    table.getColumnByName('Capacity').getRangeBetweenHeaderAndTotal().getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    let rows = range.getValues();
    let keysObject = JSON.parse(keys) as EventData[];
    for (let i=0; i < rows.length; i++){
      let row = rows[i];
      let [event, date, location, capacity] = row;
      let match = false;
      // console.log("Currently on event ID " + event + " row: " + i );
      for (let keyObject of keysObject){
        // console.log("Comparing: " + event + " with master event record: " + keyObject.event);
        if (keyObject.event === event) {
          match = true;
          if (keyObject.date !== date) {
            overallMatch = false;
            range.getCell(i, 1).getFormat()
              .getFill()
              .setColor("FFFF00");
          }
          if (keyObject.location !== location) {
            overallMatch = false;
            range.getCell(i, 2).getFormat()
              .getFill()
              .setColor("FFFF00");
          }
          if (keyObject.capacity !== capacity) {
            overallMatch = false;
            range.getCell(i, 3).getFormat()
              .getFill()
              .setColor("FFFF00");
          }   
          break;             
        }
      }
      if (!match) {
        overallMatch = false;
        range.getCell(i, 0).getFormat()
          .getFill()
          .setColor("FFFF00");      
      }
  
    }
    let returnString = "All the data is in the right order.";
    if (overallMatch === false) {
      returnString = "Mismatch found. Data requires your review.";
    }
    console.log("Returning: " + returnString);
    return returnString;
  }
  
  interface EventData {
    event: string
    date: number
    location: string
    capacity: number
  }