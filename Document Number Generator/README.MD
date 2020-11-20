# Generate a new unique identifier and add a row to table and range

This scenario helps a user generate a unique document number with a specific format and add an entry to a range or table (two sepearate scenarios). 
The new entry or a row added will contain the newly generated unique document number and few other attributes passed to the script. 

The sample contains two part. 

* Part-1: achieves this scenario by reading and adding row to a worksheet containing plain range. 
* PArt-2: achieves this scenario by reading and adding row to a table (little bit simpler). 
## Part-1 Generate document number on a range

## Video 

[Step by step recording]()

### Before the new row is added 
![Range before](Range-Before.png) 
### After the new row is added 
![Range after](Range-After.png) 

## Part-2 Generate document number on a table
### Before the new row is added 
![Table before](Table-Before.png) 
### After the new row is added 
![Table after](Table-After.png) 

## Use this in Power Automate 

## Office Scripts source code

1. [Generate and add row to range](DocNumGenForRange.ts)
1. [Generate and add row to table](DocNumGenForTable.ts)

## Excel files used

[Excel file](Table-Data-With-Hyperlinks.xlsx)
