# Count blank rows on sheet(s)

This project includes two scripts:  

* [Count blank rows on a given sheet](CountBlankRowsForOneSheet.ts): Traverses through all of the used-range on a given worksheet and returns a blank row count (data only). 
* [Count blank row on all sheets](CountBlankRowsForAllSheets.ts): Traverses through all of the used-range on _all of the worksheets_ worksheet and returns a blank row count (data only). 

Note: In this case, blank row is any row where there's no data (can have format). 

_This sheet returns blank rows of 4_

![Data with blank row](blankrows.png)

_This sheet returns blank rows of 0 (all rows have some data)_

![Data with blank row](noblankrows.png)




