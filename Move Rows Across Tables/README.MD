# Move rows across tables by saving filters, processing and re-applying the filters

This script achieves the following: 

* Select rows from the source table where the value in a column is equal to _some value_. 
* Move all selected rows into a another (target) table on another worksheet
* Re-applies the relevant filters on the source table.

![Before and after images](table-filter-s-t.png) 

## Video link

[Demo video]()

## Solution

There are two variants of the solution. It mainly differs in the way the rows are selected. 
* In one script, the rows are selected by applying the table filter and reading the visible range. [Script](MoveRowsAcrossTables/MoveRowsUsingTableFilter.ts)
* In the second, the rows are selected simply by reading the values and extracting the row values. [Script](MoveRowsAcrossTables/MoveRowsUsingTableFilter.ts)

## Input Excel file
Download the input Excel file used in this sample and try it out yourself! 

This is the [input Excel file](Excel_TableFilters.xlsx) being used. 

## Office Scripts

The solution uses two Office Scripts:

1. [Script](MoveRowsAcrossTables/MoveRowsUsingTableFilter.ts)
1. [Script](MoveRowsAcrossTables/MoveRowsUsingTableFilter.ts)


