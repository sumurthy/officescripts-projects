# Move rows across tables by saving filters, processing and re-applying the filters

This script achieves the following: 

* Remove all filters on the source table, wraps text on all rows.
* Select rows from the source table where the value in a column is equal to _some value_. 
* Move all selected rows into a another (target) table on another worksheet
* Re-applies the filter on the source table.

There are two variants of the solution. It mainly differs in the way the rows are selected. 
* In one script, the rows are selected by applying the table filter and reading the visible range. 
* In the second, the rows are selected simply by reading the values and extracting the row values.
