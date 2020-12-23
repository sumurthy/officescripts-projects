# Move rows across tables by saving filters, processing and re-applying the filters

This script achieves the following: 

* Remove all filters on the source table, wraps text on all rows.
* Select rows from the source table where the value in a column is equal to _some value_. 
* Move all selected rows into a another (target) table on another worksheet
* Re-applies the filter on the source table.

