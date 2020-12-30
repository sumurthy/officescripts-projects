# Range Basics: Get Methods

Range is the foundational object within Office Script's Excel object model. [Range APIs](https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts) allows access to both data and format that available on the grid and links other key objects within Excel such as worksheets, tables, charts, etc. 

Range is identified using its address such as "A1:B4" or using a named-item, which is a named key for a given set of cells. In Excel object model, both a cell and group of cells are referred as _range_. Range can contain cell level attributes such as data within a cell and also cell and cells-level attributes such as format, borders, etc. 
Range can also be obtained via. users selection, which consists of at least one cell. As you interact with range, it is important to keep these cell and range relationship clear. 

The following are core set of getter, setter and other useful methods most often used in scripts. This is a great starting point to begin the API jorney. The following sections divides the methods into easy to understand grouping and helps to build an easy mental model as you begin to unlock the range object's APIs. 

## Video link 

[![Watch step by step video](v_rangebasics.png)](https://youtu.be/4emjkOFdLBA "Watch step by step video")

## Scripts 

* [Basic read and write](ReadWriteBasics.ts)
* [Clear column filter](../Clear%20Table%20Filter%20For%20Active%20Cell/ClearFilterFromCell.ts)
* [Color each cell with unique color](ColorCells.ts): shows how to iterate over cells and perform operations 
* [Update range with values using 2D array](Update2DRangeValues.ts): dynamically calculates the range dimension to update based on 2D array values. 


## Office Script link 
A collection of snippets showin in the video (doesn't do anything meaningful..)

* [RangeBasics](RangeBasics.ts)


## Methods that return some range metadata 

* getAddress(), getAddressLocal()
* getCellCount()
* getRowCount(), getColumnCount()

## Methods that returns data/constants associated with a given range 

### Returned as singular cell value

* getFormula(), getFormulaLocal()
* getFormulaR1C1()
* getNumberFormat(), getNumberFormatLocal()
* getText()
* getValue()
* getValueType()

### Returned as 2-D arrays (whole range)

* getFormulas(), getFormulasLocal()
* getFormulasR1C1()
* getNumberFormatCategories()
* getNumberFormats(), getNumberFormatsLocal()
* getTexts()
* getValues()
* getValueTypes()
* getHidden()
* getIsEntireRow()
* getIsEntireColumn()

## Methods that return other range object

* getCell(row, column)
* getColumn(column)
* getColumnHidden()
* getColumnsAfter(count)
* getColumnsBefore(count)
* getEntireColumn()
* getEntireRow()
* getLastCell()
* getLastColumn()
* getLastRow()
* getRow(row)
* getRowHidden()
* getRowsAbove(count)
* getRowsBelow(count)

--- Important/interesting---

* _workbook_.getSelectedRange()
* _workbook_.getActiveCell()
* getUsedRange(valuesOnly)
* getAbsoluteResizedRange(numRows, numColumns)
* getOffsetRange(rowOffset, columnOffset)
* getResizedRange(deltaRows, deltaColumns)

## Methods that return a range object in relation to another range object 
* getBoundingRect(anotherRange)
* getIntersection(anotherRange)


## Methods that return other objects (non range objects) 

* getDirectPrecedents()
* getWorksheet()
* getTables(fullyContained)
* getPivotTables(fullyContained)
* getDataValidation()
* getPredefinedCellStyle()

## Set methods 

### Singular cell set methods

* setFormula(formula)	
* setFormulaLocal(formulaLocal)	
* setFormulaR1C1(formulaR1C1)	
* setNumberFormatLocal(numberFormatLocal)	
* setValue(value)	

### 2D / entire range set methods 

* setFormulas(formulas)	
* setFormulasLocal(formulasLocal)	
* setFormulasR1C1(formulasR1C1)	
* setNumberFormat(numberFormat)	
* setNumberFormats(numberFormats)	
* setNumberFormatsLocal(numberFormatsLocal)
* setValues(values)

## Other methods 

* merge(across)
* unmerge()

## Coming soon 

* Range edge APIs

