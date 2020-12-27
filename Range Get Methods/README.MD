# Range Basics: Get Methods

## Methods that return some range metadata 

* getAddress(), getAddressLocal()
* getCellCount()
* getRowCount(), getColumnCount()

## Methods that returns data/constants associated with a giben range 

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

# Other methods 

* merge(across)
* unmerge()

