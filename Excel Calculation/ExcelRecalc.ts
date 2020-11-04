function main(workbook: ExcelScript.Workbook) {
    // Set calculation mode
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get calculation mode 
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Calculate (for manual mode files)
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
