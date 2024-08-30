function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet
    let sheet = workbook.getActiveWorksheet(); 
  
    // Get the used range of the sheet
    let usedRange = sheet.getUsedRange();
  
    // Get the total number of rows in the used range
    let totalRows = usedRange.getRowCount();

    // Delete the last 3 rows
      sheet.getRangeByIndexes(totalRows - 3, 0, 3, usedRange.getColumnCount()).delete(ExcelScript.DeleteShiftDirection.up);
}
  