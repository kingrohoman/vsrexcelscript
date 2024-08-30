function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet
    let sheet = workbook.getActiveWorksheet();
  
    // Get the used range of the sheet
    let usedRange = sheet.getUsedRange();
  
    // Get the total number of rows in the used range
    let totalRows = usedRange.getRowCount();
  
    // Check if there are at least 3 rows to delete
    if (totalRows > 3) {
      // Delete the last 3 rows
      sheet.getRangeByIndexes(totalRows - 3, 0, 3, usedRange.getColumnCount()).delete(ExcelScript.DeleteShiftDirection.up);
    } else {
      // If there are fewer than 3 rows, delete all rows
      sheet.getRangeByIndexes(0, 0, totalRows, usedRange.getColumnCount()).delete(ExcelScript.DeleteShiftDirection.up);
    }
  }
  