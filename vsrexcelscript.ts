function main(workbook: ExcelScript.Workbook) {

  // DELETING THE LAST 3 ROWS
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

  // DELETING UNNECESSARY COLUMNS
  sheet.getRange("B:B").delete(ExcelScript.DeleteShiftDirection.left);
  sheet.getRange("E:H").delete(ExcelScript.DeleteShiftDirection.left);
  sheet.getRange("F:G").delete(ExcelScript.DeleteShiftDirection.left);
  sheet.getRange("J:M").delete(ExcelScript.DeleteShiftDirection.left);
  sheet.getRange("K:K").delete(ExcelScript.DeleteShiftDirection.left);
  sheet.getRange("Q:U").delete(ExcelScript.DeleteShiftDirection.left);
  sheet.getRange("V:W").delete(ExcelScript.DeleteShiftDirection.left);

  // DELETING DUPLICATES IN COLUMN A
  // Find the last row in column A
  let lastRow = sheet.getRange("A:A").find("*", { searchDirection: ExcelScript.SearchDirection.backwards }).getRowIndex();

  // Remove duplicates in Column A
  let columnAValues = sheet.getRange(`A2:A${lastRow + 1}`).getValues();
  let uniqueValues = new Set();

  for (let i = columnAValues.length - 1; i >= 0; i--) {
    if (uniqueValues.has(columnAValues[i][0])) {
      sheet.getRange(`A${i + 2}`).getEntireRow().delete(ExcelScript.DeleteShiftDirection.up);
    } else {
      uniqueValues.add(columnAValues[i][0]);
    }
  }

  // Update the last row after deleting duplicates
  lastRow = sheet.getRange("A:A").find("*", { searchDirection: ExcelScript.SearchDirection.backwards }).getRowIndex();


  //ADD PSKU COLUMNS AND DUPLICATE VALUES
  //Insert range B:F on sheet, move existing cells to the right
  sheet.getRange("B:F").insert(ExcelScript.InsertShiftDirection.right);

  //Paste Column A to Range B:F
  sheet.getRange("B:F").copyFrom(sheet.getRange("A:A"), ExcelScript.RangeCopyType.all, false, false);

  //Rename the Headers
  sheet.getRange("B1").setValue("Inventory CD");
  sheet.getRange("C1").setValue("SKU");
  sheet.getRange("D1").setValue("PSKU");
  sheet.getRange("E1").setValue("Parent SKU");
  sheet.getRange("F1").setValue("Parent SKUID");

  //UPDATE AGE VALUES
  let ageColumn = sheet.getRange(`I1:I${lastRow}`);
  let ageValues = ageColumn.getValues();

  for (let i = 0; i < ageValues.length; i++) {
    if (ageValues[i][0] === "Adults" || ageValues[i][0] === "All Ages" || ageValues[i][0] === "Youth + Adults") {
      ageValues[i][0] = "Adult"
    } else if (ageValues[i][0] === "Pre-School") {
      ageValues[i][0] = "Toddler";
    }
  }

  ageColumn.setValues(ageValues);
}