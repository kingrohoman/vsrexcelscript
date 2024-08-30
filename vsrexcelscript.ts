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
    sheet
      .getRangeByIndexes(totalRows - 3, 0, 3, usedRange.getColumnCount())
      .delete(ExcelScript.DeleteShiftDirection.up);
  } else {
    // If there are fewer than 3 rows, delete all rows
    sheet
      .getRangeByIndexes(0, 0, totalRows, usedRange.getColumnCount())
      .delete(ExcelScript.DeleteShiftDirection.up);
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
  let lastRow = sheet
    .getRange("A:A")
    .find("*", { searchDirection: ExcelScript.SearchDirection.backwards })
    .getRowIndex();

  // Remove duplicates in Column A
  let columnAValues = sheet.getRange(`A2:A${lastRow + 1}`).getValues();
  let uniqueValues = new Set();

  for (let i = columnAValues.length - 1; i >= 0; i--) {
    if (uniqueValues.has(columnAValues[i][0])) {
      sheet
        .getRange(`A${i + 2}`)
        .getEntireRow()
        .delete(ExcelScript.DeleteShiftDirection.up);
    } else {
      uniqueValues.add(columnAValues[i][0]);
    }
  }

  // Update the last row after deleting duplicates
  lastRow = sheet
    .getRange("A:A")
    .find("*", { searchDirection: ExcelScript.SearchDirection.backwards })
    .getRowIndex();

  //ADD PSKU COLUMNS AND DUPLICATE VALUES
  //Insert range B:F on sheet, move existing cells to the right
  sheet.getRange("B:F").insert(ExcelScript.InsertShiftDirection.right);

  //Paste Column A to Range B:F
  sheet
    .getRange("B:F")
    .copyFrom(
      sheet.getRange("A:A"),
      ExcelScript.RangeCopyType.all,
      false,
      false
    );

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
    if (
      ageValues[i][0] === "Adults" ||
      ageValues[i][0] === "All Ages" ||
      ageValues[i][0] === "Youth + Adults"
    ) {
      ageValues[i][0] = "Adult";
    } else if (ageValues[i][0] === "Pre-School") {
      ageValues[i][0] = "Toddler";
    }
  }

  ageColumn.setValues(ageValues);

  // UPDATE GENDER VALUES
  let genderColumn = sheet.getRange(`H2:H${lastRow + 1}`);
  let genderValues = genderColumn.getValues();

  for (let i = 0; i < genderValues.length; i++) {
    if (ageValues[i][0] === "Adult") {
      if (genderValues[i][0] === "Unisex" || genderValues[i][0] === "Male") {
        genderValues[i][0] = "Mens";
      } else {
        genderValues[i][0] = "Womens";
      }
    } else {
      if (genderValues[i][0] === "Unisex" || genderValues[i][0] === "Male") {
        genderValues[i][0] = "Boys";
      } else {
        genderValues[i][0] = "Girls";
      }
    }
  }

  genderColumn.setValues(genderValues);

  //CHANGE TITLE VALUES TO PROPER CASE
    let columnJ = sheet.getRange(`J2:J${lastRow + 1}`);
    let columnJValues = columnJ.getValues();

    for (let i = 0; i < columnJValues.length; i++) {
      columnJValues[i][0] = toProperCase(columnJValues[i][0]);
    }

    columnJ.setValues(columnJValues);

    // Function to convert a string to proper case
    function toProperCase(str: string): string {
      return str
        .toLowerCase()
        .replace(/\b\w/g, (char) => char.toUpperCase());
    }

  //UPDATING COUNTRY OF ORIGIN VALUES
    //Create and Rename Columns
    sheet.getRange("L:L").insert(ExcelScript.InsertShiftDirection.right);
    sheet.getRange("L:L").insert(ExcelScript.InsertShiftDirection.right);
    sheet.getRange("L1").setValue("Country of Origin Long");
    sheet.getRange("M1").setValue("Tariff");

  // Apply VLOOKUP formula to each cell in the new column
  for (let i = 2; i <= lastRow + 1; i++) {
    let cellAddress = `L${i}`;
    let lookupCell = `K${i}`;
    let formula = `=VLOOKUP(${lookupCell}, COO!A:B, 2, FALSE)`;
    sheet.getRange(cellAddress).setFormula(formula);
  }

  // Set specific values for M2, AC2, AD2, and AE2 down to the last row
  for (let i = 2; i <= lastRow + 1; i++) {
    sheet.getRange(`M${i}`).setValue("9825.10.0000");
    sheet.getRange(`AC${i}`).setValue("POUND");
    sheet.getRange(`AD${i}`).setValue(1);
    sheet.getRange(`AE${i}`).setValue(1);
  }

  // UPDATE MANCOLOR VALUES
  
    // Replacing - with / and converting to proper case in column N
    let columnN = sheet.getRange(`N2:N${lastRow + 1}`);
    let columnNValues = columnN.getValues();

    for (let i = 0; i < columnNValues.length; i++) {
      columnNValues[i][0] = toProperCase(columnNValues[i][0].replace(/-/g, '/'));
    }

    // Proper Title for Mancolor

    for (let i = 0; i < columnNValues.length; i++) {
      columnNValues[i][0] = toProperCase(columnNValues[i][0]);
    }

    columnN.setValues(columnNValues);

}

