function copyDropdownOptionsToCells() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getRange("A1"); // Cell with dropdown
  const rule = cell.getDataValidation();

  if (!rule) {
    SpreadsheetApp.getUi().alert("A1 has no data validation rule.");
    return;
  }

  const values = rule.getCriteriaValues()[0]; // Get dropdown values

  // Paste each option into B1 to B70
  for (let i = 0; i < values.length; i++) {
    sheet.getRange(i + 1, 2).setValue(values[i]); // Column B = 2
  }

  SpreadsheetApp.getUi().alert("Dropdown options copied to B1:B" + values.length);
}
