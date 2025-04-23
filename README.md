# Extract Drop-Down Options to Cells in Google Sheets

This script is designed to extract the options from a drop-down menu in a specified cell of a Google Sheet and copy them into a range of cells in the same sheet.

## Functionality

The function `copyDropdownOptionsToCells` performs the following steps:

1. Identifies the active Google Sheet and selects the specified cell (`A1` in this case) that contains a drop-down menu.
2. Verifies if the cell has a data validation rule (drop-down menu). If no data validation rule exists, an alert is displayed.
3. Retrieves the values or options from the drop-down menu.
4. Copies the extracted options into a specified column (`B1` to `B70` in this example).
5. Displays a success message with the number of options copied.

## Prerequisites

- A Google Sheet with a drop-down menu configured in cell `A1`.
- Google Apps Script editor access to copy and paste the function.

## How to Use

1. Open the Google Sheet where you want to use this script.
2. Go to `Extensions` > `Apps Script`.
3. Copy and paste the following code into the script editor:

   ```javascript
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
