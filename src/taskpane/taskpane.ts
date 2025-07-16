/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office, HTMLInputElement, HTMLSelectElement, setTimeout */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("add-charge-column").onclick = addChargeColumn;
  }
});

export async function addChargeColumn() {
  try {
    await Excel.run(async (context) => {
      // Get the active worksheet
      const worksheet = context.workbook.worksheets.getActiveWorksheet();

      // Get user inputs
      const columnHeader =
        (document.getElementById("column-header") as HTMLInputElement).value || "Charge";
      const columnPosition = (document.getElementById("column-position") as HTMLSelectElement)
        .value;

      // Get the used range to find the data
      const usedRange = worksheet.getUsedRange();
      usedRange.load(["rowCount", "columnCount", "values"]);

      await context.sync();

      if (!usedRange) {
        showMessage("No data found in the worksheet.", "error");
        return;
      }

      let targetColumn: Excel.Range;
      let columnLetter: string;

      if (columnPosition === "next") {
        // Find the next available column
        columnLetter = getColumnLetter(usedRange.columnCount + 1);
        targetColumn = worksheet.getRange(`${columnLetter}:${columnLetter}`);
      } else {
        // Use the specified column
        columnLetter = columnPosition;
        targetColumn = worksheet.getRange(`${columnLetter}:${columnLetter}`);
      }

      // Set the header in the first row
      const headerCell = worksheet.getRange(`${columnLetter}1`);
      headerCell.values = [[columnHeader]];
      headerCell.format.font.bold = true;
      headerCell.format.fill.color = "#4472C4";
      headerCell.format.font.color = "white";

      // Add data validation for the charge column (from row 2 onwards)
      const dataRange = worksheet.getRange(`${columnLetter}2:${columnLetter}${usedRange.rowCount}`);

      // Apply data validation to restrict to Y, N, or Q
      dataRange.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: "Y,N,Q",
        },
      };

      // Set default values to "Q" (Query)
      const defaultValues = [];
      for (let i = 0; i < usedRange.rowCount - 1; i++) {
        defaultValues.push(["Q"]);
      }
      dataRange.values = defaultValues;

      // Format the data cells
      dataRange.format.horizontalAlignment = "Center";

      // Auto-fit the column width
      targetColumn.format.autofitColumns();

      await context.sync();

      showMessage(
        `Successfully added '${columnHeader}' column at column ${columnLetter} with Y/N/Q options.`,
        "success"
      );
    });
  } catch (error) {
    console.error(error);
    showMessage("An error occurred: " + error.message, "error");
  }
}

function getColumnLetter(columnNumber: number): string {
  let columnLetter = "";
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnLetter;
}

function showMessage(message: string, type: string) {
  const messageDiv = document.getElementById("message");
  const messageText = document.getElementById("message-text");

  messageText.textContent = message;
  messageDiv.style.display = "block";

  // Style based on type
  if (type === "error") {
    messageDiv.className = "ms-MessageBar ms-MessageBar--error";
  } else {
    messageDiv.className = "ms-MessageBar ms-MessageBar--success";
  }

  // Hide message after 5 seconds
  setTimeout(() => {
    messageDiv.style.display = "none";
  }, 5000);
}
