/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office, HTMLInputElement, HTMLSelectElement, setTimeout, HTMLInputElement */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("format-spreadsheet").onclick = formatSpreadsheet;
    document.getElementById("add-charge-column").onclick = addChargeColumn;
    document.getElementById("color-code-rows").onclick = colorCodeRows;

    // Toggle prepopulation rules visibility
    const prepopulateCheckbox = document.getElementById("prepopulate-charge") as HTMLInputElement;
    prepopulateCheckbox.onchange = () => {
      const rulesDiv = document.getElementById("prepopulate-rules");
      rulesDiv.style.display = prepopulateCheckbox.checked ? "block" : "none";
    };

    // Toggle format options visibility
    const formatOptionsCheckbox = document.getElementById(
      "show-format-options"
    ) as HTMLInputElement;
    formatOptionsCheckbox.onchange = () => {
      const optionsDiv = document.getElementById("format-options");
      optionsDiv.style.display = formatOptionsCheckbox.checked ? "block" : "none";
    };
  }
});

export async function formatSpreadsheet() {
  try {
    await Excel.run(async (context) => {
      // Get the active worksheet
      const worksheet = context.workbook.worksheets.getActiveWorksheet();

      // Get the used range
      const usedRange = worksheet.getUsedRange();
      usedRange.load(["rowCount", "columnCount"]);

      await context.sync();

      if (!usedRange) {
        showMessage("No data found in the worksheet to format.", "error");
        return;
      }

      // Get user configuration values
      const headerBgColor = (document.getElementById("header-bg-color") as HTMLInputElement).value;
      const headerTextColor = (document.getElementById("header-text-color") as HTMLInputElement)
        .value;
      const altRowColor1 = (document.getElementById("alt-row-color1") as HTMLInputElement).value;
      const altRowColor2 = (document.getElementById("alt-row-color2") as HTMLInputElement).value;
      const borderColor = (document.getElementById("border-color") as HTMLInputElement).value;
      const maxColumnWidth = parseInt(
        (document.getElementById("max-column-width") as HTMLInputElement).value,
        10
      );
      const enableAlternatingRows = (
        document.getElementById("enable-alternating-rows") as HTMLInputElement
      ).checked;

      // Format headers (first row)
      const headerRow = usedRange.getRow(0);
      headerRow.format.font.bold = true;
      headerRow.format.fill.color = headerBgColor;
      headerRow.format.font.color = headerTextColor;
      headerRow.format.horizontalAlignment = "Center";

      // Add borders to entire used range
      const borderItems = [
        "EdgeTop",
        "EdgeBottom",
        "EdgeLeft",
        "EdgeRight",
        "InsideHorizontal",
        "InsideVertical",
      ];
      borderItems.forEach((item) => {
        usedRange.format.borders.getItem(item).style = "Continuous";
        usedRange.format.borders.getItem(item).color = borderColor;
      });

      // Auto-fit columns first
      const columns = usedRange.getEntireColumn();
      columns.format.autofitColumns();

      await context.sync();

      // Load all column widths at once to avoid sync in loop
      const columnWidths = [];
      for (let col = 0; col < usedRange.columnCount; col++) {
        const column = usedRange.getColumn(col);
        column.load("format/columnWidth");
        columnWidths.push(column);
      }

      await context.sync();

      // Now process columns that exceed max width
      for (let col = 0; col < columnWidths.length; col++) {
        const column = columnWidths[col];
        if (column.format.columnWidth > maxColumnWidth) {
          column.format.columnWidth = maxColumnWidth;
          column.format.wrapText = true;

          // Adjust row height to accommodate wrapped text
          const dataRows = usedRange
            .getOffsetRange(1, 0)
            .getResizedRange(usedRange.rowCount - 1, 0);
          dataRows.format.rowHeight = 30; // Minimum row height for wrapped text
        }
      }

      // Apply alternating row colors if enabled (skip header row)
      if (enableAlternatingRows) {
        for (let row = 1; row < usedRange.rowCount; row++) {
          const dataRow = usedRange.getRow(row);
          if (row % 2 === 0) {
            dataRow.format.fill.color = altRowColor2; // Even rows
          } else {
            dataRow.format.fill.color = altRowColor1; // Odd rows
          }
        }
      }

      await context.sync();

      showMessage(
        "Spreadsheet formatted successfully with your custom styling options.",
        "success"
      );
    });
  } catch (error) {
    console.error(error);
    showMessage("An error occurred while formatting: " + error.message, "error");
  }
}

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
      const shouldPrepopulate = (document.getElementById("prepopulate-charge") as HTMLInputElement)
        .checked;

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

      // Set values based on prepopulation setting
      let values = [];
      if (shouldPrepopulate) {
        // Find narrative/description column
        const narrativeColumnIndex = findNarrativeColumn(usedRange.values[0]);

        if (narrativeColumnIndex !== -1) {
          values = prepopulateChargeValues(usedRange.values, narrativeColumnIndex);
        } else {
          // Default to Q if no narrative column found
          for (let i = 0; i < usedRange.rowCount - 1; i++) {
            values.push(["Q"]);
          }
          showMessage(
            "No Narrative/Description column found. Defaulting to 'Q' values.",
            "warning"
          );
        }
      } else {
        // Default to "Q" (Query) when not prepopulating
        for (let i = 0; i < usedRange.rowCount - 1; i++) {
          values.push(["Q"]);
        }
      }
      dataRange.values = values;

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

function findNarrativeColumn(headerRow: any[]): number {
  const narrativeKeywords = ["narrative", "description", "desc", "notes", "comment", "details"];

  for (let i = 0; i < headerRow.length; i++) {
    if (headerRow[i]) {
      const headerText = headerRow[i].toString().toLowerCase();
      if (narrativeKeywords.some((keyword) => headerText.includes(keyword))) {
        return i;
      }
    }
  }
  return -1;
}

function prepopulateChargeValues(allValues: any[][], narrativeColumnIndex: number): string[][] {
  const noChargeKeywords = (document.getElementById("no-charge-keywords") as HTMLInputElement).value
    .toLowerCase()
    .split(",")
    .map((k) => k.trim())
    .filter((k) => k.length > 0);

  const values: string[][] = [];

  // Skip header row (start from index 1)
  for (let row = 1; row < allValues.length; row++) {
    const narrativeText = allValues[row][narrativeColumnIndex]?.toString().toLowerCase() || "";

    let chargeValue = "Y"; // Default to Yes (chargeable)

    // If narrative text is empty or only whitespace, mark as Query
    if (narrativeText.trim() === "") {
      chargeValue = "Q";
    }
    // Check for no-charge keywords
    else if (noChargeKeywords.some((keyword) => narrativeText.includes(keyword))) {
      chargeValue = "N";
    }
    // Otherwise, default to Y (chargeable)

    values.push([chargeValue]);
  }

  return values;
}

function showMessage(message: string, type: string) {
  const messageDiv = document.getElementById("message");
  const messageText = document.getElementById("message-text");

  messageText.textContent = message;
  messageDiv.style.display = "block";

  // Style based on type
  if (type === "error") {
    messageDiv.className = "ms-MessageBar ms-MessageBar--error";
  } else if (type === "warning") {
    messageDiv.className = "ms-MessageBar ms-MessageBar--warning";
  } else {
    messageDiv.className = "ms-MessageBar ms-MessageBar--success";
  }

  // Hide message after 5 seconds
  setTimeout(() => {
    messageDiv.style.display = "none";
  }, 5000);
}

export async function colorCodeRows() {
  try {
    await Excel.run(async (context) => {
      // Get the active worksheet
      const worksheet = context.workbook.worksheets.getActiveWorksheet();

      // Get the used range
      const usedRange = worksheet.getUsedRange();
      usedRange.load(["rowCount", "columnCount", "values"]);

      await context.sync();

      if (!usedRange) {
        showMessage("No data found in the worksheet.", "error");
        return;
      }

      // Find the Charge column
      const headerRow = usedRange.values[0];
      let chargeColumnIndex = -1;

      for (let i = 0; i < headerRow.length; i++) {
        if (headerRow[i] && headerRow[i].toString().toLowerCase().includes("charge")) {
          chargeColumnIndex = i;
          break;
        }
      }

      if (chargeColumnIndex === -1) {
        showMessage("No 'Charge' column found. Please add a Charge column first.", "error");
        return;
      }

      // Apply color coding to each row based on the charge value
      const values = usedRange.values;
      for (let row = 1; row < usedRange.rowCount; row++) {
        const chargeValue = values[row][chargeColumnIndex];
        const rowRange = usedRange.getRow(row);

        if (chargeValue === "Y") {
          // Pale green for Yes
          rowRange.format.fill.color = "#D4EDDA";
        } else if (chargeValue === "N") {
          // Pale red for No
          rowRange.format.fill.color = "#F8D7DA";
        } else if (chargeValue === "Q") {
          // Pale amber/yellow for Query
          rowRange.format.fill.color = "#FFF3CD";
        }
      }

      await context.sync();

      showMessage("Successfully applied color coding to rows based on Charge values.", "success");
    });
  } catch (error) {
    console.error(error);
    showMessage("An error occurred: " + error.message, "error");
  }
}
