/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office, HTMLInputElement, HTMLSelectElement, HTMLElement, setTimeout, localStorage */

// Track whether a matter is currently loaded
let currentMatterLoaded: string | null = null;

// Undo functionality
interface UndoSnapshot {
  timestamp: number;
  changes: Array<{
    row: number;
    column: number;
    oldValue: any;
    newValue: any;
  }>;
}

let lastUndoSnapshot: UndoSnapshot | null = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("apply-formatting").onclick = applyFormatting;
    document.getElementById("apply-all-rules").onclick = applyAllRules;

    // Matter profile functionality
    document.getElementById("save-matter").onclick = saveMatterProfile;
    document.getElementById("delete-matter").onclick = deleteMatterProfile;
    document.getElementById("save-current-settings").onclick = saveCurrentSettings;

    // Fee Earners functionality
    document.getElementById("add-fee-earner").onclick = () => addFeeEarnerRow();
    document.getElementById("update-from-spreadsheet").onclick = updateFeeEarnersFromSpreadsheet;
    document.getElementById("save-participants").onclick = saveParticipants;

    // Rules functionality
    document.getElementById("save-rule-settings").onclick = saveRuleSettings;
    document.getElementById("undo-name-rules").onclick = undoNameStandardisation;

    // Nickname database functionality
    document.getElementById("add-nickname").onclick = addNicknameEntry;
    document.getElementById("reset-nicknames").onclick = resetNicknamesToDefault;

    // Nickname database toggle
    const nicknameToggle = document.getElementById("use-nickname-database") as HTMLInputElement;
    nicknameToggle.onchange = () => {
      const configDiv = document.getElementById("nickname-database-config");
      configDiv.style.display = nicknameToggle.checked ? "block" : "none";
    };

    // Name Standardisation rule toggle
    const nameStandardisationToggle = document.getElementById(
      "name-standardisation-enabled"
    ) as HTMLInputElement;
    nameStandardisationToggle.onchange = () => {
      const configDiv = document.getElementById("name-standardisation-content");
      configDiv.style.display = nameStandardisationToggle.checked ? "block" : "none";
    };

    // Missing Time Entries rule toggle
    const missingTimeEntriesToggle = document.getElementById(
      "missing-time-entries-enabled"
    ) as HTMLInputElement;
    missingTimeEntriesToggle.onchange = () => {
      const configDiv = document.getElementById("missing-time-entries-content");
      configDiv.style.display = missingTimeEntriesToggle.checked ? "block" : "none";
    };

    // Make functions available globally for onclick handlers
    (window as any).removeFeeEarnerRow = removeFeeEarnerRow;
    (window as any).removeNicknameEntry = removeNicknameEntry;

    // Handle matter selection from dropdown
    const matterSelect = document.getElementById("matter-select") as HTMLSelectElement;
    matterSelect.onchange = () => {
      if (matterSelect.value === "__new__") {
        // Show new matter creation section and switch to Settings tab
        document.getElementById("new-matter-section").style.display = "block";
        switchToSettingsTab();
        // Hide Quick Actions section
        document.getElementById("quick-actions-section").style.display = "none";
      } else if (matterSelect.value) {
        // Load existing matter profile
        loadMatterProfile();
        // Hide new matter creation section
        document.getElementById("new-matter-section").style.display = "none";
      } else {
        // Hide Quick Actions section when no matter is selected
        document.getElementById("quick-actions-section").style.display = "none";
        // Hide new matter creation section
        document.getElementById("new-matter-section").style.display = "none";
      }
    };

    // Toggle prepopulation rules visibility
    const prepopulateCheckbox = document.getElementById("prepopulate-charge") as HTMLInputElement;
    prepopulateCheckbox.onchange = () => {
      const rulesDiv = document.getElementById("prepopulate-rules");
      rulesDiv.style.display = prepopulateCheckbox.checked ? "block" : "none";
    };

    // Load saved matter profiles on startup
    loadMatterProfiles();

    // Initialize fee earners table with one empty row
    loadFeeEarnersTable([]);

    // Initialize rules with defaults
    loadRulesConfig(getDefaultRules());

    // Initialize nickname database
    loadNicknameDatabase({});

    // Initialize undo button state
    updateUndoButtonState();

    // Reset table scroll position
    setTimeout(() => resetTableScroll(), 200);

    // Tab switching functionality
    const tabButtons = document.querySelectorAll(".tab-button");
    const tabContents = document.querySelectorAll(".tab-content");

    tabButtons.forEach((button) => {
      button.addEventListener("click", () => {
        const targetTab = button.getAttribute("data-tab");

        // Remove active class from all buttons and contents
        tabButtons.forEach((btn) => btn.classList.remove("active"));
        tabContents.forEach((content) => content.classList.remove("active"));

        // Add active class to clicked button and corresponding content
        button.classList.add("active");
        document.getElementById(`${targetTab}-tab`).classList.add("active");
      });
    });

    // Dropdown functionality
    const dropdownHeaders = document.querySelectorAll(".dropdown-header");
    dropdownHeaders.forEach((header) => {
      header.addEventListener("click", () => {
        const targetId = header.getAttribute("data-target");
        const content = document.getElementById(targetId);
        const arrow = header.querySelector(".dropdown-arrow");

        if (content.style.display === "none") {
          content.style.display = "block";
          header.classList.remove("collapsed");
          arrow.textContent = "▼";
        } else {
          content.style.display = "none";
          header.classList.add("collapsed");
          arrow.textContent = "▶";
        }
      });
    });
  }
});

// Combined formatting function that applies all formatting operations
async function applyFormatting() {
  try {
    // Execute all three formatting operations in sequence
    await formatSpreadsheet();
    await addColumns();
    await colorCodeRows();

    showMessage("Formatting applied successfully.", "success");
  } catch (error) {
    console.error("Error applying formatting:", error);
    showMessage("An error occurred while applying formatting: " + error.message, "error");
  }
}

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
      const verticalAlignment = (document.getElementById("vertical-alignment") as HTMLSelectElement)
        .value;

      // Format headers (first row)
      const headerRow = usedRange.getRow(0);
      headerRow.format.font.bold = true;
      headerRow.format.fill.color = headerBgColor;
      headerRow.format.font.color = headerTextColor;
      headerRow.format.horizontalAlignment = "Center";
      headerRow.format.verticalAlignment = verticalAlignment;

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

      // Apply vertical alignment to all data rows
      const dataRows = usedRange.getOffsetRange(1, 0).getResizedRange(usedRange.rowCount - 1, 0);
      dataRows.format.verticalAlignment = verticalAlignment;

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

async function addChargeColumnInternal(worksheet: Excel.Worksheet, usedRange: Excel.Range) {
  // Get user inputs
  const columnHeader =
    (document.getElementById("column-header") as HTMLInputElement).value || "Charge";
  const columnPosition = (document.getElementById("column-position") as HTMLSelectElement).value;
  const shouldPrepopulate = (document.getElementById("prepopulate-charge") as HTMLInputElement)
    .checked;

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

  // Get current matter settings for formatting
  const headerBgColor = (document.getElementById("header-bg-color") as HTMLInputElement).value;
  const headerTextColor = (document.getElementById("header-text-color") as HTMLInputElement).value;
  const borderColor = (document.getElementById("border-color") as HTMLInputElement).value;

  // Set the header in the first row with matter settings
  const headerCell = worksheet.getRange(`${columnLetter}1`);
  headerCell.values = [[columnHeader]];
  headerCell.format.font.bold = true;
  headerCell.format.fill.color = headerBgColor;
  headerCell.format.font.color = headerTextColor;
  headerCell.format.horizontalAlignment = "Center";

  // Apply borders to header cell
  const headerBorderItems = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
  headerBorderItems.forEach((item) => {
    headerCell.format.borders.getItem(item).style = "Continuous";
    headerCell.format.borders.getItem(item).color = borderColor;
  });

  // Add data validation for the charge column (from row 2 onwards)
  const dataRange = worksheet.getRange(`${columnLetter}2:${columnLetter}${usedRange.rowCount}`);

  // Apply data validation to restrict to Y, N, or Q
  dataRange.dataValidation.rule = {
    list: {
      inCellDropDown: true,
      source: "Y,N,Q",
    },
  };

  // Apply borders to data cells
  const dataBorderItems = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
  dataBorderItems.forEach((item) => {
    dataRange.format.borders.getItem(item).style = "Continuous";
    dataRange.format.borders.getItem(item).color = borderColor;
  });

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
      showMessage("No Narrative/Description column found. Defaulting to 'Q' values.", "warning");
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

  // Apply alternating row colors if enabled
  const enableAlternatingRows = (
    document.getElementById("enable-alternating-rows") as HTMLInputElement
  ).checked;
  if (enableAlternatingRows) {
    const altRowColor1 = (document.getElementById("alt-row-color1") as HTMLInputElement).value;
    const altRowColor2 = (document.getElementById("alt-row-color2") as HTMLInputElement).value;

    // Apply alternating colors to each row in the charge column
    for (let row = 2; row <= usedRange.rowCount; row++) {
      const cell = worksheet.getRange(`${columnLetter}${row}`);
      if (row % 2 === 0) {
        cell.format.fill.color = altRowColor2; // Even rows
      } else {
        cell.format.fill.color = altRowColor1; // Odd rows
      }
    }
  }

  // Auto-fit the column width
  targetColumn.format.autofitColumns();
}

async function addChargeColumnAtPosition(
  worksheet: Excel.Worksheet,
  usedRange: Excel.Range,
  columnNumber: number
) {
  // Get user inputs
  const columnHeader =
    (document.getElementById("column-header") as HTMLInputElement).value || "Charge";
  const shouldPrepopulate = (document.getElementById("prepopulate-charge") as HTMLInputElement)
    .checked;

  // Calculate column letter for the specified position
  const columnLetter = getColumnLetter(columnNumber);
  const targetColumn = worksheet.getRange(`${columnLetter}:${columnLetter}`);

  // Get current matter settings for formatting
  const headerBgColor = (document.getElementById("header-bg-color") as HTMLInputElement).value;
  const headerTextColor = (document.getElementById("header-text-color") as HTMLInputElement).value;
  const borderColor = (document.getElementById("border-color") as HTMLInputElement).value;

  // Set the header in the first row with matter settings
  const headerCell = worksheet.getRange(`${columnLetter}1`);
  headerCell.values = [[columnHeader]];
  headerCell.format.font.bold = true;
  headerCell.format.fill.color = headerBgColor;
  headerCell.format.font.color = headerTextColor;
  headerCell.format.horizontalAlignment = "Center";

  // Apply borders to header cell
  const headerBorderItems = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
  headerBorderItems.forEach((item) => {
    headerCell.format.borders.getItem(item).style = "Continuous";
    headerCell.format.borders.getItem(item).color = borderColor;
  });

  // Add data validation for the charge column (from row 2 onwards)
  const dataRange = worksheet.getRange(`${columnLetter}2:${columnLetter}${usedRange.rowCount}`);

  // Apply data validation to restrict to Y, N, or Q
  dataRange.dataValidation.rule = {
    list: {
      inCellDropDown: true,
      source: "Y,N,Q",
    },
  };

  // Apply borders to data cells
  const dataBorderItems = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
  dataBorderItems.forEach((item) => {
    dataRange.format.borders.getItem(item).style = "Continuous";
    dataRange.format.borders.getItem(item).color = borderColor;
  });

  // Set values based on prepopulation setting
  let values = [];
  if (shouldPrepopulate) {
    // Find narrative/description column in the updated range
    const narrativeColumnIndex = findNarrativeColumn(usedRange.values[0]);

    if (narrativeColumnIndex !== -1) {
      values = prepopulateChargeValues(usedRange.values, narrativeColumnIndex);
    } else {
      // Default to Q if no narrative column found
      for (let i = 0; i < usedRange.rowCount - 1; i++) {
        values.push(["Q"]);
      }
      showMessage("No Narrative/Description column found. Defaulting to 'Q' values.", "warning");
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

  // Apply alternating row colors if enabled
  const enableAlternatingRows = (
    document.getElementById("enable-alternating-rows") as HTMLInputElement
  ).checked;
  if (enableAlternatingRows) {
    const altRowColor1 = (document.getElementById("alt-row-color1") as HTMLInputElement).value;
    const altRowColor2 = (document.getElementById("alt-row-color2") as HTMLInputElement).value;

    // Apply alternating colors to each row in the charge column
    for (let row = 2; row <= usedRange.rowCount; row++) {
      const cell = worksheet.getRange(`${columnLetter}${row}`);
      if (row % 2 === 0) {
        cell.format.fill.color = altRowColor2; // Even rows
      } else {
        cell.format.fill.color = altRowColor1; // Odd rows
      }
    }
  }

  // Auto-fit the column width
  targetColumn.format.autofitColumns();
}

export async function addChargeColumn() {
  try {
    await Excel.run(async (context) => {
      // Get the active worksheet
      const worksheet = context.workbook.worksheets.getActiveWorksheet();

      // Get the used range to find the data
      const usedRange = worksheet.getUsedRange();
      usedRange.load(["rowCount", "columnCount", "values"]);

      await context.sync();

      if (!usedRange) {
        showMessage("No data found in the worksheet.", "error");
        return;
      }

      await addChargeColumnInternal(worksheet, usedRange);

      await context.sync();

      const columnHeader =
        (document.getElementById("column-header") as HTMLInputElement).value || "Charge";
      const columnPosition = (document.getElementById("column-position") as HTMLSelectElement)
        .value;

      const columnLetter =
        columnPosition === "next" ? getColumnLetter(usedRange.columnCount + 1) : columnPosition;

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

function findTimeColumn(headerRow: any[]): number {
  const timeKeywords = ["time", "hours", "mins", "minutes", "duration"];

  for (let i = 0; i < headerRow.length; i++) {
    if (headerRow[i]) {
      const headerText = headerRow[i].toString().toLowerCase();
      if (timeKeywords.some((keyword) => headerText.includes(keyword))) {
        return i;
      }
    }
  }
  return -1;
}

function findNameColumn(headerRow: any[]): number {
  const nameKeywords = ["name", "fee earner", "lawyer", "attorney", "solicitor", "person", "who"];

  for (let i = 0; i < headerRow.length; i++) {
    if (headerRow[i]) {
      const headerText = headerRow[i].toString().toLowerCase();
      if (nameKeywords.some((keyword) => headerText.includes(keyword))) {
        return i;
      }
    }
  }
  return -1;
}

function findNotesColumn(headerRow: any[]): number {
  const notesKeywords = ["notes", "note", "rules applied", "tracking"];

  for (let i = 0; i < headerRow.length; i++) {
    if (headerRow[i]) {
      const headerText = headerRow[i].toString().toLowerCase();
      if (notesKeywords.some((keyword) => headerText.includes(keyword))) {
        return i;
      }
    }
  }
  return -1;
}

async function createNotesColumn(worksheet: Excel.Worksheet, insertAfterColumn: number) {
  // Insert new column
  const insertColumn = worksheet.getCell(0, insertAfterColumn + 1).getEntireColumn();
  insertColumn.insert(Excel.InsertShiftDirection.right);

  // Set header
  const headerCell = worksheet.getCell(0, insertAfterColumn + 1);
  headerCell.values = [["Notes"]];

  // Apply header formatting
  const profiles = getMatterProfiles();
  const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;
  const currentProfile = profiles.find((p) => p.name === selectedMatter);

  if (currentProfile) {
    const headerBgColor = currentProfile.headerBgColor || "#4472C4";
    const headerTextColor = currentProfile.headerTextColor || "#FFFFFF";
    const borderColor = currentProfile.borderColor || "#D1D5DB";

    headerCell.format.fill.color = headerBgColor;
    headerCell.format.font.color = headerTextColor;
    headerCell.format.font.bold = true;
    headerCell.format.horizontalAlignment = "Center";
    headerCell.format.verticalAlignment = "Center";

    // Apply borders
    const borderItems = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
    borderItems.forEach((item) => {
      headerCell.format.borders.getItem(item).style = "Continuous";
      headerCell.format.borders.getItem(item).color = borderColor;
    });
  }

  return insertAfterColumn + 1;
}

async function createNotesColumnWithFormatting(
  worksheet: Excel.Worksheet,
  usedRange: Excel.Range
): Promise<number> {
  const insertAfterColumn = usedRange.columnCount - 1;

  // Insert new column at the far right
  const insertColumn = worksheet.getCell(0, usedRange.columnCount).getEntireColumn();
  insertColumn.insert(Excel.InsertShiftDirection.right);

  const notesColumnIndex = usedRange.columnCount;

  // Set header
  const headerCell = worksheet.getCell(0, notesColumnIndex);
  headerCell.values = [["Notes"]];

  // Apply formatting that matches the current matter profile
  const profiles = getMatterProfiles();
  const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;
  const currentProfile = profiles.find((p) => p.name === selectedMatter);

  if (currentProfile) {
    const headerBgColor = currentProfile.headerBgColor || "#4472C4";
    const headerTextColor = currentProfile.headerTextColor || "#FFFFFF";
    const borderColor = currentProfile.borderColor || "#D1D5DB";
    const altRowColor1 = currentProfile.altRowColor1 || "#FFFFFF";
    const altRowColor2 = currentProfile.altRowColor2 || "#F8F9FA";
    const enableAlternatingRows = currentProfile.enableAlternatingRows !== false;
    const verticalAlignment = currentProfile.verticalAlignment || "center";

    // Format header
    headerCell.format.fill.color = headerBgColor;
    headerCell.format.font.color = headerTextColor;
    headerCell.format.font.bold = true;
    headerCell.format.horizontalAlignment = "Center";
    headerCell.format.verticalAlignment =
      verticalAlignment === "center" ? "Center" : verticalAlignment === "top" ? "Top" : "Bottom";

    // Apply header borders
    const borderItems = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
    borderItems.forEach((item) => {
      headerCell.format.borders.getItem(item).style = "Continuous";
      headerCell.format.borders.getItem(item).color = borderColor;
    });

    // Format data cells (rows 2 onwards)
    if (usedRange.rowCount > 1) {
      const dataRange = worksheet.getRangeByIndexes(1, notesColumnIndex, usedRange.rowCount - 1, 1);

      // Apply borders to all data cells
      borderItems.forEach((item) => {
        dataRange.format.borders.getItem(item).style = "Continuous";
        dataRange.format.borders.getItem(item).color = borderColor;
      });

      // Apply vertical alignment
      dataRange.format.verticalAlignment =
        verticalAlignment === "center" ? "Center" : verticalAlignment === "top" ? "Top" : "Bottom";
      dataRange.format.horizontalAlignment = "Left";

      // Apply alternating row colors if enabled
      if (enableAlternatingRows) {
        for (let row = 1; row < usedRange.rowCount; row++) {
          const cell = worksheet.getCell(row, notesColumnIndex);
          if (row % 2 === 0) {
            cell.format.fill.color = altRowColor2;
          } else {
            cell.format.fill.color = altRowColor1;
          }
        }
      }
    }
  }

  return notesColumnIndex;
}

function addNoteToRow(notes: string, newNote: string): string {
  if (!notes || notes.trim() === "") {
    return newNote;
  }

  // Check if note already exists to avoid duplicates
  const existingNotes = notes.split(",").map((n) => n.trim());
  if (existingNotes.includes(newNote)) {
    return notes; // Note already exists
  }

  return notes + ", " + newNote;
}

function findRoleColumn(headerRow: any[]): number {
  const roleKeywords = ["role", "title", "position", "grade", "level", "rank"];

  for (let i = 0; i < headerRow.length; i++) {
    if (headerRow[i]) {
      const headerText = headerRow[i].toString().toLowerCase();
      if (roleKeywords.some((keyword) => headerText.includes(keyword))) {
        return i;
      }
    }
  }
  return -1;
}

function findRateColumn(headerRow: any[]): number {
  const rateKeywords = ["rate", "charge", "cost", "price", "fee", "bill", "amount"];

  for (let i = 0; i < headerRow.length; i++) {
    if (headerRow[i]) {
      const headerText = headerRow[i].toString().toLowerCase();
      if (rateKeywords.some((keyword) => headerText.includes(keyword))) {
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
    // Check for no-charge keywords at the start of the narrative text
    else if (noChargeKeywords.some((keyword) => narrativeText.startsWith(keyword))) {
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

async function addAmendedColumn(
  worksheet: Excel.Worksheet,
  columnIndex: number,
  originalName: string,
  amendedName: string,
  usedRange: Excel.Range,
  insertionOffsetCount: number
): Promise<string> {
  const adjustedIndex = columnIndex + insertionOffsetCount;
  const columnLetter = getColumnLetter(adjustedIndex + 1);
  const amendedColumnLetter = getColumnLetter(adjustedIndex + 2);

  // Get current matter settings for formatting
  const headerBgColor = (document.getElementById("header-bg-color") as HTMLInputElement).value;
  const headerTextColor = (document.getElementById("header-text-color") as HTMLInputElement).value;
  const borderColor = (document.getElementById("border-color") as HTMLInputElement).value;
  const enableAlternatingRows = (
    document.getElementById("enable-alternating-rows") as HTMLInputElement
  ).checked;
  const altRowColor1 = (document.getElementById("alt-row-color1") as HTMLInputElement).value;
  const altRowColor2 = (document.getElementById("alt-row-color2") as HTMLInputElement).value;

  // Rename existing column
  const originalHeaderCell = worksheet.getRange(`${columnLetter}1`);
  originalHeaderCell.values = [[originalName]];

  // Insert new column to the right
  const insertRange = worksheet.getRange(`${amendedColumnLetter}:${amendedColumnLetter}`);
  insertRange.insert(Excel.InsertShiftDirection.right);

  // Add amended header
  const amendedHeaderCell = worksheet.getRange(`${amendedColumnLetter}1`);
  amendedHeaderCell.values = [[amendedName]];
  amendedHeaderCell.format.font.bold = true;
  amendedHeaderCell.format.fill.color = headerBgColor;
  amendedHeaderCell.format.font.color = headerTextColor;
  amendedHeaderCell.format.horizontalAlignment = "Center";

  // Apply borders to header
  const headerBorderItems = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
  headerBorderItems.forEach((item) => {
    amendedHeaderCell.format.borders.getItem(item).style = "Continuous";
    amendedHeaderCell.format.borders.getItem(item).color = borderColor;
  });

  // Format data cells in the new column
  const amendedDataRange = worksheet.getRange(
    `${amendedColumnLetter}2:${amendedColumnLetter}${usedRange.rowCount}`
  );
  const dataBorderItems = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
  dataBorderItems.forEach((item) => {
    amendedDataRange.format.borders.getItem(item).style = "Continuous";
    amendedDataRange.format.borders.getItem(item).color = borderColor;
  });

  // Apply alternating row colors if enabled
  if (enableAlternatingRows) {
    for (let row = 2; row <= usedRange.rowCount; row++) {
      const cell = worksheet.getRange(`${amendedColumnLetter}${row}`);
      if (row % 2 === 0) {
        cell.format.fill.color = altRowColor2;
      } else {
        cell.format.fill.color = altRowColor1;
      }
    }
  }

  // Auto-fit column widths for both original (renamed) and amended columns
  const originalColumn = worksheet.getRange(`${columnLetter}:${columnLetter}`);
  const amendedColumn = worksheet.getRange(`${amendedColumnLetter}:${amendedColumnLetter}`);

  originalColumn.format.autofitColumns();
  amendedColumn.format.autofitColumns();

  return amendedName.split(" ")[1]; // Return "Narrative" or "Time"
}

export async function addColumns() {
  try {
    await Excel.run(async (context) => {
      // Get the active worksheet
      const worksheet = context.workbook.worksheets.getActiveWorksheet();

      // Get the initial used range
      let usedRange = worksheet.getUsedRange();
      usedRange.load(["rowCount", "columnCount", "values"]);

      await context.sync();

      if (!usedRange) {
        showMessage("No data found in the worksheet.", "error");
        return;
      }

      // Get user settings
      const columnHeader = (
        document.getElementById("column-header") as HTMLInputElement
      ).value.trim();
      const prepopulateCharge = (document.getElementById("prepopulate-charge") as HTMLInputElement)
        .checked;
      const shouldAddCharge = columnHeader !== "" || prepopulateCharge;
      const shouldAddAmendedNarrative = (
        document.getElementById("add-amended-narrative") as HTMLInputElement
      ).checked;
      const shouldAddAmendedTime = (document.getElementById("add-amended-time") as HTMLInputElement)
        .checked;
      const shouldAddNotes = (document.getElementById("add-notes-column") as HTMLInputElement)
        .checked;

      // Also check if Name Standardisation is enabled - if so, we should add Notes column
      const nameStandardisationEnabled = (
        document.getElementById("name-standardisation-enabled") as HTMLInputElement
      ).checked;
      const shouldAddNotesForRules = shouldAddNotes || nameStandardisationEnabled;

      const headerRow = usedRange.values[0];
      let processedColumns = [];

      // Find column indices once at the beginning
      const narrativeColumnIndex = findNarrativeColumn(headerRow);
      const timeColumnIndex = findTimeColumn(headerRow);

      // PHASE 1: Process amended columns (right to left)
      const columnsToProcess = [];

      if (shouldAddAmendedNarrative && narrativeColumnIndex !== -1) {
        columnsToProcess.push({
          index: narrativeColumnIndex,
          originalName: "Original Narrative",
          amendedName: "Amended Narrative",
          type: "Narrative",
        });
      }

      if (shouldAddAmendedTime && timeColumnIndex !== -1) {
        columnsToProcess.push({
          index: timeColumnIndex,
          originalName: "Original Time",
          amendedName: "Amended Time",
          type: "Time",
        });
      }

      // Sort by column index descending (process rightmost columns first)
      columnsToProcess.sort((a, b) => b.index - a.index);

      // Process each amended column
      for (const column of columnsToProcess) {
        const columnType = await addAmendedColumn(
          worksheet,
          column.index,
          column.originalName,
          column.amendedName,
          usedRange,
          0 // No offset needed since we process from right to left
        );
        processedColumns.push(columnType);
      }

      // PHASE 2: Add charge column at the far right (if requested)
      if (shouldAddCharge) {
        // After amended columns are added, refresh the used range to get the updated column count
        usedRange = worksheet.getUsedRange();
        usedRange.load(["rowCount", "columnCount", "values"]);
        await context.sync();

        // Add charge column at the next available column (far right)
        await addChargeColumnAtPosition(worksheet, usedRange, usedRange.columnCount + 1);
        processedColumns.push("Charge");
      }

      // PHASE 3: Add Notes column at the far right (if requested or if Name Standardisation is enabled)
      if (shouldAddNotesForRules) {
        // After charge column is added, refresh the used range to get the updated column count
        usedRange = worksheet.getUsedRange();
        usedRange.load(["rowCount", "columnCount", "values"]);
        await context.sync();

        // Check if Notes column already exists
        const updatedHeaderRow = usedRange.values[0];
        const existingNotesIndex = findNotesColumn(updatedHeaderRow);

        if (existingNotesIndex === -1) {
          // Create Notes column at the far right with full formatting
          await createNotesColumnWithFormatting(worksheet, usedRange);
          processedColumns.push("Notes");
        } else {
          showMessage("Notes column already exists in the worksheet.", "info");
        }
      }

      // Final auto-fit pass to ensure all columns are properly sized
      if (processedColumns.length > 0) {
        usedRange = worksheet.getUsedRange();
        usedRange.load(["columnCount"]);
        await context.sync();

        // Auto-fit all columns in the used range to account for any formatting changes
        const allColumns = usedRange.getEntireColumn();
        allColumns.format.autofitColumns();
      }

      await context.sync();

      if (processedColumns.length > 0) {
        const message = `Successfully added ${processedColumns.join(", ")} column${processedColumns.length > 1 ? "s" : ""}.`;
        showMessage(message, "success");
      } else {
        showMessage(
          "No columns were configured to be added. Please check your settings.",
          "warning"
        );
      }
    });
  } catch (error) {
    console.error(error);
    showMessage("An error occurred while adding columns: " + error.message, "error");
  }
}

// Fee Earner Management
interface FeeEarner {
  name: string;
  role: string;
  rate: number;
  email: string;
  billingContact: "Fee Earner" | "Other";
  billingContactName: string;
  billingContactEmail: string;
  isDefaultForName?: boolean;
  nameVariations?: string[];
}

// Rules Management
interface NameStandardisationRule {
  enabled: boolean;
  caseSensitive: boolean;
  allowPartialMatches: boolean;
  useDateMatching: boolean;
  replaceOnlyFirstOccurrence: boolean;
  excludedNames: string[];
  minPartialMatchLength?: number; // Optional for backward compatibility
  useNicknameDatabase?: boolean; // Optional for backward compatibility
  customNicknames?: Record<string, string>; // nickname -> full name mapping
}

interface MissingTimeEntriesRule {
  enabled: boolean;
  dateTolerance: number; // days ±0 for exact match
  meetingKeywords: string[]; // words that indicate meetings/calls
  requireExactTimeMatch: boolean; // optional stricter matching
  createMissingEntries: boolean; // auto-create placeholder entries
}

interface RulesConfig {
  nameStandardisation: NameStandardisationRule;
  missingTimeEntries: MissingTimeEntriesRule;
}

// Built-in nickname database
const DEFAULT_NICKNAMES: Record<string, string> = {
  // Common male nicknames
  bill: "william",
  billy: "william",
  will: "william",
  bob: "robert",
  bobby: "robert",
  rob: "robert",
  robbie: "robert",
  dick: "richard",
  rick: "richard",
  ricky: "richard",
  rich: "richard",
  richie: "richard",
  jim: "james",
  jimmy: "james",
  jamie: "james",
  joe: "joseph",
  joey: "joseph",
  mike: "michael",
  mickey: "michael",
  mick: "michael",
  dave: "david",
  davey: "david",
  dan: "daniel",
  danny: "daniel",
  tom: "thomas",
  tommy: "thomas",
  chris: "christopher",
  matt: "matthew",
  steve: "stephen",
  phil: "philip",
  phil: "phillip",
  tony: "anthony",
  andy: "andrew",
  drew: "andrew",
  nick: "nicholas",
  john: "jonathan",
  johnny: "jonathan",
  ben: "benjamin",
  benny: "benjamin",
  alex: "alexander",
  sam: "samuel",
  sammy: "samuel",
  ed: "edward",
  eddie: "edward",
  ted: "edward",
  teddy: "edward",
  charlie: "charles",
  chuck: "charles",
  tim: "timothy",

  // Common female nicknames
  liz: "elizabeth",
  lizzy: "elizabeth",
  beth: "elizabeth",
  betty: "elizabeth",
  sue: "susan",
  susie: "susan",
  suzy: "susan",
  kate: "katherine",
  katie: "katherine",
  kathy: "katherine",
  kit: "katherine",
  kitty: "katherine",
  jen: "jennifer",
  jenny: "jennifer",
  jess: "jessica",
  jessie: "jessica",
  lisa: "elizabeth",
  mel: "melissa",
  amy: "amanda",
  mandy: "amanda",
  chris: "christine",
  chrissy: "christine",
  tina: "christina",
  cindy: "cynthia",
  patty: "patricia",
  pat: "patricia",
  trish: "patricia",
  nancy: "nan",
  ann: "anne",
  annie: "anne",
  maggie: "margaret",
  meg: "margaret",
  peggy: "margaret",
  carol: "caroline",
  carrie: "caroline",
  julie: "julia",
  jules: "julia",
  marie: "mary",
  sally: "sarah",
  sara: "sarah",
  alex: "alexandra",
  lexi: "alexandra",
  sam: "samantha",
  sammie: "samantha",
};

// Matter Profile Management
interface MatterProfile {
  name: string;
  headerBgColor: string;
  headerTextColor: string;
  altRowColor1: string;
  altRowColor2: string;
  borderColor: string;
  maxColumnWidth: number;
  enableAlternatingRows: boolean;
  verticalAlignment: string;
  columnHeader: string;
  columnPosition: string;
  prepopulateCharge: boolean;
  noChargeKeywords: string;
  addAmendedNarrative: boolean;
  addAmendedTime: boolean;
  addNotesColumn: boolean;
  feeEarners: FeeEarner[];
  rules: RulesConfig;
}

function getCurrentSettings(): MatterProfile {
  return {
    name: "",
    headerBgColor: (document.getElementById("header-bg-color") as HTMLInputElement).value,
    headerTextColor: (document.getElementById("header-text-color") as HTMLInputElement).value,
    altRowColor1: (document.getElementById("alt-row-color1") as HTMLInputElement).value,
    altRowColor2: (document.getElementById("alt-row-color2") as HTMLInputElement).value,
    borderColor: (document.getElementById("border-color") as HTMLInputElement).value,
    maxColumnWidth: parseInt(
      (document.getElementById("max-column-width") as HTMLInputElement).value,
      10
    ),
    enableAlternatingRows: (document.getElementById("enable-alternating-rows") as HTMLInputElement)
      .checked,
    verticalAlignment: (document.getElementById("vertical-alignment") as HTMLSelectElement).value,
    columnHeader: (document.getElementById("column-header") as HTMLInputElement).value,
    columnPosition: (document.getElementById("column-position") as HTMLSelectElement).value,
    prepopulateCharge: (document.getElementById("prepopulate-charge") as HTMLInputElement).checked,
    noChargeKeywords: (document.getElementById("no-charge-keywords") as HTMLInputElement).value,
    addAmendedNarrative: (document.getElementById("add-amended-narrative") as HTMLInputElement)
      .checked,
    addAmendedTime: (document.getElementById("add-amended-time") as HTMLInputElement).checked,
    addNotesColumn: (document.getElementById("add-notes-column") as HTMLInputElement).checked,
    feeEarners: getCurrentFeeEarners(),
    rules: getCurrentRules(),
  };
}

function applySettings(profile: MatterProfile) {
  (document.getElementById("header-bg-color") as HTMLInputElement).value = profile.headerBgColor;
  (document.getElementById("header-text-color") as HTMLInputElement).value =
    profile.headerTextColor;
  (document.getElementById("alt-row-color1") as HTMLInputElement).value = profile.altRowColor1;
  (document.getElementById("alt-row-color2") as HTMLInputElement).value = profile.altRowColor2;
  (document.getElementById("border-color") as HTMLInputElement).value = profile.borderColor;
  (document.getElementById("max-column-width") as HTMLInputElement).value =
    profile.maxColumnWidth.toString();
  (document.getElementById("enable-alternating-rows") as HTMLInputElement).checked =
    profile.enableAlternatingRows;
  (document.getElementById("vertical-alignment") as HTMLSelectElement).value =
    profile.verticalAlignment || "center";

  // Apply charge column settings with backward compatibility defaults
  (document.getElementById("column-header") as HTMLInputElement).value =
    profile.columnHeader || "Charge";
  (document.getElementById("column-position") as HTMLSelectElement).value =
    profile.columnPosition || "next";
  (document.getElementById("prepopulate-charge") as HTMLInputElement).checked =
    profile.prepopulateCharge || false;
  (document.getElementById("no-charge-keywords") as HTMLInputElement).value =
    profile.noChargeKeywords;

  // Apply amended column settings with backward compatibility defaults
  (document.getElementById("add-amended-narrative") as HTMLInputElement).checked =
    profile.addAmendedNarrative || false;
  (document.getElementById("add-amended-time") as HTMLInputElement).checked =
    profile.addAmendedTime || false;
  (document.getElementById("add-notes-column") as HTMLInputElement).checked =
    profile.addNotesColumn || false;

  // Apply fee earners settings with backward compatibility defaults
  const feeEarners = profile.feeEarners || [];
  loadFeeEarnersTable(feeEarners);

  // Apply rules settings with backward compatibility defaults
  const rules = profile.rules || getDefaultRules();
  loadRulesConfig(rules);

  // Update prepopulation rules visibility based on checkbox state
  const rulesDiv = document.getElementById("prepopulate-rules");
  rulesDiv.style.display = profile.prepopulateCharge || false ? "block" : "none";
}

function loadMatterProfiles() {
  const profiles = getMatterProfiles();
  const selectElement = document.getElementById("matter-select") as HTMLSelectElement;

  // Clear existing options and rebuild with default options
  selectElement.innerHTML = '<option value="">-- Select a Matter --</option>';

  // Debug: Show what matters are currently saved (temporary)
  console.log(
    `DEBUG: Loaded ${profiles.length} matter profiles:`,
    profiles.map((p) => p.name).join(", ")
  );

  // If no profiles exist, create a default one
  if (profiles.length === 0) {
    console.log("No matter profiles found - they may have been cleared from localStorage");

    // Create a default "Project Apricot" matter profile to help user get started
    const defaultProfile: MatterProfile = {
      name: "Project Apricot",
      headerBgColor: "#4472C4",
      headerTextColor: "#FFFFFF",
      altRowColor1: "#FFFFFF",
      altRowColor2: "#F8F9FA",
      borderColor: "#D1D5DB",
      enableAlternatingRows: true,
      maxColumnWidth: 300,
      verticalAlignment: "center",
      columnHeader: "Charge",
      columnPosition: "next",
      prepopulateCharge: false,
      noChargeKeywords: ["NC", "DO NOT CHARGE", "Non Chargeable"],
      addAmendedNarrative: false,
      addAmendedTime: false,
      addNotesColumn: true,
      participants: {
        feeEarners: [
          {
            name: "John Smith",
            role: "Partner",
            rate: 500,
            email: "john.smith@example.com",
            useAsDefault: true,
            billingContact: false,
          },
          {
            name: "Jane Doe",
            role: "Associate",
            rate: 300,
            email: "jane.doe@example.com",
            useAsDefault: false,
            billingContact: false,
          },
        ],
      },
      rules: {
        nameStandardisation: {
          enabled: false,
          caseSensitive: false,
          allowPartialMatches: true,
          useDateMatching: true,
          replaceOnlyFirstOccurrence: true,
          excludedNames: [],
          minPartialMatchLength: 3,
          useNicknameDatabase: true,
          customNicknames: {},
        },
        missingTimeEntries: {
          enabled: false,
          dateTolerance: 0,
          meetingKeywords: ["meeting", "call", "conference", "discussion", "telephone", "phone"],
          requireExactTimeMatch: false,
          createMissingEntries: false,
        },
      },
    };

    saveMatterProfiles([defaultProfile]);
    console.log("Created default 'Project Apricot' matter profile");
  }

  // Get current profiles (including any newly created default)
  const currentProfiles = getMatterProfiles();

  // Add all saved profiles to the dropdown
  currentProfiles.forEach((profile) => {
    const option = document.createElement("option");
    option.value = profile.name;
    option.textContent = profile.name;
    selectElement.appendChild(option);
  });

  // Add the "Add New Matter" option at the end
  const addNewOption = document.createElement("option");
  addNewOption.value = "__new__";
  addNewOption.textContent = "+ Add New Matter";
  addNewOption.style.fontStyle = "italic";
  selectElement.appendChild(addNewOption);

  // If a matter is loaded, update the dropdown
  if (currentMatterLoaded) {
    selectElement.value = currentMatterLoaded;
  }
}

function getMatterProfiles(): MatterProfile[] {
  const stored = localStorage.getItem("fixmytime-matter-profiles");
  return stored ? JSON.parse(stored) : [];
}

function saveMatterProfiles(profiles: MatterProfile[]) {
  localStorage.setItem("fixmytime-matter-profiles", JSON.stringify(profiles));
}

function saveMatterProfile() {
  const matterName = (document.getElementById("new-matter-name") as HTMLInputElement).value.trim();

  if (!matterName) {
    showMessage("Please enter a matter name.", "error");
    return;
  }

  const currentSettings = getCurrentSettings();
  currentSettings.name = matterName;

  const profiles = getMatterProfiles();
  const existingIndex = profiles.findIndex((p) => p.name === matterName);

  if (existingIndex >= 0) {
    profiles[existingIndex] = currentSettings;
    showMessage(`Matter profile "${matterName}" updated successfully.`, "success");
  } else {
    profiles.push(currentSettings);
    showMessage(`Matter profile "${matterName}" saved successfully.`, "success");
  }

  saveMatterProfiles(profiles);
  loadMatterProfiles();

  // Clear the input and select the saved matter
  (document.getElementById("new-matter-name") as HTMLInputElement).value = "";
  (document.getElementById("matter-select") as HTMLSelectElement).value = matterName;
}

function loadMatterProfile() {
  const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;

  if (!selectedMatter) {
    showMessage("Please select a matter to load.", "error");
    return;
  }

  const profiles = getMatterProfiles();
  const profile = profiles.find((p) => p.name === selectedMatter);

  if (profile) {
    applySettings(profile);
    currentMatterLoaded = selectedMatter;

    // Show current matter display
    document.getElementById("current-matter-display").style.display = "block";
    document.getElementById("current-matter-name").textContent = selectedMatter;

    // Show Quick Actions section now that a matter is selected
    document.getElementById("quick-actions-section").style.display = "block";

    showMessage(`Matter profile "${selectedMatter}" loaded successfully.`, "success");
  } else {
    showMessage("Matter profile not found.", "error");
  }
}

export function switchToSettingsTab() {
  // Remove active class from all buttons and contents
  const tabButtons = document.querySelectorAll(".tab-button");
  const tabContents = document.querySelectorAll(".tab-content");

  tabButtons.forEach((btn) => btn.classList.remove("active"));
  tabContents.forEach((content) => content.classList.remove("active"));

  // Add active class to settings tab
  const settingsButton = document.querySelector('[data-tab="settings"]') as HTMLElement;
  settingsButton.classList.add("active");
  document.getElementById("settings-tab").classList.add("active");
}

export function showNewMatterSection() {
  document.getElementById("new-matter-section").style.display = "block";
  // Focus on the input field
  (document.getElementById("new-matter-name") as HTMLInputElement).focus();
}

export function hideNewMatterSection() {
  document.getElementById("new-matter-section").style.display = "none";
}

function deleteMatterProfile() {
  const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;

  if (!selectedMatter) {
    showMessage("Please select a matter to delete.", "error");
    return;
  }

  const profiles = getMatterProfiles();
  const filteredProfiles = profiles.filter((p) => p.name !== selectedMatter);

  if (filteredProfiles.length < profiles.length) {
    saveMatterProfiles(filteredProfiles);
    loadMatterProfiles();

    // If the deleted matter was the currently loaded one, reset the UI
    if (currentMatterLoaded === selectedMatter) {
      currentMatterLoaded = null;
      document.getElementById("current-matter-display").style.display = "none";
    }

    showMessage(`Matter profile "${selectedMatter}" deleted successfully.`, "success");
  } else {
    showMessage("Matter profile not found.", "error");
  }
}
function saveCurrentSettings() {
  const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;

  if (!selectedMatter) {
    showMessage(
      "Please select a matter from the dropdown to update, or create a new one in the Matter Profile Management section.",
      "error"
    );
    return;
  }

  const currentSettings = getCurrentSettings();
  currentSettings.name = selectedMatter;

  const profiles = getMatterProfiles();
  const existingIndex = profiles.findIndex((p) => p.name === selectedMatter);

  if (existingIndex >= 0) {
    profiles[existingIndex] = currentSettings;
    saveMatterProfiles(profiles);
    showMessage(
      `Matter profile "${selectedMatter}" updated successfully with current settings.`,
      "success"
    );
  } else {
    showMessage("Selected matter profile not found. Please create a new profile first.", "error");
  }
}

// Fee Earners Management Functions
function getCurrentFeeEarners(): FeeEarner[] {
  const tbody = document.getElementById("fee-earners-tbody");
  if (!tbody) return [];

  const rows = tbody.querySelectorAll("tr");
  const feeEarners: FeeEarner[] = [];

  rows.forEach((row) => {
    const nameInput = row.querySelector(".name-input") as HTMLInputElement;
    const roleInput = row.querySelector(".role-input") as HTMLInputElement;
    const rateInput = row.querySelector(".rate-input") as HTMLInputElement;
    const emailInput = row.querySelector(".email-input") as HTMLInputElement;
    const billingContactSelect = row.querySelector(".billing-contact-select") as HTMLSelectElement;
    const billingContactNameInput = row.querySelector(
      ".billing-contact-name-input"
    ) as HTMLInputElement;
    const billingContactEmailInput = row.querySelector(
      ".billing-contact-email-input"
    ) as HTMLInputElement;

    if (nameInput && roleInput && rateInput && emailInput && billingContactSelect) {
      const useAsDefaultCheckbox = row.querySelector(
        ".use-as-default-checkbox"
      ) as HTMLInputElement;

      feeEarners.push({
        name: nameInput.value.trim(),
        role: roleInput.value.trim(),
        rate: parseFloat(rateInput.value) || 0,
        email: emailInput.value.trim(),
        billingContact: billingContactSelect.value as "Fee Earner" | "Other",
        billingContactName: billingContactNameInput ? billingContactNameInput.value.trim() : "",
        billingContactEmail: billingContactEmailInput ? billingContactEmailInput.value.trim() : "",
        isDefaultForName: useAsDefaultCheckbox ? useAsDefaultCheckbox.checked : false,
      });
    }
  });

  return feeEarners;
}

function loadFeeEarnersTable(feeEarners: FeeEarner[]) {
  const tbody = document.getElementById("fee-earners-tbody");
  if (!tbody) return;

  tbody.innerHTML = "";

  if (feeEarners.length === 0) {
    // Add one empty row if no fee earners exist
    addFeeEarnerRow();
  } else {
    feeEarners.forEach((feeEarner) => {
      addFeeEarnerRowWithData(feeEarner);
    });

    // Run duplicate detection after loading all fee earners
    setTimeout(() => {
      detectDuplicateNames();
      resetTableScroll();
    }, 100);
  }
}

function addFeeEarnerRow() {
  const emptyFeeEarner: FeeEarner = {
    name: "",
    role: "",
    rate: 0,
    email: "",
    billingContact: "Fee Earner",
    billingContactName: "",
    billingContactEmail: "",
  };
  addFeeEarnerRowWithData(emptyFeeEarner);
}

function addFeeEarnerRowWithData(feeEarner: FeeEarner) {
  const tbody = document.getElementById("fee-earners-tbody");
  if (!tbody) return;

  const row = document.createElement("tr");
  const isOtherBilling = feeEarner.billingContact === "Other";

  row.innerHTML = `
    <td><input type="text" class="name-input" value="${feeEarner.name}" placeholder="Enter name"></td>
    <td><input type="text" class="role-input" value="${feeEarner.role}" placeholder="Enter role"></td>
    <td><input type="number" class="rate-input" value="${feeEarner.rate || ""}" placeholder="0.00" step="0.01" min="0"></td>
    <td><input type="email" class="email-input" value="${feeEarner.email}" placeholder="Enter email"></td>
    <td style="text-align: center;">
      <input type="checkbox" class="use-as-default-checkbox" ${feeEarner.isDefaultForName ? "checked" : ""}>
    </td>
    <td>
      <select class="billing-contact-select">
        <option value="Fee Earner" ${feeEarner.billingContact === "Fee Earner" ? "selected" : ""}>Fee Earner</option>
        <option value="Other" ${feeEarner.billingContact === "Other" ? "selected" : ""}>Other</option>
      </select>
    </td>
    <td class="${!isOtherBilling ? "disabled-field" : ""}">
      <input type="text" class="billing-contact-name-input" value="${feeEarner.billingContactName}" 
             placeholder="Enter name" ${!isOtherBilling ? "disabled" : ""}>
    </td>
    <td class="${!isOtherBilling ? "disabled-field" : ""}">
      <input type="email" class="billing-contact-email-input" value="${feeEarner.billingContactEmail}" 
             placeholder="Enter email" ${!isOtherBilling ? "disabled" : ""}>
    </td>
    <td>
      <button type="button" class="remove-fee-earner" onclick="removeFeeEarnerRow(this)">Remove</button>
    </td>
  `;

  tbody.appendChild(row);

  // Add event listener for billing contact change
  const billingContactSelect = row.querySelector(".billing-contact-select");
  billingContactSelect?.addEventListener("change", handleBillingContactChange);

  // Add event listeners for duplicate detection and default management
  const nameInput = row.querySelector(".name-input");
  const useAsDefaultCheckbox = row.querySelector(".use-as-default-checkbox");

  nameInput?.addEventListener("input", () => detectDuplicateNames());
  useAsDefaultCheckbox?.addEventListener("change", (event) =>
    handleDefaultCheckboxChange(event, row)
  );

  // Ensure scroll container can scroll to the beginning
  resetTableScroll();
}

function removeFeeEarnerRow(button: HTMLButtonElement) {
  const row = button.closest("tr");
  if (row) {
    row.remove();
    // Re-run duplicate detection after removal
    setTimeout(() => detectDuplicateNames(), 50);
  }
}

function handleBillingContactChange(event: Event) {
  const select = event.target as HTMLSelectElement;
  const row = select.closest("tr");
  if (!row) return;

  const isOther = select.value === "Other";
  const billingNameCell = row.children[6] as HTMLTableCellElement;
  const billingEmailCell = row.children[7] as HTMLTableCellElement;
  const billingNameInput = billingNameCell.querySelector("input") as HTMLInputElement;
  const billingEmailInput = billingEmailCell.querySelector("input") as HTMLInputElement;

  if (isOther) {
    // Enable billing contact fields
    billingNameCell.classList.remove("disabled-field");
    billingEmailCell.classList.remove("disabled-field");
    billingNameInput.disabled = false;
    billingEmailInput.disabled = false;
  } else {
    // Disable and clear billing contact fields
    billingNameCell.classList.add("disabled-field");
    billingEmailCell.classList.add("disabled-field");
    billingNameInput.disabled = true;
    billingEmailInput.disabled = true;
    billingNameInput.value = "";
    billingEmailInput.value = "";
  }
}

function detectDuplicateNames() {
  const tbody = document.getElementById("fee-earners-tbody");
  if (!tbody) return;

  const rows = Array.from(tbody.querySelectorAll("tr"));
  const nameGroups = new Map<string, HTMLTableRowElement[]>();
  let hasDuplicates = false;

  // Group rows by first name (case insensitive)
  rows.forEach((row) => {
    const nameInput = row.querySelector(".name-input") as HTMLInputElement;
    if (nameInput && nameInput.value.trim()) {
      const firstName = nameInput.value.trim().split(" ")[0].toLowerCase();
      if (!nameGroups.has(firstName)) {
        nameGroups.set(firstName, []);
      }
      nameGroups.get(firstName)!.push(row);
    }
  });

  // Apply duplicate styling and manage default checkboxes
  rows.forEach((row) => {
    const nameInput = row.querySelector(".name-input") as HTMLInputElement;
    const useAsDefaultCheckbox = row.querySelector(".use-as-default-checkbox") as HTMLInputElement;

    if (nameInput && nameInput.value.trim()) {
      const firstName = nameInput.value.trim().split(" ")[0].toLowerCase();
      const duplicateRows = nameGroups.get(firstName) || [];

      if (duplicateRows.length > 1) {
        // This is a duplicate name - highlight the row
        row.classList.add("duplicate-name");
        useAsDefaultCheckbox.style.display = "block";
        hasDuplicates = true;

        // Ensure only one is marked as default
        const checkedDefaults = duplicateRows.filter(
          (r) => (r.querySelector(".use-as-default-checkbox") as HTMLInputElement).checked
        );

        if (checkedDefaults.length === 0) {
          // Auto-select the first one as default
          (duplicateRows[0].querySelector(".use-as-default-checkbox") as HTMLInputElement).checked =
            true;
        }
      } else {
        // Not a duplicate - remove highlighting and hide checkbox
        row.classList.remove("duplicate-name");
        useAsDefaultCheckbox.style.display = "none";
        useAsDefaultCheckbox.checked = false;
      }
    }
  });

  // Show/hide duplicate names info panel
  const duplicateNamesInfo = document.getElementById("duplicate-names-info");
  if (duplicateNamesInfo) {
    duplicateNamesInfo.style.display = hasDuplicates ? "block" : "none";
  }
}

function resetTableScroll() {
  // Reset horizontal scroll position to beginning
  const container = document.querySelector(".fee-earners-container") as HTMLElement;
  if (container) {
    // Force scroll to leftmost position
    container.scrollLeft = 0;
    container.scrollTo({ left: 0, behavior: "auto" });

    // Also ensure proper CSS reset
    container.style.scrollBehavior = "auto";
    setTimeout(() => {
      container.scrollLeft = 0;
      container.style.scrollBehavior = "smooth";
    }, 50);
  }
}

function handleDefaultCheckboxChange(event: Event, currentRow: HTMLTableRowElement) {
  const checkbox = event.target as HTMLInputElement;
  const nameInput = currentRow.querySelector(".name-input") as HTMLInputElement;

  if (checkbox.checked && nameInput.value.trim()) {
    const firstName = nameInput.value.trim().split(" ")[0].toLowerCase();
    const tbody = document.getElementById("fee-earners-tbody");
    if (!tbody) return;

    // Uncheck all other default checkboxes for the same first name
    const rows = Array.from(tbody.querySelectorAll("tr"));
    rows.forEach((row) => {
      if (row !== currentRow) {
        const rowNameInput = row.querySelector(".name-input") as HTMLInputElement;
        if (rowNameInput && rowNameInput.value.trim()) {
          const rowFirstName = rowNameInput.value.trim().split(" ")[0].toLowerCase();
          if (rowFirstName === firstName) {
            const rowCheckbox = row.querySelector(".use-as-default-checkbox") as HTMLInputElement;
            rowCheckbox.checked = false;
          }
        }
      }
    });
  }
}

// Stage 3: Core Name Matching Logic
function applyNameStandardisationRule(
  worksheetData: any[],
  feeEarners: FeeEarner[],
  ruleConfig: NameStandardisationRule
): any[] {
  if (!ruleConfig.enabled || feeEarners.length === 0) {
    return worksheetData;
  }

  const processedData = [...worksheetData];

  processedData.forEach((row, rowIndex) => {
    // Find the source narrative column (Original Narrative or Narrative)
    const sourceNarrativeKey = findSourceNarrativeColumn(row);
    if (!sourceNarrativeKey) return;

    const narrativeText = row[sourceNarrativeKey];
    if (!narrativeText || typeof narrativeText !== "string") return;

    // Process the narrative text for name standardisation
    const processedText = processNarrativeForNames(
      narrativeText,
      feeEarners,
      ruleConfig,
      row.Date || row.date || null // Try to get date for matching
    );

    // Only update the amended narrative column if changes were made
    const amendedColumnKey = getOrCreateAmendedNarrativeColumn(row);
    if (amendedColumnKey && processedText !== narrativeText) {
      row[amendedColumnKey] = processedText;
    }
  });

  return processedData;
}

function findSourceNarrativeColumn(row: any): string | null {
  const keys = Object.keys(row);

  // First, look for "Original Narrative"
  const originalNarrative = keys.find(
    (key) => key.toLowerCase().includes("original") && key.toLowerCase().includes("narrative")
  );
  if (originalNarrative) return originalNarrative;

  // Then look for just "Narrative" (but not "Amended Narrative")
  const narrative = keys.find(
    (key) =>
      key.toLowerCase().includes("narrative") &&
      !key.toLowerCase().includes("amended") &&
      !key.toLowerCase().includes("original")
  );
  if (narrative) return narrative;

  // Finally, check for "Description"
  const description = keys.find((key) => key.toLowerCase().includes("description"));
  return description || null;
}

function getOrCreateAmendedNarrativeColumn(row: any): string | null {
  const keys = Object.keys(row);

  // Look for existing "Amended Narrative" column
  const amendedColumn = keys.find(
    (key) => key.toLowerCase().includes("amended") && key.toLowerCase().includes("narrative")
  );

  if (amendedColumn) {
    return amendedColumn;
  }

  // If not found, we'll need to create it - return the expected name
  return "Amended Narrative";
}

function processNarrativeForNames(
  narrativeText: string,
  feeEarners: FeeEarner[],
  ruleConfig: NameStandardisationRule,
  rowDate?: string | Date | null
): string {
  if (!narrativeText) return narrativeText;

  let processedText = narrativeText;
  const excludedNames = ruleConfig.excludedNames.map((name) => name.toLowerCase().trim());

  // Create a map of first names to fee earners for quick lookup
  const nameMap = createFeeEarnerNameMap(feeEarners, ruleConfig.allowPartialMatches);

  // Process each word in the narrative
  let hasReplacements = false;

  // Split text but keep track of original spacing and punctuation
  const wordPattern = /\b(\w+)\b/g;
  let match;

  while ((match = wordPattern.exec(narrativeText)) !== null) {
    const word = match[1];
    const cleanWord = word.toLowerCase();

    // Skip if word is excluded
    if (excludedNames.includes(cleanWord)) continue;

    // Skip if word is too short to be a meaningful name
    if (cleanWord.length < 2) continue;

    // Check if this word matches any fee earner names
    const minLength = ruleConfig.minPartialMatchLength || 3; // Default to 3 if not set
    let matchingFeeEarners = findMatchingFeeEarners(
      cleanWord,
      nameMap,
      ruleConfig.allowPartialMatches,
      minLength
    );

    // If no direct match and nickname database is enabled, check nicknames
    if (matchingFeeEarners.length === 0 && ruleConfig.useNicknameDatabase !== false) {
      const expandedName = findNicknameMatch(cleanWord, ruleConfig.customNicknames || {});
      if (expandedName) {
        // Try to find fee earners with this expanded name
        matchingFeeEarners = findMatchingFeeEarners(expandedName, nameMap, false, minLength);
      }
    }

    if (matchingFeeEarners.length > 0) {
      // Determine which fee earner to use
      const selectedFeeEarner = selectBestFeeEarnerMatch(
        matchingFeeEarners,
        rowDate,
        ruleConfig.useDateMatching
      );

      if (selectedFeeEarner && selectedFeeEarner.name !== word) {
        // Check if this word is already part of a full name
        if (isAlreadyFullName(word, selectedFeeEarner.name, narrativeText, match.index!)) {
          continue; // Skip replacement if it's already a full name
        }

        // Replace the first name with the full name
        if (ruleConfig.replaceOnlyFirstOccurrence && hasReplacements) {
          // Skip if we've already done a replacement and only first occurrence is enabled
          continue;
        }

        // Create a regex that preserves the original case and word boundaries
        const regex = new RegExp(`\\b${escapeRegExp(word)}\\b`, hasReplacements ? "g" : "");
        processedText = processedText.replace(regex, selectedFeeEarner.name);
        hasReplacements = true;

        // If only replacing first occurrence, we can stop after the first replacement
        if (ruleConfig.replaceOnlyFirstOccurrence) {
          break;
        }
      }
    }
  }

  return processedText;
}

function isAlreadyFullName(
  foundWord: string,
  fullName: string,
  narrativeText: string,
  wordIndex: number
): boolean {
  // Look for any word that follows the found word
  // Get text starting from after the found word
  const afterWord = narrativeText.slice(wordIndex + foundWord.length);

  // Use regex to find the next word
  const nextWordMatch = afterWord.match(/^\s+(\w+)/);

  if (nextWordMatch) {
    const nextWord = nextWordMatch[1];

    // Check if the next word looks like a surname (capitalized and not obviously not a name)
    if (isLikelySurname(nextWord)) {
      return true; // Already appears to be part of a full name, don't replace
    }

    // Additional check: if the fee earner's name matches exactly what we found
    const nameParts = fullName.split(/\s+/);
    if (nameParts.length >= 2) {
      const firstName = nameParts[0].toLowerCase();
      const lastName = nameParts[nameParts.length - 1].toLowerCase();

      // Check if found word + next word matches the fee earner's name exactly
      if (foundWord.toLowerCase() === firstName && nextWord.toLowerCase() === lastName) {
        return true; // This is exactly the fee earner's full name, don't replace
      }
    }
  }

  return false; // Not a full name, safe to replace
}

function isLikelySurname(word: string): boolean {
  // Check if a word is likely to be a surname based on common patterns

  // Must be capitalized (proper noun)
  if (word[0] !== word[0].toUpperCase()) {
    return false;
  }

  // Must be at least 2 characters
  if (word.length < 2) {
    return false;
  }

  // Exclude common words that might be capitalized but aren't surnames
  const excludedWords = [
    "THE",
    "AND",
    "OR",
    "BUT",
    "FOR",
    "NOR",
    "SO",
    "YET",
    "IN",
    "ON",
    "AT",
    "TO",
    "FROM",
    "BY",
    "WITH",
    "ABOUT",
    "MONDAY",
    "TUESDAY",
    "WEDNESDAY",
    "THURSDAY",
    "FRIDAY",
    "SATURDAY",
    "SUNDAY",
    "JANUARY",
    "FEBRUARY",
    "MARCH",
    "APRIL",
    "MAY",
    "JUNE",
    "JULY",
    "AUGUST",
    "SEPTEMBER",
    "OCTOBER",
    "NOVEMBER",
    "DECEMBER",
    "THIS",
    "THAT",
    "THESE",
    "THOSE",
    "HIS",
    "HER",
    "THEIR",
    "OUR",
    "YOUR",
    "WORKED",
    "ATTENDED",
    "REVIEWED",
    "PREPARED",
    "DRAFTED",
    "MEETING",
    "CALL",
  ];

  if (excludedWords.includes(word.toUpperCase())) {
    return false;
  }

  // If it passes these tests, it's likely a surname
  return true;
}

function findNicknameMatch(
  searchWord: string,
  customNicknames: Record<string, string>
): string | null {
  const lowerSearchWord = searchWord.toLowerCase();

  // Check custom nicknames first (they override defaults)
  if (customNicknames[lowerSearchWord]) {
    return customNicknames[lowerSearchWord];
  }

  // Check built-in nickname database
  if (DEFAULT_NICKNAMES[lowerSearchWord]) {
    return DEFAULT_NICKNAMES[lowerSearchWord];
  }

  return null;
}

function createFeeEarnerNameMap(
  feeEarners: FeeEarner[],
  allowPartialMatches: boolean
): Map<string, FeeEarner[]> {
  const nameMap = new Map<string, FeeEarner[]>();

  feeEarners.forEach((feeEarner) => {
    if (!feeEarner.name) return;

    const names = feeEarner.name.split(/\s+/);
    const firstName = names[0].toLowerCase();

    // Add the first name to the map
    if (!nameMap.has(firstName)) {
      nameMap.set(firstName, []);
    }
    nameMap.get(firstName)!.push(feeEarner);

    // If partial matches are allowed, also add name variations
    if (allowPartialMatches && feeEarner.nameVariations) {
      feeEarner.nameVariations.forEach((variation) => {
        const variationKey = variation.toLowerCase().trim();
        if (!nameMap.has(variationKey)) {
          nameMap.set(variationKey, []);
        }
        nameMap.get(variationKey)!.push(feeEarner);
      });
    }
  });

  return nameMap;
}

function findMatchingFeeEarners(
  searchName: string,
  nameMap: Map<string, FeeEarner[]>,
  allowPartialMatches: boolean,
  minPartialMatchLength: number = 3
): FeeEarner[] {
  // Direct match
  if (nameMap.has(searchName)) {
    return nameMap.get(searchName)!;
  }

  // Partial matching if enabled
  if (allowPartialMatches) {
    const matches: FeeEarner[] = [];
    const uniqueFeeEarners = new Set<FeeEarner>();

    nameMap.forEach((feeEarners, mappedName) => {
      // For partial matching, we want to be more careful:
      // 1. Both names should meet minimum length requirement
      // 2. Only match prefixes, not arbitrary substrings

      if (
        searchName.length >= minPartialMatchLength &&
        mappedName.length >= minPartialMatchLength
      ) {
        // Check if search name is a prefix of mapped name (e.g., "John" matches "Johnny")
        if (mappedName.startsWith(searchName)) {
          feeEarners.forEach((fe) => uniqueFeeEarners.add(fe));
        }
        // Check if mapped name is a prefix of search name (e.g., "Johnny" typed, "John" in system)
        else if (searchName.startsWith(mappedName)) {
          feeEarners.forEach((fe) => uniqueFeeEarners.add(fe));
        }
      }
    });

    return Array.from(uniqueFeeEarners);
  }

  return [];
}

function selectBestFeeEarnerMatch(
  matchingFeeEarners: FeeEarner[],
  rowDate: string | Date | null,
  useDateMatching: boolean
): FeeEarner | null {
  if (matchingFeeEarners.length === 0) return null;
  if (matchingFeeEarners.length === 1) return matchingFeeEarners[0];

  // If date matching is enabled and we have a date, try to find the best match
  if (useDateMatching && rowDate) {
    const parsedRowDate = parseDate(rowDate);
    if (parsedRowDate) {
      // Try to find a fee earner with a matching date within ±5 days
      const dateMatchedFeeEarner = findFeeEarnerByDateRange(matchingFeeEarners, parsedRowDate, 5);
      if (dateMatchedFeeEarner) {
        return dateMatchedFeeEarner;
      }
    }
  }

  // Find the fee earner marked as default for this name group
  const defaultFeeEarner = matchingFeeEarners.find((fe) => fe.isDefaultForName);
  if (defaultFeeEarner) return defaultFeeEarner;

  // Fall back to the first one if no default is set
  return matchingFeeEarners[0];
}

function findFeeEarnerByDateRange(
  feeEarners: FeeEarner[],
  targetDate: Date,
  daysTolerance: number
): FeeEarner | null {
  // In a full implementation, this would check against fee earner assignment dates
  // stored in the fee earner records or a separate tracking system

  // For now, we'll implement a placeholder that could be extended
  // to check against actual date records when that functionality is added

  const targetTime = targetDate.getTime();
  const toleranceMs = daysTolerance * 24 * 60 * 60 * 1000; // Convert days to milliseconds

  for (const feeEarner of feeEarners) {
    // Placeholder: In a real implementation, this would check against
    // stored dates for when each fee earner was assigned to work
    // For now, we'll return null to fall back to default logic
    // Future enhancement: Check feeEarner.assignmentDates or similar
    // if (feeEarner.assignmentDates) {
    //   for (const assignmentDate of feeEarner.assignmentDates) {
    //     const assignmentTime = parseDate(assignmentDate)?.getTime();
    //     if (assignmentTime && Math.abs(targetTime - assignmentTime) <= toleranceMs) {
    //       return feeEarner;
    //     }
    //   }
    // }
  }

  return null; // No date-based match found, fall back to default logic
}

// Enhanced date parsing with multiple format support
function parseDate(dateInput: string | Date): Date | null {
  if (!dateInput) return null;

  if (dateInput instanceof Date) return dateInput;

  // Try multiple date formats
  const dateString = dateInput.toString().trim();

  // Common formats to try
  const formats = [
    // ISO format
    /^\d{4}-\d{2}-\d{2}$/,
    // US format MM/DD/YYYY
    /^\d{1,2}\/\d{1,2}\/\d{4}$/,
    // UK format DD/MM/YYYY
    /^\d{1,2}\/\d{1,2}\/\d{4}$/,
    // Short format with dashes
    /^\d{1,2}-\d{1,2}-\d{4}$/,
    // Excel date number (if it's a number)
    /^\d+(\.\d+)?$/,
  ];

  // Try standard JavaScript Date parsing first
  let date = new Date(dateString);
  if (!isNaN(date.getTime())) {
    return date;
  }

  // If that fails, try Excel serial date number conversion
  const numericValue = parseFloat(dateString);
  if (!isNaN(numericValue) && numericValue > 40000 && numericValue < 50000) {
    // Excel serial date (days since January 1, 1900)
    // Note: Excel incorrectly treats 1900 as a leap year, so we need to adjust
    const excelEpoch = new Date(1899, 11, 30); // December 30, 1899
    date = new Date(excelEpoch.getTime() + numericValue * 24 * 60 * 60 * 1000);
    if (!isNaN(date.getTime())) {
      return date;
    }
  }

  return null;
}

function escapeRegExp(string: string): string {
  return string.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Apply name standardisation rules to the current worksheet
async function applyNameStandardisationToWorksheet() {
  try {
    // Get current matter and its rules
    const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;
    if (!selectedMatter) {
      showMessage(
        "Please select a matter profile before applying name standardisation rules.",
        "error"
      );
      return;
    }

    const profiles = getMatterProfiles();
    const currentProfile = profiles.find((p) => p.name === selectedMatter);
    if (!currentProfile || !currentProfile.rules) {
      showMessage("No rules found for the selected matter profile.", "error");
      return;
    }

    const nameRule = currentProfile.rules.nameStandardisation;
    if (!nameRule.enabled) {
      showMessage("Name standardisation rule is not enabled for this matter.", "info");
      return;
    }

    const feeEarners = currentProfile.feeEarners || [];
    if (feeEarners.length === 0) {
      showMessage("No fee earners found for this matter. Please add fee earners first.", "error");
      return;
    }

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = worksheet.getUsedRange();

      if (!usedRange) {
        showMessage("No data found in the worksheet to process.", "error");
        return;
      }

      usedRange.load(["values", "formulas", "rowCount", "columnCount"]);
      await context.sync();

      // Convert Excel data to a workable format
      const headers = usedRange.values[0] as string[];
      const worksheetData: any[] = [];

      for (let i = 1; i < usedRange.values.length; i++) {
        const row = usedRange.values[i];
        const rowData: any = {};

        headers.forEach((header, colIndex) => {
          rowData[header] = row[colIndex];
        });

        worksheetData.push(rowData);
      }

      // Apply name standardisation rules
      const processedData = applyNameStandardisationRule(worksheetData, feeEarners, nameRule);

      // Update the worksheet with processed data
      let updatedCount = 0;

      // First check if "Amended Narrative" column exists
      let amendedNarrativeCol = headers.findIndex(
        (h) => h.toLowerCase().includes("amended") && h.toLowerCase().includes("narrative")
      );

      // If column doesn't exist, we need to check if we need to create it
      const needsAmendedColumn =
        amendedNarrativeCol < 0 &&
        processedData.some((row) => row["Amended Narrative"] !== undefined);

      if (needsAmendedColumn) {
        // We'll need to add the column after the source narrative column
        const sourceCol = headers.findIndex(
          (h) =>
            (h.toLowerCase().includes("original") && h.toLowerCase().includes("narrative")) ||
            (h.toLowerCase().includes("narrative") && !h.toLowerCase().includes("amended"))
        );

        if (sourceCol >= 0) {
          // Insert new column after the source narrative column
          const insertCol = sourceCol + 1;
          const newColumn = worksheet.getCell(0, insertCol).getEntireColumn();
          newColumn.insert(Excel.InsertShiftDirection.right);

          // Set the header for the new column
          const headerCell = worksheet.getCell(0, insertCol);
          headerCell.values = [["Amended Narrative"]];

          // Update our tracking
          amendedNarrativeCol = insertCol;
          headers.splice(insertCol, 0, "Amended Narrative");

          await context.sync();
        }
      }

      // Now update the data
      for (let i = 0; i < processedData.length; i++) {
        const processedRow = processedData[i];
        const amendedValue = processedRow["Amended Narrative"];

        if (amendedValue !== undefined && amendedNarrativeCol >= 0) {
          // Update the cell in Excel (i+1 because row 0 is headers)
          const cell = worksheet.getCell(i + 1, amendedNarrativeCol);
          cell.values = [[amendedValue]];
          updatedCount++;
        }
      }

      await context.sync();

      if (updatedCount > 0) {
        showMessage(
          `Name standardisation applied successfully. Updated ${updatedCount} rows.`,
          "success"
        );
      } else {
        showMessage("Name standardisation completed, but no changes were needed.", "info");
      }
    });
  } catch (error) {
    console.error("Error applying name standardisation:", error);
    showMessage("An error occurred while applying name standardisation: " + error.message, "error");
  }
}

async function updateFeeEarnersFromSpreadsheet() {
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

      const headerRow = usedRange.values[0];
      const nameColumnIndex = findNameColumn(headerRow);
      const roleColumnIndex = findRoleColumn(headerRow);
      const rateColumnIndex = findRateColumn(headerRow);

      if (nameColumnIndex === -1) {
        showMessage("No Name column found in the spreadsheet.", "error");
        return;
      }

      // Extract unique fee earner combinations
      const uniqueFeeEarners = new Map<string, FeeEarner>();

      // Start from row 1 (skip header row)
      for (let row = 1; row < usedRange.rowCount; row++) {
        const nameValue = usedRange.values[row][nameColumnIndex];
        const roleValue = roleColumnIndex !== -1 ? usedRange.values[row][roleColumnIndex] : "";
        const rateValue = rateColumnIndex !== -1 ? usedRange.values[row][rateColumnIndex] : "";

        if (nameValue && nameValue.toString().trim()) {
          const name = nameValue.toString().trim();
          const role = roleValue ? roleValue.toString().trim() : "";
          const rate = rateValue ? parseFloat(rateValue.toString()) || 0 : 0;

          // Create a unique key for this combination
          const key = `${name}-${role}-${rate}`;

          if (!uniqueFeeEarners.has(key)) {
            uniqueFeeEarners.set(key, {
              name: name,
              role: role,
              rate: rate,
              email: "", // Will be manually filled
              billingContact: "Fee Earner", // Default
              billingContactName: "",
              billingContactEmail: "",
            });
          }
        }
      }

      // Update the fee earners table with found data
      const feeEarners = Array.from(uniqueFeeEarners.values());

      if (feeEarners.length > 0) {
        loadFeeEarnersTable(feeEarners);

        const foundColumns = [];
        if (nameColumnIndex !== -1) foundColumns.push("Name");
        if (roleColumnIndex !== -1) foundColumns.push("Role");
        if (rateColumnIndex !== -1) foundColumns.push("Rate");

        showMessage(
          `Found ${feeEarners.length} unique fee earner${feeEarners.length > 1 ? "s" : ""} from ${foundColumns.join(", ")} column${foundColumns.length > 1 ? "s" : ""}. Please fill in missing information manually.`,
          "success"
        );
      } else {
        showMessage("No fee earners found in the spreadsheet data.", "warning");
      }
    });
  } catch (error) {
    console.error(error);
    showMessage("An error occurred while scanning the spreadsheet: " + error.message, "error");
  }
}

function saveParticipants() {
  const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;

  if (!selectedMatter) {
    showMessage("Please select a matter from the dropdown to save participants to.", "error");
    return;
  }

  // Get current fee earners from the table
  const currentFeeEarners = getCurrentFeeEarners();

  // Filter out completely empty rows
  const validFeeEarners = currentFeeEarners.filter(
    (feeEarner) =>
      feeEarner.name.trim() !== "" ||
      feeEarner.role.trim() !== "" ||
      feeEarner.rate > 0 ||
      feeEarner.email.trim() !== ""
  );

  // Get existing matter profiles
  const profiles = getMatterProfiles();
  const existingIndex = profiles.findIndex((p) => p.name === selectedMatter);

  if (existingIndex >= 0) {
    // Update the existing profile with new fee earners data
    profiles[existingIndex].feeEarners = validFeeEarners;
    saveMatterProfiles(profiles);

    showMessage(
      `Successfully saved ${validFeeEarners.length} fee earner${validFeeEarners.length !== 1 ? "s" : ""} to matter profile "${selectedMatter}".`,
      "success"
    );
  } else {
    showMessage("Selected matter profile not found. Please create a new profile first.", "error");
  }
}

// Nickname Database Management
function addNicknameEntry() {
  const nicknameList = document.getElementById("nickname-list");
  if (!nicknameList) return;

  const entry = createNicknameEntry("", "", false);
  nicknameList.appendChild(entry);

  // Focus the first input in the new entry
  const firstInput = entry.querySelector("input");
  if (firstInput) {
    firstInput.focus();
  }
}

function createNicknameEntry(
  nickname: string,
  fullName: string,
  isBuiltIn: boolean = false
): HTMLElement {
  const entry = document.createElement("div");
  entry.className = `nickname-entry ${isBuiltIn ? "built-in" : ""}`;

  entry.innerHTML = `
    <input type="text" class="nickname-input" value="${nickname}" placeholder="Nickname" ${isBuiltIn ? "readonly" : ""}>
    <span class="nickname-arrow">→</span>
    <input type="text" class="fullname-input" value="${fullName}" placeholder="Full Name" ${isBuiltIn ? "readonly" : ""}>
    <button type="button" class="nickname-remove" onclick="removeNicknameEntry(this)">
      ${isBuiltIn ? "Hide" : "Remove"}
    </button>
  `;

  return entry;
}

function removeNicknameEntry(button: HTMLButtonElement) {
  const entry = button.closest(".nickname-entry");
  if (entry) {
    entry.remove();
  }
}

function resetNicknamesToDefault() {
  loadNicknameDatabase(DEFAULT_NICKNAMES);
}

function loadNicknameDatabase(nicknames: Record<string, string>) {
  const nicknameList = document.getElementById("nickname-list");
  if (!nicknameList) return;

  nicknameList.innerHTML = "";

  // Add built-in nicknames (read-only) - sorted alphabetically by nickname
  const sortedBuiltInNicknames = Object.entries(DEFAULT_NICKNAMES).sort(
    ([nicknameA], [nicknameB]) => nicknameA.toLowerCase().localeCompare(nicknameB.toLowerCase())
  );

  sortedBuiltInNicknames.forEach(([nickname, fullName]) => {
    const entry = createNicknameEntry(nickname, fullName, true);
    nicknameList.appendChild(entry);
  });

  // Add custom nicknames (editable) - sorted alphabetically by nickname
  const sortedCustomNicknames = Object.entries(nicknames)
    .filter(([nickname]) => !DEFAULT_NICKNAMES[nickname])
    .sort(([nicknameA], [nicknameB]) =>
      nicknameA.toLowerCase().localeCompare(nicknameB.toLowerCase())
    );

  sortedCustomNicknames.forEach(([nickname, fullName]) => {
    const entry = createNicknameEntry(nickname, fullName, false);
    nicknameList.appendChild(entry);
  });
}

function getCurrentNicknames(): Record<string, string> {
  const nicknames: Record<string, string> = {};
  const nicknameList = document.getElementById("nickname-list");

  if (nicknameList) {
    const entries = nicknameList.querySelectorAll(".nickname-entry");
    entries.forEach((entry) => {
      const nicknameInput = entry.querySelector(".nickname-input") as HTMLInputElement;
      const fullNameInput = entry.querySelector(".fullname-input") as HTMLInputElement;

      if (
        nicknameInput &&
        fullNameInput &&
        nicknameInput.value.trim() &&
        fullNameInput.value.trim()
      ) {
        nicknames[nicknameInput.value.trim().toLowerCase()] = fullNameInput.value
          .trim()
          .toLowerCase();
      }
    });
  }

  return nicknames;
}

// Undo and Auto-Apply Functionality
async function undoNameStandardisation() {
  if (!lastUndoSnapshot) {
    showMessage("No changes to undo.", "info");
    return;
  }

  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();

      // Restore all changed values
      for (const change of lastUndoSnapshot.changes) {
        const cell = worksheet.getCell(change.row, change.column);
        cell.values = [[change.oldValue]];
      }

      await context.sync();

      showMessage(
        `Undid ${lastUndoSnapshot.changes.length} changes from name standardisation.`,
        "success"
      );

      // Clear the undo snapshot and disable undo button
      lastUndoSnapshot = null;
      updateUndoButtonState();
    });
  } catch (error) {
    console.error("Error undoing changes:", error);
    showMessage("An error occurred while undoing changes: " + error.message, "error");
  }
}

function updateUndoButtonState() {
  const undoButton = document.getElementById("undo-name-rules") as HTMLElement;
  if (undoButton) {
    if (lastUndoSnapshot) {
      undoButton.removeAttribute("disabled");
      undoButton.style.opacity = "1";
      undoButton.style.cursor = "pointer";
    } else {
      undoButton.setAttribute("disabled", "true");
      undoButton.style.opacity = "0.5";
      undoButton.style.cursor = "not-allowed";
    }
  }
}

// Apply all enabled rules
async function applyAllRules() {
  try {
    // Get current matter and its rules
    const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;
    if (!selectedMatter) {
      showMessage("Please select a matter profile before applying rules.", "error");
      return;
    }

    const profiles = getMatterProfiles();
    const currentProfile = profiles.find((p) => p.name === selectedMatter);
    if (!currentProfile || !currentProfile.rules) {
      showMessage("No rules found for the selected matter profile.", "error");
      return;
    }

    // DEBUG: Show what rules are enabled
    const nameStandardisationEnabled = currentProfile.rules.nameStandardisation?.enabled || false;
    const missingTimeEnabled = currentProfile.rules.missingTimeEntries?.enabled || false;

    showMessage(
      `DEBUG: Starting rules for ${selectedMatter}. Name Standardisation: ${nameStandardisationEnabled ? "ENABLED" : "disabled"}, Missing Time: ${missingTimeEnabled ? "ENABLED" : "disabled"}`,
      "info"
    );

    let appliedRules = [];
    let totalUpdatedRows = 0;

    // Apply Name Standardisation Rule if enabled
    if (currentProfile.rules.nameStandardisation?.enabled) {
      showMessage("Applying Name Standardisation rule...", "info");

      const result = await applyNameStandardisationRuleWithResult();
      if (result.success) {
        appliedRules.push("Name Standardisation");
        totalUpdatedRows += result.updatedRows;
      } else if (result.error) {
        showMessage(`Name Standardisation failed: ${result.error}`, "error");
        return;
      }
    }

    // Apply Missing Time Entries Rule if enabled
    if (currentProfile.rules.missingTimeEntries?.enabled) {
      showMessage("Applying Missing Time Entries rule...", "info");

      const result = await applyMissingTimeEntriesRuleWithResult();

      if (result.success) {
        appliedRules.push("Missing Time Entries");
        totalUpdatedRows += result.updatedRows;
        showMessage(`Missing Time Entries completed: ${result.updatedRows} rows updated`, "info");
      } else if (result.error) {
        showMessage(`Missing Time Entries failed: ${result.error}`, "error");
        return;
      }
    } else {
      showMessage("Missing Time Entries rule is disabled - skipping", "info");
    }

    // Show final result (temporarily disabling formatting reapplication to debug Notes issue)
    if (appliedRules.length > 0) {
      console.log("DEBUGGING: Skipping formatSpreadsheet() to preserve Notes");

      const rulesText = appliedRules.join(", ");
      showMessage(
        `Successfully applied ${rulesText}. Updated ${totalUpdatedRows} rows total. (Format reapplication disabled for debugging). Undo is available.`,
        "success"
      );
    } else {
      showMessage("No rules were enabled for this matter profile.", "info");
    }
  } catch (error) {
    console.error("Error applying rules:", error);
    showMessage("An error occurred while applying rules: " + error.message, "error");
  }
}

// Helper function to apply Name Standardisation and return result
async function applyNameStandardisationRuleWithResult(): Promise<{
  success: boolean;
  updatedRows: number;
  error?: string;
}> {
  try {
    const updatedCount = await applyNameStandardisationToWorksheetWithUndo();
    return { success: true, updatedRows: updatedCount };
  } catch (error) {
    return { success: false, updatedRows: 0, error: error.message };
  }
}

// Missing Time Entries rule implementation
async function applyMissingTimeEntriesRuleWithResult(): Promise<{
  success: boolean;
  updatedRows: number;
  error?: string;
}> {
  try {
    // Get current matter and its rules
    const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;
    if (!selectedMatter) {
      return { success: false, updatedRows: 0, error: "No matter selected" };
    }

    const profiles = getMatterProfiles();
    const currentProfile = profiles.find((p) => p.name === selectedMatter);
    if (!currentProfile || !currentProfile.rules || !currentProfile.rules.missingTimeEntries) {
      return { success: false, updatedRows: 0, error: "Missing time entries rule not configured" };
    }

    const missingTimeRule = currentProfile.rules.missingTimeEntries;
    if (!missingTimeRule.enabled) {
      return { success: false, updatedRows: 0, error: "Missing time entries rule not enabled" };
    }

    return await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = worksheet.getUsedRange();

      if (!usedRange) {
        return { success: false, updatedRows: 0, error: "No data found in worksheet" };
      }

      usedRange.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      // Parse data into structured format
      const headers = usedRange.values[0] as string[];
      const entries = [];

      for (let i = 1; i < usedRange.values.length; i++) {
        const row = usedRange.values[i];
        const entry: any = { rowIndex: i };

        headers.forEach((header, colIndex) => {
          entry[header.toLowerCase().replace(/\s+/g, "_")] = row[colIndex];
        });

        entries.push(entry);
      }

      // Find required columns with more flexible matching
      const feeEarnerCol = headers.findIndex((h) => {
        const headerLower = h.toLowerCase();
        return (
          (headerLower.includes("fee") && headerLower.includes("earner")) ||
          headerLower.includes("name") ||
          headerLower.includes("person") ||
          headerLower.includes("who") ||
          headerLower.includes("user")
        );
      });

      const dateCol = headers.findIndex((h) => {
        const headerLower = h.toLowerCase();
        return (
          headerLower.includes("date") ||
          headerLower.includes("day") ||
          headerLower.includes("when")
        );
      });

      const narrativeCol = headers.findIndex((h) => {
        const headerLower = h.toLowerCase();
        return (
          headerLower.includes("narrative") ||
          headerLower.includes("description") ||
          headerLower.includes("note") ||
          headerLower.includes("detail") ||
          headerLower.includes("work") ||
          headerLower.includes("activity")
        );
      });

      if (feeEarnerCol === -1 || dateCol === -1 || narrativeCol === -1) {
        return {
          success: false,
          updatedRows: 0,
          error: `Required columns not found. Found: FeeEarner(${feeEarnerCol}), Date(${dateCol}), Narrative(${narrativeCol}). Headers: ${headers.join(", ")}`,
        };
      }

      console.log(
        `Missing Time Rule: Using columns - FeeEarner: ${headers[feeEarnerCol]}, Date: ${headers[dateCol]}, Narrative: ${headers[narrativeCol]}`
      );

      // Get fee earners list for name matching
      const feeEarners = currentProfile.participants?.feeEarners || [];
      console.log(`Fee earners available: ${feeEarners.map((fe) => fe.name).join(", ")}`);
      console.log(`Meeting keywords: ${missingTimeRule.meetingKeywords.join(", ")}`);
      console.log(`Date tolerance: ${missingTimeRule.dateTolerance} days`);
      console.log(`Total entries to process: ${entries.length}`);

      // Show initial debugging info
      showMessage(
        `DEBUG: Starting with ${entries.length} entries, ${feeEarners.length} fee earners: ${feeEarners.map((fe) => fe.name).join(", ")}`,
        "info"
      );
      showMessage(
        `DEBUG: Meeting keywords: "${missingTimeRule.meetingKeywords.join('", "')}"`,
        "info"
      );

      const missingEntries = [];

      // Process each entry looking for meeting keywords
      let processedCount = 0;
      let meetingEntriesFound = 0;

      for (const entry of entries) {
        processedCount++;
        const narrative = (entry[headers[narrativeCol].toLowerCase().replace(/\s+/g, "_")] || "")
          .toString()
          .toLowerCase();
        const entryFeeEarner = (
          entry[headers[feeEarnerCol].toLowerCase().replace(/\s+/g, "_")] || ""
        ).toString();
        const entryDate = entry[headers[dateCol].toLowerCase().replace(/\s+/g, "_")];

        // Skip if no narrative or missing key data
        if (!narrative || !entryFeeEarner || !entryDate) {
          console.log(
            `Skipping entry ${processedCount}: missing data - narrative: ${!!narrative}, feeEarner: ${!!entryFeeEarner}, date: ${!!entryDate}`
          );
          continue;
        }

        // Debug log for Callum Reyes entries
        if (narrative.includes("callum") || narrative.includes("reyes")) {
          console.log(
            `DEBUG: Found potential Callum Reyes entry - Row: ${entry.rowIndex}, Narrative: "${narrative}", FeeEarner: "${entryFeeEarner}", Date: "${entryDate}"`
          );
        }

        // Check if narrative contains meeting keywords (with word boundary check)
        const containsMeetingKeyword = missingTimeRule.meetingKeywords.some((keyword) => {
          // Escape special regex characters
          const escapeRegex = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

          // Create a regex for word boundary matching
          const keywordRegex = new RegExp(`\\b${escapeRegex(keyword.toLowerCase())}\\b`, "i");
          const matches = keywordRegex.test(narrative);
          if (matches && processedCount <= 5) {
            console.log(
              `  Entry ${processedCount}: Keyword "${keyword}" matched in: "${narrative.substring(0, 80)}..."`
            );
          }
          return matches;
        });

        if (containsMeetingKeyword) {
          meetingEntriesFound++;
          console.log(
            `Found meeting entry ${meetingEntriesFound}: "${narrative}" by ${entryFeeEarner} on ${entryDate}`
          );

          // Show first few meeting entries found (to avoid spam)
          if (meetingEntriesFound <= 3) {
            showMessage(
              `DEBUG: Found meeting entry ${meetingEntriesFound}: "${narrative.substring(0, 50)}..." by ${entryFeeEarner}`,
              "info"
            );
          }
          // Find mentioned fee earners in the narrative
          const mentionedFeeEarners = feeEarners
            .filter((feeEarner) => {
              const firstName = feeEarner.name.split(" ")[0].toLowerCase();
              const lastName = feeEarner.name.split(" ").slice(1).join(" ").toLowerCase();
              const fullName = feeEarner.name.toLowerCase();

              // Escape special regex characters
              const escapeRegex = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

              // Create word boundary regex patterns to match names more accurately
              const firstNameRegex = new RegExp(`\\b${escapeRegex(firstName)}\\b`, "i");
              const lastNameRegex = lastName
                ? new RegExp(`\\b${escapeRegex(lastName)}\\b`, "i")
                : null;
              const fullNameRegex = new RegExp(`\\b${escapeRegex(fullName)}\\b`, "i");

              // Check various name patterns
              const ismentioned =
                fullNameRegex.test(narrative) ||
                (firstNameRegex.test(narrative) && lastNameRegex && lastNameRegex.test(narrative));

              if (ismentioned) {
                console.log(
                  `  Found mentioned fee earner: ${feeEarner.name} in narrative: "${narrative.substring(0, 100)}..."`
                );
              } else {
                // Additional debug logging for names that weren't matched
                if (
                  narrative.toLowerCase().includes(firstName) ||
                  narrative.toLowerCase().includes(fullName)
                ) {
                  console.log(
                    `  DEBUG: ${feeEarner.name} partially matches but failed regex test in: "${narrative.substring(0, 100)}..."`
                  );
                }
              }
              return ismentioned;
            })
            .filter((feeEarner) => {
              const isDifferentPerson = feeEarner.name !== entryFeeEarner;
              if (!isDifferentPerson) {
                console.log(`  Excluding ${feeEarner.name} as they are the entry creator`);
              }
              return isDifferentPerson;
            });

          console.log(
            `  Final mentioned fee earners for reciprocal check: ${mentionedFeeEarners.map((fe) => fe.name).join(", ")}`
          );

          // Check if mentioned fee earners have reciprocal entries
          for (const mentionedFeeEarner of mentionedFeeEarners) {
            const hasReciprocalEntry = entries.some((otherEntry) => {
              const otherFeeEarner = (
                otherEntry[headers[feeEarnerCol].toLowerCase().replace(/\s+/g, "_")] || ""
              ).toString();
              const otherDate = otherEntry[headers[dateCol].toLowerCase().replace(/\s+/g, "_")];
              const otherNarrative = (
                otherEntry[headers[narrativeCol].toLowerCase().replace(/\s+/g, "_")] || ""
              )
                .toString()
                .toLowerCase();

              // Check if it's the right fee earner
              if (otherFeeEarner !== mentionedFeeEarner.name) return false;

              // Check date match (considering tolerance)
              const datesMatch = datesWithinTolerance(
                entryDate,
                otherDate,
                missingTimeRule.dateTolerance
              );
              if (!datesMatch && mentionedFeeEarner.name.toLowerCase().includes("callum")) {
                console.log(
                  `      Date mismatch for ${mentionedFeeEarner.name}: entry date ${entryDate} vs other date ${otherDate}`
                );
              }
              if (!datesMatch) return false;

              // Optionally check if the reciprocal narrative mentions the original fee earner
              const originalFirstName = entryFeeEarner.split(" ")[0].toLowerCase();
              const originalFullName = entryFeeEarner.toLowerCase();

              return (
                otherNarrative.includes(originalFirstName) ||
                otherNarrative.includes(originalFullName) ||
                missingTimeRule.meetingKeywords.some((keyword) =>
                  otherNarrative.includes(keyword.toLowerCase())
                )
              );
            });

            if (!hasReciprocalEntry) {
              console.log(
                `    Missing reciprocal entry: ${mentionedFeeEarner.name} should have entry for ${entryDate}`
              );
              missingEntries.push({
                originalEntry: entry,
                missingFeeEarner: mentionedFeeEarner,
                date: entryDate,
                narrative: narrative,
              });
            } else {
              console.log(`    Found reciprocal entry for ${mentionedFeeEarner.name}`);
            }
          }
        }
      }

      console.log(
        `Processing complete: ${processedCount} entries processed, ${meetingEntriesFound} meeting entries found, ${missingEntries.length} missing reciprocal entries identified`
      );

      // Show detailed processing results in the UI
      showMessage(
        `DEBUG: Processed ${processedCount} entries, found ${meetingEntriesFound} meeting entries, identified ${missingEntries.length} missing reciprocal entries`,
        "info"
      );

      // Add notes to rows about missing entries and apply formatting
      // Get the header row to find Notes column
      const headerRow = headers;
      let notesColumnIndex = findNotesColumn(headerRow);

      if (notesColumnIndex === -1) {
        console.log("Notes column not found, creating new one...");
        notesColumnIndex = await createNotesColumnWithFormatting(worksheet, usedRange);
        // Reload used range after adding column
        usedRange.load(["values", "rowCount", "columnCount"]);
        await context.sync();
        console.log(`Created Notes column at index ${notesColumnIndex}`);
      } else {
        console.log(`Found existing Notes column at index ${notesColumnIndex}`);
      }

      let updatedCount = 0;
      for (const missing of missingEntries) {
        const rowIndex = missing.originalEntry.rowIndex;
        const noteText = `Missing Time: ${missing.missingFeeEarner.name} should have entry for ${missing.date}`;

        // Get current Notes cell value and update it
        const notesCell = worksheet.getCell(rowIndex, notesColumnIndex);
        notesCell.load("values");
        await context.sync();

        const existingNotes = notesCell.values[0][0]?.toString() || "";
        const updatedNotes = addNoteToRow(existingNotes, noteText);

        console.log(
          `Row ${rowIndex}: Existing notes: "${existingNotes}", Adding: "${noteText}", Result: "${updatedNotes}"`
        );

        if (updatedNotes !== existingNotes) {
          // Update the Notes cell
          notesCell.values = [[updatedNotes]];

          console.log(
            `Updated Notes cell at row ${rowIndex}, column ${notesColumnIndex} with: "${updatedNotes}"`
          );

          // Apply pale blue formatting to the entire row to highlight missing time
          const rowRange = worksheet.getCell(rowIndex, 0).getEntireRow();
          rowRange.format.fill.color = "#E3F2FD"; // Pale blue background

          await context.sync();
          updatedCount++;

          console.log(`Applied pale blue formatting to row ${rowIndex}`);
        } else {
          console.log(`No change needed for row ${rowIndex} - note already exists`);
        }

        // Create placeholder entry if enabled
        if (missingTimeRule.createMissingEntries) {
          await createPlaceholderEntry(worksheet, missing, rowIndex, headers, currentProfile);
          await context.sync();
        }
      }

      return { success: true, updatedRows: updatedCount };
    });
  } catch (error) {
    console.error("Error applying missing time entries rule:", error);
    return { success: false, updatedRows: 0, error: error.message };
  }
}

// Helper function to create placeholder entry for missing time
async function createPlaceholderEntry(
  worksheet: Excel.Worksheet,
  missing: any,
  parentRowIndex: number,
  headers: string[],
  currentProfile: any
) {
  // Insert a new row right after the parent entry
  const insertRowIndex = parentRowIndex + 1;
  const insertRange = worksheet.getCell(insertRowIndex, 0).getEntireRow();
  insertRange.insert(Excel.InsertShiftDirection.down);

  // Get original entry data
  const originalEntry = missing.originalEntry;
  const originalFeeEarner =
    originalEntry[
      headers
        .find(
          (h) =>
            (h.toLowerCase().includes("fee") && h.toLowerCase().includes("earner")) ||
            h.toLowerCase().includes("name")
        )
        ?.toLowerCase()
        .replace(/\s+/g, "_")
    ] || "";

  const originalNarrative =
    originalEntry[
      headers
        .find(
          (h) => h.toLowerCase().includes("narrative") || h.toLowerCase().includes("description")
        )
        ?.toLowerCase()
        .replace(/\s+/g, "_")
    ] || "";

  // Create new row data with swapped names
  for (let colIndex = 0; colIndex < headers.length; colIndex++) {
    const header = headers[colIndex];
    const headerKey = header.toLowerCase().replace(/\s+/g, "_");
    let cellValue = originalEntry[headerKey];

    // Handle fee earner column - swap to missing person
    if (
      (header.toLowerCase().includes("fee") && header.toLowerCase().includes("earner")) ||
      header.toLowerCase().includes("name")
    ) {
      cellValue = missing.missingFeeEarner.name;
    }
    // Handle narrative column - swap names in narrative
    else if (
      header.toLowerCase().includes("narrative") ||
      header.toLowerCase().includes("description")
    ) {
      if (cellValue) {
        // Replace original fee earner name with missing fee earner name in narrative
        let swappedNarrative = cellValue.toString();

        // Try to replace first name
        const originalFirstName = originalFeeEarner.split(" ")[0];
        const missingFirstName = missing.missingFeeEarner.name.split(" ")[0];
        swappedNarrative = swappedNarrative.replace(
          new RegExp(originalFirstName, "gi"),
          missingFirstName
        );

        // Also try full name replacement
        swappedNarrative = swappedNarrative.replace(
          new RegExp(originalFeeEarner, "gi"),
          missing.missingFeeEarner.name
        );

        cellValue = swappedNarrative;
      }
    }

    // Set the cell value
    if (cellValue !== undefined && cellValue !== null) {
      const newCell = worksheet.getCell(insertRowIndex, colIndex);
      newCell.values = [[cellValue]];
    }
  }

  // Apply consistent formatting to match parent row
  const parentRow = worksheet.getCell(parentRowIndex, 0).getEntireRow();
  const newRow = worksheet.getCell(insertRowIndex, 0).getEntireRow();

  // Copy formatting from parent (this is basic - could be enhanced)
  if (currentProfile.enableAlternatingRows !== false) {
    const altRowColor1 = currentProfile.altRowColor1 || "#FFFFFF";
    const altRowColor2 = currentProfile.altRowColor2 || "#F8F9FA";
    // Use alternating color logic
    const useAltColor = insertRowIndex % 2 === 0;
    newRow.format.fill.color = useAltColor ? altRowColor2 : altRowColor1;
  }
}

// Helper function to check if dates are within tolerance
function datesWithinTolerance(date1: any, date2: any, toleranceDays: number): boolean {
  try {
    const d1 = new Date(date1);
    const d2 = new Date(date2);

    if (isNaN(d1.getTime()) || isNaN(d2.getTime())) {
      return false; // Invalid dates
    }

    const diffMs = Math.abs(d1.getTime() - d2.getTime());
    const diffDays = diffMs / (1000 * 60 * 60 * 24);

    return diffDays <= toleranceDays;
  } catch {
    return false;
  }
}

// Modified apply function to support undo
async function applyNameStandardisationToWorksheetWithUndo(): Promise<number> {
  try {
    // Get current matter and its rules
    const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;
    if (!selectedMatter) {
      showMessage(
        "Please select a matter profile before applying name standardisation rules.",
        "error"
      );
      return 0;
    }

    const profiles = getMatterProfiles();
    const currentProfile = profiles.find((p) => p.name === selectedMatter);
    if (!currentProfile || !currentProfile.rules) {
      showMessage("No rules found for the selected matter profile.", "error");
      return 0;
    }

    const nameRule = currentProfile.rules.nameStandardisation;
    if (!nameRule.enabled) {
      showMessage("Name standardisation rule is not enabled for this matter.", "info");
      return 0;
    }

    const feeEarners = currentProfile.feeEarners || [];
    if (feeEarners.length === 0) {
      showMessage("No fee earners found for this matter. Please add fee earners first.", "error");
      return 0;
    }

    return await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = worksheet.getUsedRange();

      if (!usedRange) {
        showMessage("No data found in the worksheet to process.", "error");
        return 0;
      }

      usedRange.load(["values", "formulas", "rowCount", "columnCount"]);
      await context.sync();

      // Convert Excel data to a workable format
      const headers = usedRange.values[0] as string[];
      const worksheetData: any[] = [];

      for (let i = 1; i < usedRange.values.length; i++) {
        const row = usedRange.values[i];
        const rowData: any = {};

        headers.forEach((header, colIndex) => {
          rowData[header] = row[colIndex];
        });

        worksheetData.push(rowData);
      }

      // Apply name standardisation rules
      const processedData = applyNameStandardisationRule(worksheetData, feeEarners, nameRule);

      // Create undo snapshot
      const undoSnapshot: UndoSnapshot = {
        timestamp: Date.now(),
        changes: [],
      };

      // Update the worksheet with processed data and track changes
      let updatedCount = 0;

      // Check if "Amended Narrative" column exists
      let amendedNarrativeCol = headers.findIndex(
        (h) => h.toLowerCase().includes("amended") && h.toLowerCase().includes("narrative")
      );

      // If column doesn't exist, we need to check if we need to create it
      const needsAmendedColumn =
        amendedNarrativeCol < 0 &&
        processedData.some((row) => row["Amended Narrative"] !== undefined);

      if (needsAmendedColumn) {
        // Handle column creation (same as before)
        const sourceCol = headers.findIndex(
          (h) =>
            (h.toLowerCase().includes("original") && h.toLowerCase().includes("narrative")) ||
            (h.toLowerCase().includes("narrative") && !h.toLowerCase().includes("amended"))
        );

        if (sourceCol >= 0) {
          const insertCol = sourceCol + 1;
          const newColumn = worksheet.getCell(0, insertCol).getEntireColumn();
          newColumn.insert(Excel.InsertShiftDirection.right);

          const headerCell = worksheet.getCell(0, insertCol);
          headerCell.values = [["Amended Narrative"]];

          amendedNarrativeCol = insertCol;
          headers.splice(insertCol, 0, "Amended Narrative");

          await context.sync();
        }
      }

      // First, update the Amended Narrative data
      for (let i = 0; i < processedData.length; i++) {
        const processedRow = processedData[i];
        const amendedValue = processedRow["Amended Narrative"];

        if (amendedValue !== undefined && amendedNarrativeCol >= 0) {
          // Get the old value for undo
          const oldValue = usedRange.values[i + 1][amendedNarrativeCol] || "";

          if (amendedValue !== oldValue) {
            // Track this change for undo
            undoSnapshot.changes.push({
              row: i + 1,
              column: amendedNarrativeCol,
              oldValue: oldValue,
              newValue: amendedValue,
            });

            // Update the cell
            const cell = worksheet.getCell(i + 1, amendedNarrativeCol);
            cell.values = [[amendedValue]];
            updatedCount++;
          }
        }
      }

      // Sync the amended narrative changes before adding notes
      await context.sync();

      // Find Notes column (should exist if Name Standardisation is enabled)
      let notesCol = findNotesColumn(headers);

      // Now add notes for rows that were changed
      for (let i = 0; i < processedData.length; i++) {
        const processedRow = processedData[i];
        const amendedValue = processedRow["Amended Narrative"];

        if (amendedValue !== undefined && amendedNarrativeCol >= 0 && notesCol >= 0) {
          // Check if this row had changes
          const hadChanges = undoSnapshot.changes.some(
            (change) => change.row === i + 1 && change.column === amendedNarrativeCol
          );

          if (hadChanges) {
            // Get current Notes cell value
            const notesCell = worksheet.getCell(i + 1, notesCol);
            notesCell.load("values");
            await context.sync();

            const existingNotes = notesCell.values[0][0]?.toString() || "";
            const updatedNotes = addNoteToRow(existingNotes, "Name Standardised");

            if (updatedNotes !== existingNotes) {
              // Track Notes column change for undo
              undoSnapshot.changes.push({
                row: i + 1,
                column: notesCol,
                oldValue: existingNotes,
                newValue: updatedNotes,
              });

              // Update the Notes cell
              notesCell.values = [[updatedNotes]];
            }
          }
        }
      }

      await context.sync();

      // Store undo snapshot only if changes were made
      if (undoSnapshot.changes.length > 0) {
        lastUndoSnapshot = undoSnapshot;
        updateUndoButtonState();
      }

      if (updatedCount > 0) {
        showMessage(
          `Name standardisation applied successfully. Updated ${updatedCount} rows. Undo is available.`,
          "success"
        );
      } else {
        showMessage("Name standardisation completed, but no changes were needed.", "info");
      }

      return updatedCount;
    });
  } catch (error) {
    console.error("Error applying name standardisation:", error);
    showMessage("An error occurred while applying name standardisation: " + error.message, "error");
    return 0;
  }
}

// Rules Management Functions
function getDefaultRules(): RulesConfig {
  return {
    nameStandardisation: {
      enabled: false,
      caseSensitive: false,
      allowPartialMatches: true,
      useDateMatching: true,
      replaceOnlyFirstOccurrence: true,
      excludedNames: [],
      minPartialMatchLength: 3,
      useNicknameDatabase: true,
      customNicknames: {},
    },
    missingTimeEntries: {
      enabled: false,
      dateTolerance: 0, // exact date match
      meetingKeywords: ["meeting", "call", "conference", "discussion", "telephone", "phone"],
      requireExactTimeMatch: false,
      createMissingEntries: false,
    },
  };
}

function getCurrentRules(): RulesConfig {
  const excludedNamesText = (document.getElementById("excluded-names") as HTMLInputElement).value;
  const excludedNames = excludedNamesText
    .split(",")
    .map((name) => name.trim())
    .filter((name) => name.length > 0);

  const minPartialMatchLength = parseInt(
    (document.getElementById("min-partial-match-length") as HTMLInputElement).value || "3",
    10
  );

  return {
    nameStandardisation: {
      enabled: (document.getElementById("name-standardisation-enabled") as HTMLInputElement)
        .checked,
      caseSensitive: false, // Case sensitivity removed from Stage 1
      allowPartialMatches: (document.getElementById("partial-matches") as HTMLInputElement).checked,
      useDateMatching: (document.getElementById("date-matching") as HTMLInputElement).checked,
      replaceOnlyFirstOccurrence: (
        document.getElementById("first-occurrence-only") as HTMLInputElement
      ).checked,
      excludedNames: excludedNames,
      minPartialMatchLength: minPartialMatchLength,
      useNicknameDatabase: (document.getElementById("use-nickname-database") as HTMLInputElement)
        .checked,
      customNicknames: getCurrentNicknames(),
    },
    missingTimeEntries: {
      enabled:
        (document.getElementById("missing-time-entries-enabled") as HTMLInputElement)?.checked ||
        false,
      dateTolerance: parseInt(
        (document.getElementById("date-tolerance") as HTMLInputElement)?.value || "0",
        10
      ),
      meetingKeywords: (
        (document.getElementById("meeting-keywords") as HTMLInputElement)?.value ||
        "meeting,call,conference,discussion,telephone,phone"
      )
        .split(",")
        .map((keyword) => keyword.trim())
        .filter((keyword) => keyword.length > 0),
      requireExactTimeMatch:
        (document.getElementById("exact-time-match") as HTMLInputElement)?.checked || false,
      createMissingEntries:
        (document.getElementById("create-missing-entries") as HTMLInputElement)?.checked || false,
    },
  };
}

function loadRulesConfig(rules: RulesConfig) {
  // Load Name Standardisation rule settings
  const nameRule = rules.nameStandardisation;

  (document.getElementById("name-standardisation-enabled") as HTMLInputElement).checked =
    nameRule.enabled;
  (document.getElementById("partial-matches") as HTMLInputElement).checked =
    nameRule.allowPartialMatches;
  (document.getElementById("date-matching") as HTMLInputElement).checked = nameRule.useDateMatching;
  (document.getElementById("first-occurrence-only") as HTMLInputElement).checked =
    nameRule.replaceOnlyFirstOccurrence;
  (document.getElementById("excluded-names") as HTMLInputElement).value =
    nameRule.excludedNames.join(", ");
  (document.getElementById("min-partial-match-length") as HTMLInputElement).value = (
    nameRule.minPartialMatchLength || 3
  ).toString();
  (document.getElementById("use-nickname-database") as HTMLInputElement).checked =
    nameRule.useNicknameDatabase !== false;

  // Load nickname database
  const customNicknames = nameRule.customNicknames || {};
  loadNicknameDatabase(customNicknames);

  // Show/hide configuration based on enabled state
  const configDiv = document.getElementById("name-standardisation-content");
  configDiv.style.display = nameRule.enabled ? "block" : "none";

  // Show/hide nickname database config
  const nicknameConfigDiv = document.getElementById("nickname-database-config");
  nicknameConfigDiv.style.display = nameRule.useNicknameDatabase !== false ? "block" : "none";

  // Load Missing Time Entries rule settings (with null checks for backward compatibility)
  const missingTimeRule = rules.missingTimeEntries || getDefaultRules().missingTimeEntries;

  const missingTimeEnabledEl = document.getElementById(
    "missing-time-entries-enabled"
  ) as HTMLInputElement;
  if (missingTimeEnabledEl) {
    missingTimeEnabledEl.checked = missingTimeRule.enabled;
  }

  const dateToleranceEl = document.getElementById("date-tolerance") as HTMLInputElement;
  if (dateToleranceEl) {
    dateToleranceEl.value = missingTimeRule.dateTolerance.toString();
  }

  const meetingKeywordsEl = document.getElementById("meeting-keywords") as HTMLInputElement;
  if (meetingKeywordsEl) {
    meetingKeywordsEl.value = missingTimeRule.meetingKeywords.join(", ");
  }

  const exactTimeMatchEl = document.getElementById("exact-time-match") as HTMLInputElement;
  if (exactTimeMatchEl) {
    exactTimeMatchEl.checked = missingTimeRule.requireExactTimeMatch;
  }

  const createMissingEntriesEl = document.getElementById(
    "create-missing-entries"
  ) as HTMLInputElement;
  if (createMissingEntriesEl) {
    createMissingEntriesEl.checked = missingTimeRule.createMissingEntries;
  }

  // Show/hide missing time entries configuration based on enabled state
  const missingTimeConfigDiv = document.getElementById("missing-time-entries-content");
  if (missingTimeConfigDiv) {
    missingTimeConfigDiv.style.display = missingTimeRule.enabled ? "block" : "none";
  }
}

function saveRuleSettings() {
  const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;

  if (!selectedMatter) {
    showMessage("Please select a matter from the dropdown to save rules to.", "error");
    return;
  }

  // Get current rules from the form
  const currentRules = getCurrentRules();

  // Get existing matter profiles
  const profiles = getMatterProfiles();
  const existingIndex = profiles.findIndex((p) => p.name === selectedMatter);

  if (existingIndex >= 0) {
    // Update the existing profile with new rules data
    profiles[existingIndex].rules = currentRules;
    saveMatterProfiles(profiles);

    // Generate detailed feedback about what was saved
    const enabledRules = [];
    if (currentRules.nameStandardisation?.enabled) {
      enabledRules.push("Name Standardisation");
    }
    if (currentRules.missingTimeEntries?.enabled) {
      enabledRules.push("Missing Time Entries");
    }

    const ruleCount = enabledRules.length;
    const rulesList = enabledRules.length > 0 ? ` (${enabledRules.join(", ")})` : "";

    showMessage(
      `Successfully saved rule settings${rulesList} to matter profile "${selectedMatter}". ${ruleCount} rule${ruleCount !== 1 ? "s" : ""} enabled.`,
      "success"
    );
  } else {
    showMessage("Selected matter profile not found. Please create a new profile first.", "error");
  }
}
