/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office, HTMLInputElement, HTMLSelectElement, HTMLElement, setTimeout, localStorage */

// Track whether a matter is currently loaded
let currentMatterLoaded: string | null = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("format-spreadsheet").onclick = formatSpreadsheet;
    document.getElementById("add-charge-column").onclick = addChargeColumn;
    document.getElementById("color-code-rows").onclick = colorCodeRows;

    // Matter profile functionality
    document.getElementById("save-matter").onclick = saveMatterProfile;
    document.getElementById("delete-matter").onclick = deleteMatterProfile;
    document.getElementById("save-current-settings").onclick = saveCurrentSettings;

    // Handle matter selection from dropdown
    const matterSelect = document.getElementById("matter-select") as HTMLSelectElement;
    matterSelect.onchange = () => {
      if (matterSelect.value) {
        loadMatterProfile();
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

      // Get current matter settings for formatting
      const headerBgColor = (document.getElementById("header-bg-color") as HTMLInputElement).value;
      const headerTextColor = (document.getElementById("header-text-color") as HTMLInputElement)
        .value;
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

  // Update prepopulation rules visibility based on checkbox state
  const rulesDiv = document.getElementById("prepopulate-rules");
  rulesDiv.style.display = profile.prepopulateCharge || false ? "block" : "none";
}

function loadMatterProfiles() {
  const profiles = getMatterProfiles();
  const selectElement = document.getElementById("matter-select") as HTMLSelectElement;

  // Clear existing options except the first one
  selectElement.innerHTML = '<option value="">-- Select a Matter --</option>';

  // Add saved profiles
  profiles.forEach((profile) => {
    const option = document.createElement("option");
    option.value = profile.name;
    option.textContent = profile.name;
    selectElement.appendChild(option);
  });

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
