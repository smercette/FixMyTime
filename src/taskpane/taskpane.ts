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
  try {
    console.log("Office.onReady called, host:", info.host);
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("apply-formatting").onclick = applyFormatting;
    document.getElementById("apply-all-rules").onclick = applyAllRules;
    // Clear any previous debug info on initialization
    clearDebugInfo();
    
    // Test debug functionality
    console.log("FixMyTime add-in Office.onReady called");
    addDebugInfo("FixMyTime add-in loaded successfully");

    // Matter profile functionality
    document.getElementById("save-matter").onclick = saveMatterProfile;
    document.getElementById("delete-matter").onclick = deleteMatterProfile;
    document.getElementById("save-current-settings").onclick = saveCurrentSettings;
    document.getElementById("save-matter-profile").onclick = saveMatterProfileFromDropdown;

    // Debug: Add a function to check localStorage
    (window as any).debugMatterProfiles = () => {
      const profiles = getMatterProfiles();
      addDebugInfo(`Current matter profiles: ${JSON.stringify(profiles, null, 2)}`);
      const stored = localStorage.getItem("fixmytime-matter-profiles");
      addDebugInfo(`Raw localStorage data: ${stored}`);
      return profiles;
    };

    // Fee Earners functionality
    document.getElementById("add-fee-earner").onclick = () => addFeeEarnerRow();
    document.getElementById("update-from-spreadsheet").onclick = updateFeeEarnersFromSpreadsheet;
    document.getElementById("save-participants").onclick = saveParticipants;

    // Rules functionality
    document.getElementById("save-rule-settings").onclick = saveRuleSettings;
    document.getElementById("undo-name-rules").onclick = undoNameStandardisation;

    // Nickname database functionality (temporarily commented out due to missing functions)
    // document.getElementById("add-nickname").onclick = addNicknameEntry;
    // document.getElementById("reset-nicknames").onclick = resetNicknamesToDefault;

    // Nickname database toggle
    const nicknameToggle = document.getElementById("use-nickname-database") as HTMLInputElement;
    nicknameToggle.onchange = () => {
      const configDiv = document.getElementById("nickname-database-config");
      configDiv.style.display = nicknameToggle.checked ? "block" : "none";
    };

    // TimeFormat rule toggle
    const timeFormatToggle = document.getElementById("time-format-enabled") as HTMLInputElement;
    if (timeFormatToggle) {
      timeFormatToggle.onchange = () => {
        const configDiv = document.getElementById("time-format-content");
        if (configDiv) {
          configDiv.style.display = timeFormatToggle.checked ? "block" : "none";
        }
      };
    }

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

    // NeedsDetail rule toggle
    const needsDetailToggle = document.getElementById("needs-detail-enabled") as HTMLInputElement;
    if (needsDetailToggle) {
      needsDetailToggle.onchange = () => {
        const configDiv = document.getElementById("needs-detail-content");
        if (configDiv) {
          configDiv.style.display = needsDetailToggle.checked ? "block" : "none";
        }
      };
    }

    // Travel rule toggle
    const travelToggle = document.getElementById("travel-enabled") as HTMLInputElement;
    if (travelToggle) {
      travelToggle.onchange = () => {
        const configDiv = document.getElementById("travel-content");
        if (configDiv) {
          configDiv.style.display = travelToggle.checked ? "block" : "none";
        }
      };
    }

    // Non Chargeable rule toggle
    const nonChargeableToggle = document.getElementById(
      "non-chargeable-enabled"
    ) as HTMLInputElement;
    if (nonChargeableToggle) {
      nonChargeableToggle.onchange = () => {
        const configDiv = document.getElementById("non-chargeable-content");
        if (configDiv) {
          configDiv.style.display = nonChargeableToggle.checked ? "block" : "none";
        }
      };
    }

    // Max Daily Hours rule toggle
    const maxDailyHoursToggle = document.getElementById(
      "max-daily-hours-enabled"
    ) as HTMLInputElement;
    if (maxDailyHoursToggle) {
      maxDailyHoursToggle.onchange = () => {
        const configDiv = document.getElementById("max-daily-hours-content");
        if (configDiv) {
          configDiv.style.display = maxDailyHoursToggle.checked ? "block" : "none";
        }
      };
    }

    // Make functions available globally for onclick handlers
    (window as any).removeFeeEarnerRow = removeFeeEarnerRow;
    // (window as any).removeNicknameEntry = removeNicknameEntry; // temporarily commented out
    // Handle matter selection from dropdown
    const matterSelect = document.getElementById("matter-select") as HTMLSelectElement;
    matterSelect.onchange = handleMatterSelect;

    // Initialize with current state
    currentMatterLoaded = matterSelect.value || null;
    updateUIForMatterState();

    // Update matter dropdown with saved profiles
    updateMatterDropdown();
    
    // Set up tab functionality
    const tabButtons = document.querySelectorAll('.tab-button');
    tabButtons.forEach(button => {
      button.addEventListener('click', function() {
        const targetTab = this.getAttribute('data-tab');
        
        // Remove active class from all buttons and content
        tabButtons.forEach(btn => btn.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        
        // Add active class to clicked button and corresponding content
        this.classList.add('active');
        const targetContent = document.getElementById(targetTab + '-tab');
        if (targetContent) {
          targetContent.classList.add('active');
        }
        
        addDebugInfo(`Switched to ${targetTab} tab`);
      });
    });
    
    // Set up dropdown functionality for Settings tab
    const dropdownHeaders = document.querySelectorAll('.dropdown-header[data-target]');
    dropdownHeaders.forEach(header => {
      // Add cursor pointer to make it clear it's clickable
      header.style.cursor = 'pointer';
      
      header.addEventListener('click', function() {
        const targetId = this.getAttribute('data-target');
        const targetContent = document.getElementById(targetId);
        const arrow = this.querySelector('.dropdown-arrow');
        
        if (targetContent && arrow) {
          // Check current visibility (accounting for inline style and computed style)
          const currentDisplay = targetContent.style.display;
          const isVisible = currentDisplay !== 'none' && currentDisplay !== '';
          
          // Toggle visibility
          targetContent.style.display = isVisible ? 'none' : 'block';
          
          // Toggle arrow direction
          arrow.textContent = isVisible ? '▼' : '▲';
          
          addDebugInfo(`Toggled dropdown: ${targetId} - now ${isVisible ? 'hidden' : 'visible'}`);
        }
      });
      
      // Set initial arrow direction based on content visibility
      const targetId = header.getAttribute('data-target');
      const targetContent = document.getElementById(targetId);
      const arrow = header.querySelector('.dropdown-arrow');
      
      if (targetContent && arrow) {
        const currentDisplay = targetContent.style.display;
        const isVisible = currentDisplay !== 'none' && currentDisplay !== '';
        arrow.textContent = isVisible ? '▲' : '▼';
      }
    });
  }
  } catch (error) {
    console.error("Error in Office.onReady:", error);
    // Show error visibly in the task pane
    const appBody = document.getElementById("app-body");
    if (appBody) {
      const errorDiv = document.createElement("div");
      errorDiv.style.backgroundColor = "#ffebee";
      errorDiv.style.color = "#c62828";
      errorDiv.style.padding = "10px";
      errorDiv.style.border = "1px solid #f44336";
      errorDiv.style.marginBottom = "10px";
      errorDiv.innerHTML = `<strong>JavaScript Error:</strong><br>${error.message}<br><br>Stack: ${error.stack}`;
      appBody.insertBefore(errorDiv, appBody.firstChild);
    }
  }
});

// Debug functionality
function clearDebugInfo() {
  const debugDiv = document.querySelector(".debug-section .debug-content");
  if (debugDiv) {
    debugDiv.innerHTML = "";
  }
}

function addDebugInfo(message: string) {
  // Make debug area visible
  const debugArea = document.getElementById("debug-area");
  if (debugArea) {
    debugArea.style.display = "block";
  }
  
  // Add debug message
  const debugDiv = document.getElementById("debug-content");
  if (debugDiv) {
    const timestamp = new Date().toLocaleTimeString();
    const debugLine = document.createElement("div");
    debugLine.innerHTML = `[${timestamp}] ${message}`;
    debugDiv.appendChild(debugLine);

    // Auto-scroll to bottom
    debugDiv.scrollTop = debugDiv.scrollHeight;
  } else {
    // Fallback to console if debug div not found
    console.log(`[DEBUG] ${message}`);
  }
}

// Utility function to get column index by header name (case insensitive)
function getColumnIndex(headers: string[], columnName: string): number {
  return headers.findIndex(
    (header) => header && header.toString().toLowerCase() === columnName.toLowerCase()
  );
}

// Helper function to safely get cell value
function getCellValue(values: any[][], row: number, col: number): string {
  if (row >= 0 && row < values.length && col >= 0 && col < values[row].length) {
    const value = values[row][col];
    return value ? value.toString() : "";
  }
  return "";
}

// Validate that required columns exist
function validateColumns(
  headers: string[],
  requiredColumns: string[]
): { isValid: boolean; missingColumns: string[] } {
  const missingColumns = requiredColumns.filter(
    (col) => !headers.some((h) => h && h.toLowerCase() === col.toLowerCase())
  );
  return {
    isValid: missingColumns.length === 0,
    missingColumns,
  };
}

function updateUIForMatterState() {
  const quickActionsSection = document.getElementById("quick-actions-section");

  if (currentMatterLoaded) {
    // Update current matter display
    const currentMatterName = document.getElementById("current-matter-name");
    if (currentMatterName) {
      currentMatterName.textContent = currentMatterLoaded;
    }
    
    // Show the Quick Actions section with buttons
    if (quickActionsSection) {
      quickActionsSection.style.display = "block";
    }

    // Show matter selection in Settings tab header if not already there
    const settingsHeader = document.querySelector("#settings-tab .settings-matter-info");
    if (settingsHeader) {
      settingsHeader.textContent = `Current Matter: ${currentMatterLoaded}`;
      settingsHeader.style.display = "block";
    }
  } else {
    // Hide the Quick Actions section when no matter is selected
    if (quickActionsSection) {
      quickActionsSection.style.display = "none";
    }

    // Hide matter info in Settings tab
    const settingsHeader = document.querySelector("#settings-tab .settings-matter-info");
    if (settingsHeader) {
      settingsHeader.style.display = "none";
    }
  }
}

function handleMatterSelect() {
  const matterSelect = document.getElementById("matter-select") as HTMLSelectElement;
  const selectedMatter = matterSelect.value;
  
  console.log(`handleMatterSelect called with value: "${selectedMatter}"`);
  addDebugInfo(`handleMatterSelect called with value: "${selectedMatter}"`);

  if (selectedMatter === "__new__") {
    // Handle "Add New Matter" selection
    console.log("Detected Add New Matter selection - calling showAddNewMatterUI");
    addDebugInfo("Detected Add New Matter selection - calling showAddNewMatterUI");
    showAddNewMatterUI();
    addDebugInfo("Add New Matter selected - showing creation UI");
  } else if (selectedMatter) {
    currentMatterLoaded = selectedMatter;

    // Load matter profile settings into the form
    const profiles = getMatterProfiles();
    const profile = profiles.find((p) => p.name === selectedMatter);

    if (profile) {
      loadMatterProfile(profile);
      addDebugInfo(`Loaded matter profile: ${selectedMatter}`);
    }
  } else {
    currentMatterLoaded = null;
  }

  updateUIForMatterState();
}

function showAddNewMatterUI() {
  addDebugInfo("showAddNewMatterUI function called");
  
  // Switch to the Settings tab
  const settingsTab = document.querySelector('[data-tab="settings"]') as HTMLElement;
  const mainTab = document.querySelector('[data-tab="main"]') as HTMLElement;
  const settingsContent = document.getElementById("settings-tab") as HTMLElement;
  const mainContent = document.getElementById("main-tab") as HTMLElement;
  
  addDebugInfo(`Found elements - settingsTab: ${!!settingsTab}, mainTab: ${!!mainTab}, settingsContent: ${!!settingsContent}, mainContent: ${!!mainContent}`);
  
  if (settingsTab && mainTab && settingsContent && mainContent) {
    // Update tab buttons
    mainTab.classList.remove("active");
    settingsTab.classList.add("active");
    
    // Update tab content
    mainContent.classList.remove("active");
    settingsContent.classList.add("active");
    
    // Show the new matter section
    const newMatterSection = document.getElementById("new-matter-section") as HTMLElement;
    if (newMatterSection) {
      newMatterSection.style.display = "block";
    }
    
    // Clear and focus the new matter name input
    const newMatterInput = document.getElementById("new-matter-name") as HTMLInputElement;
    if (newMatterInput) {
      newMatterInput.value = "";
      newMatterInput.focus();
    }
    
    // Reset the dropdown to default state
    const matterSelect = document.getElementById("matter-select") as HTMLSelectElement;
    if (matterSelect) {
      matterSelect.value = "";
    }
  }
}

function showMessage(message: string, type: "success" | "error" | "info" = "info") {
  // Create or update the message element
  let messageEl = document.getElementById("message-display");
  if (!messageEl) {
    messageEl = document.createElement("div");
    messageEl.id = "message-display";
    messageEl.style.cssText = `
      position: fixed;
      top: 10px;
      right: 10px;
      padding: 10px 15px;
      border-radius: 4px;
      font-weight: bold;
      z-index: 1000;
      max-width: 300px;
      word-wrap: break-word;
      box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    `;
    document.body.appendChild(messageEl);
  }

  // Set message and style based on type
  messageEl.textContent = message;
  messageEl.className = `message-${type}`;

  // Style based on type
  switch (type) {
    case "success":
      messageEl.style.backgroundColor = "#d4edda";
      messageEl.style.color = "#155724";
      messageEl.style.border = "1px solid #c3e6cb";
      break;
    case "error":
      messageEl.style.backgroundColor = "#f8d7da";
      messageEl.style.color = "#721c24";
      messageEl.style.border = "1px solid #f5c6cb";
      break;
    default: // info
      messageEl.style.backgroundColor = "#d1ecf1";
      messageEl.style.color = "#0c5460";
      messageEl.style.border = "1px solid #bee5eb";
  }

  // Show the message
  messageEl.style.display = "block";

  // Hide after 3 seconds
  setTimeout(() => {
    if (messageEl && messageEl.parentNode) {
      messageEl.style.display = "none";
    }
  }, 3000);
}

async function applyFormatting() {
  try {
    addDebugInfo("Starting formatting process...");
    
    // Execute all three formatting operations in sequence
    addDebugInfo("Step 1: Applying spreadsheet formatting...");
    await formatSpreadsheet();
    
    addDebugInfo("Step 2: Adding columns...");
    await addColumns();
    
    addDebugInfo("Step 3: Color coding rows...");
    await colorCodeRows();

    addDebugInfo("Formatting completed successfully!");
    showMessage("Formatting applied successfully.", "success");
  } catch (error) {
    console.error("Error applying formatting:", error);
    addDebugInfo(`Formatting error: ${error.message}`);
    showMessage("An error occurred while applying formatting: " + error.message, "error");
  }
}

export async function formatSpreadsheet() {
  try {
    addDebugInfo("formatSpreadsheet: Starting Excel.run...");
    await Excel.run(async (context) => {
      addDebugInfo("formatSpreadsheet: Getting active worksheet...");
      // Get the active worksheet
      const worksheet = context.workbook.worksheets.getActiveWorksheet();

      addDebugInfo("formatSpreadsheet: Getting used range...");
      // Get the used range
      const usedRange = worksheet.getUsedRange();
      usedRange.load(["rowCount", "columnCount"]);

      addDebugInfo("formatSpreadsheet: First context.sync...");
      await context.sync();

      addDebugInfo("formatSpreadsheet: Checking if usedRange exists...");
      if (!usedRange) {
        showMessage("No data found in the worksheet to format.", "error");
        return;
      }

      addDebugInfo("formatSpreadsheet: Getting user configuration values...");
      // Get user configuration values with defaults in case elements don't exist
      const headerBgColorEl = document.getElementById("header-bg-color") as HTMLInputElement;
      const headerBgColor = headerBgColorEl ? headerBgColorEl.value : "#0078d4";
      
      const headerTextColorEl = document.getElementById("header-text-color") as HTMLInputElement;
      const headerTextColor = headerTextColorEl ? headerTextColorEl.value : "#ffffff";
      
      const altRowColor1El = document.getElementById("alt-row-color1") as HTMLInputElement;
      const altRowColor1 = altRowColor1El ? altRowColor1El.value : "#f9f9f9";
      
      const altRowColor2El = document.getElementById("alt-row-color2") as HTMLInputElement;
      const altRowColor2 = altRowColor2El ? altRowColor2El.value : "#ffffff";
      
      const borderColorEl = document.getElementById("border-color") as HTMLInputElement;
      const borderColor = borderColorEl ? borderColorEl.value : "#d1d1d1";
      
      const maxColumnWidthEl = document.getElementById("max-column-width") as HTMLInputElement;
      const maxColumnWidth = maxColumnWidthEl ? parseInt(maxColumnWidthEl.value, 10) : 200;
      
      const enableAlternatingRowsEl = document.getElementById("enable-alternating-rows") as HTMLInputElement;
      const enableAlternatingRows = enableAlternatingRowsEl ? enableAlternatingRowsEl.checked : true;
      
      const verticalAlignmentEl = document.getElementById("vertical-alignment") as HTMLSelectElement;
      const verticalAlignment = verticalAlignmentEl ? verticalAlignmentEl.value : "Top";
      
      addDebugInfo(`formatSpreadsheet: Config - maxColumnWidth: ${maxColumnWidth}, enableAlternatingRows: ${enableAlternatingRows}`);

      addDebugInfo("formatSpreadsheet: Starting header formatting...");
      // Format headers (first row)
      addDebugInfo("formatSpreadsheet: Getting header row...");
      const headerRow = usedRange.getRow(0);
      
      addDebugInfo("formatSpreadsheet: Setting header font bold...");
      headerRow.format.font.bold = true;
      
      addDebugInfo("formatSpreadsheet: Setting header background color...");
      headerRow.format.fill.color = headerBgColor;
      
      addDebugInfo("formatSpreadsheet: Setting header text color...");
      headerRow.format.font.color = headerTextColor;

      addDebugInfo("formatSpreadsheet: Syncing header formatting...");
      await context.sync();

      addDebugInfo("formatSpreadsheet: Starting border formatting...");
      // Apply borders to the entire range
      addDebugInfo("formatSpreadsheet: Setting InsideHorizontal border...");
      usedRange.format.borders.getItem("InsideHorizontal").style = "Continuous";
      
      addDebugInfo("formatSpreadsheet: Setting InsideHorizontal border color...");
      usedRange.format.borders.getItem("InsideHorizontal").color = borderColor;
      
      addDebugInfo("formatSpreadsheet: Setting InsideVertical border...");
      usedRange.format.borders.getItem("InsideVertical").style = "Continuous";
      usedRange.format.borders.getItem("InsideVertical").color = borderColor;
      
      addDebugInfo("formatSpreadsheet: Setting remaining borders...");
      usedRange.format.borders.getItem("EdgeTop").style = "Continuous";
      usedRange.format.borders.getItem("EdgeTop").color = borderColor;
      usedRange.format.borders.getItem("EdgeBottom").style = "Continuous";
      usedRange.format.borders.getItem("EdgeBottom").color = borderColor;
      usedRange.format.borders.getItem("EdgeLeft").style = "Continuous";
      usedRange.format.borders.getItem("EdgeLeft").color = borderColor;
      usedRange.format.borders.getItem("EdgeRight").style = "Continuous";
      usedRange.format.borders.getItem("EdgeRight").color = borderColor;
      
      addDebugInfo("formatSpreadsheet: Syncing after border formatting...");
      await context.sync();

      // Apply alternating row colors (skip header row)
      if (enableAlternatingRows) {
        for (let i = 1; i < usedRange.rowCount; i++) {
          const row = usedRange.getRow(i);
          if (i % 2 === 1) {
            row.format.fill.color = altRowColor1;
          } else {
            row.format.fill.color = altRowColor2;
          }
        }
      }

      // Set vertical alignment for all cells
      usedRange.format.verticalAlignment = verticalAlignment;

      // Auto-fit columns (temporarily disabled columnWidth constraint to debug)
      addDebugInfo("formatSpreadsheet: Auto-fitting columns...");
      for (let col = 0; col < usedRange.columnCount; col++) {
        const column = usedRange.getColumn(col);
        column.format.autofitColumns();
      }
      
      addDebugInfo("formatSpreadsheet: Syncing after auto-fit...");
      await context.sync();
      
      addDebugInfo("formatSpreadsheet: Skipping columnWidth constraint for debugging...");
    });
  } catch (error) {
    console.error("Error in formatSpreadsheet:", error);
    throw error;
  }
}

export async function addColumns() {
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = worksheet.getUsedRange();
      usedRange.load(["rowCount", "columnCount", "values"]);

      await context.sync();

      if (!usedRange) {
        showMessage("No data found in the worksheet.", "error");
        return;
      }

      const values = usedRange.values;
      const headers = values[0] as string[];

      // Get user preferences
      const columnHeader = (document.getElementById("column-header") as HTMLInputElement).value;
      const columnPosition = (document.getElementById("column-position") as HTMLSelectElement)
        .value;
      const prepopulateCharge = (document.getElementById("prepopulate-charge") as HTMLInputElement)
        .checked;
      const noChargeKeywords = (document.getElementById("no-charge-keywords") as HTMLInputElement)
        .value;
      const addAmendedNarrative = (
        document.getElementById("add-amended-narrative") as HTMLInputElement
      ).checked;
      const addAmendedTime = (document.getElementById("add-amended-time") as HTMLInputElement)
        .checked;
      const addNotesColumn = (document.getElementById("add-notes-column") as HTMLInputElement)
        .checked;

      let insertIndex: number;
      if (columnPosition === "beginning") {
        insertIndex = 0;
      } else {
        insertIndex = headers.length;
      }

      let columnsToAdd: string[] = [];
      let newColumnCount = 0;

      // Build list of columns to add
      if (addAmendedNarrative) {
        columnsToAdd.push("Amended Narrative");
        newColumnCount++;
      }
      if (addAmendedTime) {
        columnsToAdd.push("Amended Time");
        newColumnCount++;
      }
      if (columnHeader && columnHeader.trim()) {
        columnsToAdd.push(columnHeader.trim());
        newColumnCount++;
      }
      if (addNotesColumn) {
        columnsToAdd.push("Notes");
        newColumnCount++;
      }

      if (newColumnCount === 0) {
        addDebugInfo("No columns were configured to be added.");
        return;
      }

      // Rename original columns if amended columns are being added
      if (addAmendedNarrative || addAmendedTime) {
        // First, check if columns exist and rename them
        if (addAmendedNarrative) {
          const narrativeCol = headers.findIndex((h) => h && h.toLowerCase() === "narrative");
          if (narrativeCol !== -1) {
            const narrativeHeaderCell = usedRange.getCell(0, narrativeCol);
            narrativeHeaderCell.values = [["Original Narrative"]];
            addDebugInfo("Renamed 'Narrative' column to 'Original Narrative'");
          }
        }

        if (addAmendedTime) {
          const timeCol = headers.findIndex((h) => h && h.toLowerCase() === "time");
          if (timeCol !== -1) {
            const timeHeaderCell = usedRange.getCell(0, timeCol);
            timeHeaderCell.values = [["Original Time"]];
            addDebugInfo("Renamed 'Time' column to 'Original Time'");
          }
        }
      }

      // Insert new columns
      const insertRange = worksheet
        .getRange(
          `${getColumnLetter(insertIndex + 1)}:${getColumnLetter(insertIndex + newColumnCount)}`
        )
        .getEntireColumn();
      insertRange.insert(Excel.InsertShiftDirection.right);

      // Add headers for new columns
      for (let i = 0; i < columnsToAdd.length; i++) {
        const headerCell = worksheet.getCell(0, insertIndex + i);
        headerCell.values = [[columnsToAdd[i]]];
        addDebugInfo(`Added column: ${columnsToAdd[i]}`);
      }

      // Prepopulate charge column if requested
      if (prepopulateCharge && columnHeader && columnHeader.trim()) {
        const chargeColumnIndex = insertIndex + columnsToAdd.indexOf(columnHeader.trim());
        if (chargeColumnIndex >= insertIndex) {
          const keywords = noChargeKeywords
            .toLowerCase()
            .split(",")
            .map((k) => k.trim())
            .filter((k) => k);

          // Find narrative column (check for "Original Narrative" first, then "Narrative")
          let narrativeCol = headers.findIndex(
            (h) => h && h.toLowerCase() === "original narrative"
          );
          if (narrativeCol === -1) {
            narrativeCol = headers.findIndex((h) => h && h.toLowerCase() === "narrative");
          }

          if (narrativeCol !== -1) {
            const updatedRange = worksheet.getUsedRange();
            updatedRange.load(["rowCount", "values"]);
            await context.sync();

            const updatedValues = updatedRange.values;

            for (let row = 1; row < updatedValues.length; row++) {
              const narrative = (updatedValues[row][narrativeCol] || "").toString().toLowerCase();
              let chargeValue = "Y"; // Default to billable

              if (keywords.some((keyword) => narrative.includes(keyword))) {
                chargeValue = "N";
              }

              const chargeCell = worksheet.getCell(row, chargeColumnIndex);
              chargeCell.values = [[chargeValue]];
            }
            addDebugInfo(`Prepopulated ${columnHeader} column based on narrative keywords`);
          }
        }
      }

      await context.sync();
    });
  } catch (error) {
    console.error("Error in addColumns:", error);
    throw error;
  }
}

function getColumnLetter(columnNumber: number): string {
  let result = "";
  while (columnNumber > 0) {
    columnNumber--;
    result = String.fromCharCode(65 + (columnNumber % 26)) + result;
    columnNumber = Math.floor(columnNumber / 26);
  }
  return result;
}

export async function colorCodeRows() {
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = worksheet.getUsedRange();
      usedRange.load(["rowCount", "columnCount", "values"]);

      await context.sync();

      if (!usedRange) {
        showMessage("No data found in the worksheet.", "error");
        return;
      }

      const values = usedRange.values;
      const headers = values[0] as string[];

      // Find the charge column
      const chargeCol = headers.findIndex((h) => h && h.toLowerCase() === "charge");

      if (chargeCol === -1) {
        addDebugInfo("Charge column not found. Skipping color coding.");
        return;
      }

      // Color code based on charge values
      for (let row = 1; row < usedRange.rowCount; row++) {
        const chargeValue = values[row][chargeCol];
        const rowRange = usedRange.getRow(row);

        if (chargeValue === "Y") {
          // Green for billable
          rowRange.format.fill.color = "#d4edda";
        } else if (chargeValue === "N") {
          // Red for non-billable
          rowRange.format.fill.color = "#f8d7da";
        } else if (chargeValue === "Q") {
          // Yellow for query
          rowRange.format.fill.color = "#fff3cd";
        }
        // Keep default/alternating colors for other values
      }

      await context.sync();
      addDebugInfo("Applied color coding to rows based on Charge column values");
    });
  } catch (error) {
    console.error("Error in colorCodeRows:", error);
    throw error;
  }
}

// Matter profile management
interface FeeEarner {
  name: string;
  role: string;
  rate: number;
  billing_name: string;
  billing_email: string;
  useAsDefault?: boolean;
}

interface NameStandardisationRule {
  enabled: boolean;
  caseSensitive: boolean;
  allowPartialMatches: boolean;
  useDateMatching: boolean;
  replaceOnlyFirstOccurrence: boolean;
  excludedNames?: string[];
  minPartialMatchLength?: number;
  useNicknameDatabase?: boolean;
  customNicknames?: { [key: string]: string };
}

interface MissingTimeEntriesRule {
  enabled: boolean;
  dateTolerance: number; // days ±0 for exact match
  meetingKeywords: string[]; // words that indicate meetings/calls
  requireExactTimeMatch: boolean; // optional stricter matching
  createMissingEntries: boolean; // auto-create placeholder entries
}

interface TimeFormatRule {
  enabled: boolean;
  outputFormat: "HH:MM" | "XX.YY"; // HH:MM (hours:minutes) or XX.YY (decimal hours)
  roundToSixMinutes: boolean; // round to nearest 6-minute increment
}

interface NeedsDetailRule {
  enabled: boolean;
  minWordCount: number; // minimum number of words required in narrative
}

interface TravelRule {
  enabled: boolean;
  keywords: string[]; // travel-related keywords to detect
  caseSensitive: boolean; // whether keyword matching is case sensitive
  chargeValue: string; // value to set in Charge column (typically "N")
  noteText: string; // note to add to Notes column
}

interface NonChargeableSubcategory {
  enabled: boolean;
  keywords: string[];
}

interface NonChargeableRule {
  enabled: boolean;
  caseSensitive: boolean; // whether keyword matching is case sensitive
  chargeValue: string; // value to set in Charge column (typically "N")
  subcategories: {
    clericalAdmin: NonChargeableSubcategory;
    audit: NonChargeableSubcategory;
    ownError: NonChargeableSubcategory;
    research: NonChargeableSubcategory;
  };
}

interface MaxDailyHoursRule {
  enabled: boolean;
  maxHours: number; // configurable maximum hours allowed per day
  chargeValue: string; // value to set in Charge column (typically "Q")
  noteText: string; // note to add to Notes column
}

interface RulesConfig {
  timeFormat: TimeFormatRule;
  nameStandardisation: NameStandardisationRule;
  missingTimeEntries: MissingTimeEntriesRule;
  needsDetail: NeedsDetailRule;
  travel: TravelRule;
  nonChargeable: NonChargeableRule;
  maxDailyHours: MaxDailyHoursRule;
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
  rich: "richard",
  richie: "richard",
  mike: "michael",
  mick: "michael",
  mickey: "michael",
  mikey: "michael",
  jim: "james",
  jimmy: "james",
  jamie: "james",
  jack: "john",
  johnny: "john",
  jon: "jonathan",
  joe: "joseph",
  joey: "joseph",
  dave: "david",
  davey: "david",
  dan: "daniel",
  danny: "daniel",
  tom: "thomas",
  tommy: "thomas",
  tony: "anthony",
  ant: "anthony",
  chris: "christopher",
  ed: "edward",
  eddie: "edward",
  ted: "edward",
  teddy: "edward",
  steve: "stephen",
  stevie: "stephen",
  matt: "matthew",
  matty: "matthew",
  pete: "peter",
  andy: "andrew",
  drew: "andrew",
  phil: "philip",
  al: "alan",
  alex: "alexander",
  ben: "benjamin",
  benny: "benjamin",
  charlie: "charles",
  chuck: "charles",
  frank: "francis",
  frankie: "francis",
  greg: "gregory",
  harry: "harold",
  hank: "henry",
  len: "leonard",
  leo: "leonard",
  max: "maximilian",
  nick: "nicholas",
  pat: "patrick",
  paddy: "patrick",
  ray: "raymond",
  sam: "samuel",
  sammy: "samuel",
  tim: "timothy",
  vinny: "vincent",
  vince: "vincent",
  walt: "walter",
  wally: "walter",

  // Common female nicknames
  sue: "susan",
  susie: "susan",
  suzy: "susan",
  liz: "elizabeth",
  lizzy: "elizabeth",
  beth: "elizabeth",
  betty: "elizabeth",
  libby: "elizabeth",
  eliza: "elizabeth",
  lisa: "elizabeth",
  kate: "katherine",
  katie: "katherine",
  kathy: "katherine",
  kitty: "katherine",
  kit: "katherine",
  cathy: "catherine",
  cat: "catherine",
  annie: "anne",
  nan: "anne",
  nancy: "anne",
  maggie: "margaret",
  meg: "margaret",
  peggy: "margaret",
  peg: "margaret",
  jen: "jennifer",
  jenny: "jennifer",
  jess: "jessica",
  jessie: "jessica",
  becky: "rebecca",
  becca: "rebecca",
  debbie: "deborah",
  deb: "deborah",
  cindy: "cynthia",
  sandy: "sandra",
  mandy: "amanda",
  amy: "amelia",
  mel: "melissa",
  missy: "melissa",
  steph: "stephanie",
  steffi: "stephanie",
  christie: "christine",
  chris: "christine",
  tina: "christina",
  patty: "patricia",
  pat: "patricia",
  trish: "patricia",
  angie: "angela",
  carol: "caroline",
  carrie: "caroline",
  donna: "madonna",
  diane: "diana",
  fran: "francine",
  ginny: "virginia",
  ginger: "virginia",
  helen: "helena",
  jo: "joanne",
  joanie: "joanne",
  judy: "judith",
  jude: "judith",
  linda: "belinda",
  lynn: "carolyn",
  marie: "mary",
  molly: "mary",
  polly: "mary",
  nina: "antonina",
  penny: "penelope",
  rosie: "rosemary",
  sally: "sarah",
  shelly: "michelle",
  tammy: "tamara",
  terry: "theresa",
  val: "valerie",
  vicky: "victoria",
  wendy: "gwendolyn",
};

interface MatterProfile {
  name: string;
  clientName?: string;
  clientNumber?: string;
  matterNumber?: string;
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
    clientName: (document.getElementById("client-name") as HTMLInputElement)?.value || "",
    clientNumber: (document.getElementById("client-number") as HTMLInputElement)?.value || "",
    matterNumber: (document.getElementById("matter-number") as HTMLInputElement)?.value || "",
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
    feeEarners: getFeeEarnersFromForm(),
    rules: {
      timeFormat: {
        enabled: false,
        outputFormat: "HH:MM",
        roundToSixMinutes: true,
      },
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
      needsDetail: {
        enabled: false,
        minWordCount: 3,
      },
      travel: {
        enabled: false,
        keywords: [
          "travel",
          "travelling",
          "drive",
          "driving",
          "airport",
          "flight",
          "hotel",
          "accommodation",
          "train",
          "taxi",
          "uber",
          "journey",
          "commute",
          "transport",
        ],
        caseSensitive: false,
        chargeValue: "N",
        noteText: "NonBillable - Travel",
      },
      nonChargeable: {
        enabled: false,
        caseSensitive: false,
        chargeValue: "N",
        subcategories: {
          clericalAdmin: {
            enabled: false,
            keywords: [
              "filing",
              "admin",
              "administration",
              "clerical",
              "photocopying",
              "scanning",
              "organizing",
              "office",
              "paperwork",
              "housekeeping",
            ],
          },
          audit: {
            enabled: false,
            keywords: [
              "audit",
              "auditing",
              "compliance",
              "review",
              "checking",
              "verification",
              "quality control",
              "monitoring",
            ],
          },
          ownError: {
            enabled: false,
            keywords: [
              "mistake",
              "error",
              "correction",
              "fix",
              "redo",
              "revise",
              "amend",
              "rectify",
              "wrong",
              "incorrect",
            ],
          },
          research: {
            enabled: false,
            keywords: [
              "research",
              "investigating",
              "learning",
              "studying",
              "training",
              "education",
              "reading",
              "background",
              "familiarization",
            ],
          },
        },
      },
      maxDailyHours: {
        enabled: false,
        maxHours: 10,
        chargeValue: "Q",
        noteText: "Max Daily Hours Exceeded",
      },
    },
  };
}

function loadMatterProfile(profile: MatterProfile) {
  // Load basic settings
  (document.getElementById("client-name") as HTMLInputElement).value = profile.clientName || "";
  (document.getElementById("client-number") as HTMLInputElement).value = profile.clientNumber || "";
  (document.getElementById("matter-number") as HTMLInputElement).value = profile.matterNumber || "";
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
    profile.verticalAlignment;
  (document.getElementById("column-header") as HTMLInputElement).value = profile.columnHeader;
  (document.getElementById("column-position") as HTMLSelectElement).value = profile.columnPosition;
  (document.getElementById("prepopulate-charge") as HTMLInputElement).checked =
    profile.prepopulateCharge;
  (document.getElementById("no-charge-keywords") as HTMLInputElement).value =
    profile.noChargeKeywords;
  (document.getElementById("add-amended-narrative") as HTMLInputElement).checked =
    profile.addAmendedNarrative;
  (document.getElementById("add-amended-time") as HTMLInputElement).checked =
    profile.addAmendedTime;
  (document.getElementById("add-notes-column") as HTMLInputElement).checked =
    profile.addNotesColumn;

  // Load fee earners
  loadFeeEarnersIntoForm(profile.feeEarners);

  // Load rules if they exist
  if (profile.rules) {
    loadRulesConfig(profile.rules);
  }
}

function saveMatterProfile() {
  // Check both possible input fields for matter name
  const matterNameInput = document.getElementById("matter-name") as HTMLInputElement;
  const newMatterNameInput = document.getElementById("new-matter-name") as HTMLInputElement;
  
  let name = "";
  if (newMatterNameInput && newMatterNameInput.value.trim()) {
    name = newMatterNameInput.value.trim();
  } else if (matterNameInput && matterNameInput.value.trim()) {
    name = matterNameInput.value.trim();
  }

  if (!name) {
    showMessage("Please enter a matter name.", "error");
    return;
  }

  const profiles = getMatterProfiles();
  const existingIndex = profiles.findIndex((p) => p.name === name);

  const newProfile = getCurrentSettings();
  newProfile.name = name;

  if (existingIndex >= 0) {
    profiles[existingIndex] = newProfile;
    showMessage(`Matter profile "${name}" updated successfully.`, "success");
  } else {
    profiles.push(newProfile);
    showMessage(`Matter profile "${name}" saved successfully.`, "success");
  }

  saveMatterProfiles(profiles);
  updateMatterDropdown();

  // Clear both matter name inputs
  if (matterNameInput) {
    matterNameInput.value = "";
  }
  if (newMatterNameInput) {
    newMatterNameInput.value = "";
    
    // Hide the new matter section after successful save
    const newMatterSection = document.getElementById("new-matter-section") as HTMLElement;
    if (newMatterSection) {
      newMatterSection.style.display = "none";
    }
  }
  
  // Set the newly created matter as selected in the dropdown
  const matterSelect = document.getElementById("matter-select") as HTMLSelectElement;
  if (matterSelect) {
    matterSelect.value = name;
    currentMatterLoaded = name;
    updateUIForMatterState();
  }
}

function deleteMatterProfile() {
  const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;

  if (!selectedMatter) {
    showMessage("Please select a matter profile to delete.", "error");
    return;
  }

  const profiles = getMatterProfiles();
  const filteredProfiles = profiles.filter((p) => p.name !== selectedMatter);

  if (filteredProfiles.length === profiles.length) {
    showMessage("Matter profile not found.", "error");
    return;
  }

  saveMatterProfiles(filteredProfiles);
  updateMatterDropdown();
  showMessage(`Matter profile "${selectedMatter}" deleted successfully.`, "success");

  // Reset UI state
  currentMatterLoaded = null;
  updateUIForMatterState();
}

function saveCurrentSettings() {
  const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;

  if (!selectedMatter) {
    showMessage("Please select a matter profile to save settings to.", "error");
    return;
  }

  const profiles = getMatterProfiles();
  const existingIndex = profiles.findIndex((p) => p.name === selectedMatter);

  if (existingIndex >= 0) {
    const updatedProfile = getCurrentSettings();
    updatedProfile.name = selectedMatter;
    profiles[existingIndex] = updatedProfile;
    saveMatterProfiles(profiles);
    showMessage(`Settings saved to matter profile "${selectedMatter}".`, "success");
  } else {
    showMessage("Selected matter profile not found.", "error");
  }
}

function saveMatterProfileFromDropdown() {
  const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;

  if (!selectedMatter) {
    showMessage("Please select a matter profile first.", "error");
    return;
  }

  // This is the same as saveCurrentSettings, just different messaging
  saveCurrentSettings();
}

function getMatterProfiles(): MatterProfile[] {
  const stored = localStorage.getItem("fixmytime-matter-profiles");
  if (stored) {
    try {
      return JSON.parse(stored);
    } catch (e) {
      console.error("Error parsing stored matter profiles:", e);
    }
  }
  return [];
}

function saveMatterProfiles(profiles: MatterProfile[]) {
  try {
    localStorage.setItem("fixmytime-matter-profiles", JSON.stringify(profiles));
    addDebugInfo(`Saved ${profiles.length} matter profiles to localStorage`);
  } catch (e) {
    console.error("Error saving matter profiles:", e);
    showMessage("Error saving matter profiles: " + e.message, "error");
  }
}

function updateMatterDropdown() {
  const dropdown = document.getElementById("matter-select") as HTMLSelectElement;
  const profiles = getMatterProfiles();

  // Clear existing options except the first one
  dropdown.innerHTML = '<option value="">Select a matter profile...</option>';

  profiles.forEach((profile) => {
    const option = document.createElement("option");
    option.value = profile.name;
    option.textContent = profile.name;
    dropdown.appendChild(option);
  });

  // Add the "Add New Matter" option
  const addNewOption = document.createElement("option");
  addNewOption.value = "__new__";
  addNewOption.textContent = "+ Add New Matter";
  addNewOption.style.fontStyle = "italic";
  dropdown.appendChild(addNewOption);

  addDebugInfo(`Updated matter dropdown with ${profiles.length} profiles`);
}

// Fee Earners functionality
function getFeeEarnersFromForm(): FeeEarner[] {
  const tbody = document.getElementById("fee-earners-tbody");
  if (!tbody) return [];

  const rows = tbody.querySelectorAll("tr");
  const feeEarners: FeeEarner[] = [];

  rows.forEach((row) => {
    const nameInput = row.querySelector(".name-input") as HTMLInputElement;
    const roleInput = row.querySelector(".role-input") as HTMLInputElement;
    const rateInput = row.querySelector(".rate-input") as HTMLInputElement;
    const billingNameInput = row.querySelector(".billing-name-input") as HTMLInputElement;
    const billingEmailInput = row.querySelector(".billing-email-input") as HTMLInputElement;
    const useAsDefaultCheckbox = row.querySelector(".use-as-default-checkbox") as HTMLInputElement;

    if (nameInput && nameInput.value.trim()) {
      feeEarners.push({
        name: nameInput.value.trim(),
        role: roleInput?.value.trim() || "",
        rate: parseFloat(rateInput?.value) || 0,
        billing_name: billingNameInput?.value.trim() || "",
        billing_email: billingEmailInput?.value.trim() || "",
        useAsDefault: useAsDefaultCheckbox?.checked || false,
      });
    }
  });

  return feeEarners;
}

function loadFeeEarnersIntoForm(feeEarners: FeeEarner[]) {
  const tbody = document.getElementById("fee-earners-tbody");
  if (!tbody) return;

  // Clear existing rows
  tbody.innerHTML = "";

  // Add fee earners
  feeEarners.forEach((feeEarner) => {
    addFeeEarnerRow(feeEarner);
  });

  // Refresh duplicate detection
  detectDuplicateNames();
}

function addFeeEarnerRow(feeEarner?: FeeEarner) {
  const tbody = document.getElementById("fee-earners-tbody");
  if (!tbody) return;

  const row = document.createElement("tr");
  row.innerHTML = `
    <td><input type="text" class="name-input" value="${feeEarner?.name || ""}" placeholder="Enter name" oninput="detectDuplicateNames()"></td>
    <td><input type="text" class="role-input" value="${feeEarner?.role || ""}" placeholder="Enter role" onchange="updateBillingFields(this)"></td>
    <td><input type="number" class="rate-input" value="${feeEarner?.rate || ""}" placeholder="0.00" step="0.01" min="0"></td>
    <td class="billing-name-cell ${!feeEarner?.role ? "disabled-field" : ""}">
      <input type="text" class="billing-name-input" value="${feeEarner?.billing_name || ""}" placeholder="Auto-generated" ${!feeEarner?.role ? "disabled" : ""}>
    </td>
    <td class="billing-email-cell ${!feeEarner?.role ? "disabled-field" : ""}">
      <input type="text" class="billing-email-input" value="${feeEarner?.billing_email || ""}" placeholder="Auto-generated" ${!feeEarner?.role ? "disabled" : ""}>
    </td>
    <td>
      <input type="checkbox" class="use-as-default-checkbox" ${feeEarner?.useAsDefault ? "checked" : ""} style="display: none;" onchange="handleDefaultCheckboxChange(event, this.closest('tr'))">
    </td>
    <td><button type="button" onclick="removeFeeEarnerRow(this)">Remove</button></td>
  `;

  tbody.appendChild(row);
  detectDuplicateNames();
}

function removeFeeEarnerRow(button: HTMLButtonElement) {
  const row = button.closest("tr");
  if (row) {
    row.remove();
    detectDuplicateNames(); // Refresh after removal
  }
}

function updateBillingFields(roleInput: HTMLInputElement) {
  const row = roleInput.closest("tr");
  if (!row) return;

  const nameInput = row.querySelector(".name-input") as HTMLInputElement;
  const billingNameInput = row.querySelector(".billing-name-input") as HTMLInputElement;
  const billingEmailInput = row.querySelector(".billing-email-input") as HTMLInputElement;
  const billingNameCell = row.querySelector(".billing-name-cell");
  const billingEmailCell = row.querySelector(".billing-email-cell");

  const role = roleInput.value.trim();
  const name = nameInput.value.trim();

  if (role && name) {
    // Enable billing fields
    billingNameCell.classList.remove("disabled-field");
    billingEmailCell.classList.remove("disabled-field");
    billingNameInput.disabled = false;
    billingEmailInput.disabled = false;

    // Auto-generate billing name and email if empty
    if (!billingNameInput.value.trim()) {
      billingNameInput.value = `${name} (${role})`;
    }
    if (!billingEmailInput.value.trim()) {
      const emailName = name.toLowerCase().replace(/\s+/g, ".");
      const roleShort = role.toLowerCase().replace(/\s+/g, "");
      billingEmailInput.value = `${emailName}.${roleShort}@firm.com`;
    }
  } else {
    // Disable billing fields
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
        const otherNameInput = row.querySelector(".name-input") as HTMLInputElement;
        const otherCheckbox = row.querySelector(".use-as-default-checkbox") as HTMLInputElement;

        if (otherNameInput && otherNameInput.value.trim()) {
          const otherFirstName = otherNameInput.value.trim().split(" ")[0].toLowerCase();
          if (otherFirstName === firstName) {
            otherCheckbox.checked = false;
          }
        }
      }
    });
  }
}

function updateFeeEarnersFromSpreadsheet() {
  Excel.run(async (context) => {
    try {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = worksheet.getUsedRange();
      usedRange.load(["values"]);

      await context.sync();

      if (!usedRange) {
        showMessage("No data found in the worksheet.", "error");
        return;
      }

      const values = usedRange.values;
      const headers = values[0] as string[];

      // Find the name column
      const nameCol = headers.findIndex((h) => h && h.toLowerCase() === "name");

      if (nameCol === -1) {
        showMessage("Name column not found in the spreadsheet.", "error");
        return;
      }

      // Extract unique names
      const uniqueNames = new Set<string>();
      for (let i = 1; i < values.length; i++) {
        const name = values[i][nameCol];
        if (name && typeof name === "string" && name.trim()) {
          uniqueNames.add(name.trim());
        }
      }

      // Get current fee earners to preserve existing data
      const currentFeeEarners = getFeeEarnersFromForm();
      const currentFeeEarnerMap = new Map<string, FeeEarner>();
      currentFeeEarners.forEach((fe) => {
        currentFeeEarnerMap.set(fe.name.toLowerCase(), fe);
      });

      // Build updated fee earner list
      const updatedFeeEarners: FeeEarner[] = [];
      uniqueNames.forEach((name) => {
        const existing = currentFeeEarnerMap.get(name.toLowerCase());
        if (existing) {
          // Keep existing data
          updatedFeeEarners.push(existing);
        } else {
          // Add new fee earner with default values
          updatedFeeEarners.push({
            name: name,
            role: "",
            rate: 0,
            billing_name: "",
            billing_email: "",
            useAsDefault: false,
          });
        }
      });

      // Update the form
      loadFeeEarnersIntoForm(updatedFeeEarners);
      resetTableScroll();

      showMessage(
        `Updated fee earners list. Found ${uniqueNames.size} unique names in the spreadsheet.`,
        "success"
      );
      addDebugInfo(`Extracted ${uniqueNames.size} unique names from Name column`);
    } catch (error) {
      console.error("Error updating fee earners from spreadsheet:", error);
      showMessage("Error reading from spreadsheet: " + error.message, "error");
    }
  });
}

function saveParticipants() {
  const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;

  if (!selectedMatter) {
    showMessage("Please select a matter profile to save participants to.", "error");
    return;
  }

  const profiles = getMatterProfiles();
  const existingIndex = profiles.findIndex((p) => p.name === selectedMatter);

  if (existingIndex >= 0) {
    // Update the existing profile with new fee earners data
    profiles[existingIndex].feeEarners = getFeeEarnersFromForm();
    saveMatterProfiles(profiles);

    const feeEarnerCount = profiles[existingIndex].feeEarners.length;
    showMessage(
      `Successfully saved ${feeEarnerCount} fee earner${feeEarnerCount !== 1 ? "s" : ""} to matter profile "${selectedMatter}".`,
      "success"
    );
  } else {
    showMessage("Selected matter profile not found. Please create a new profile first.", "error");
  }
}

// Rules functionality

let missingTimeRowMapping: Map<string, number> = new Map();

function clearMissingTimeRowTracking() {
  missingTimeRowMapping.clear();
}

async function applyAllRules() {
  try {
    // Clear any previous Missing Time row tracking
    clearMissingTimeRowTracking();

    // Clear debug info before starting
    clearDebugInfo();

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

    const appliedRules: string[] = [];
    let totalUpdatedRows = 0;

    // Apply TimeFormat Rule if enabled
    if (currentProfile.rules.timeFormat?.enabled) {
      showMessage("Applying TimeFormat rule...", "info");
      const result = await applyTimeFormatRuleWithResult();
      if (result.success) {
        appliedRules.push("TimeFormat");
        totalUpdatedRows += result.updatedRows;
      } else if (result.error) {
        showMessage(`TimeFormat failed: ${result.error}`, "error");
        return;
      }
    }

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
      } else if (result.error) {
        showMessage(`Missing Time Entries failed: ${result.error}`, "error");
        return;
      }
    }

    // Apply NeedsDetail Rule if enabled
    if (currentProfile.rules.needsDetail?.enabled) {
      showMessage("Applying NeedsDetail rule...", "info");
      const result = await applyNeedsDetailRuleWithResult();
      if (result.success) {
        appliedRules.push("NeedsDetail");
        totalUpdatedRows += result.updatedRows;
      } else if (result.error) {
        showMessage(`NeedsDetail failed: ${result.error}`, "error");
        return;
      }
    }

    // Apply Travel Rule if enabled
    if (currentProfile.rules.travel?.enabled) {
      showMessage("Applying Travel rule...", "info");
      const result = await applyTravelRuleWithResult();
      if (result.success) {
        appliedRules.push("Travel");
        totalUpdatedRows += result.updatedRows;
      } else if (result.error) {
        showMessage(`Travel failed: ${result.error}`, "error");
        return;
      }
    }

    // Apply Max Daily Hours Rule if enabled
    if (currentProfile.rules.maxDailyHours?.enabled) {
      showMessage("Applying Max Daily Hours rule...", "info");
      const result = await applyMaxDailyHoursRuleWithResult();
      if (result.success) {
        appliedRules.push("Max Daily Hours");
        totalUpdatedRows += result.updatedRows;
      } else if (result.error) {
        showMessage(`Max Daily Hours failed: ${result.error}`, "error");
        return;
      }
    }

    // Apply Non Chargeable Rule if enabled
    if (currentProfile.rules.nonChargeable?.enabled) {
      showMessage("Applying Non Chargeable rule...", "info");
      const result = await applyNonChargeableRuleWithResult();
      if (result.success) {
        appliedRules.push("Non Chargeable");
        totalUpdatedRows += result.updatedRows;
      } else if (result.error) {
        showMessage(`Non Chargeable failed: ${result.error}`, "error");
        return;
      }
    }

    // Show final result and re-apply formatting
    if (appliedRules.length > 0) {
      showMessage(
        `Successfully applied ${appliedRules.length} rule${appliedRules.length !== 1 ? "s" : ""}: ${appliedRules.join(", ")}. Updated ${totalUpdatedRows} row${totalUpdatedRows !== 1 ? "s" : ""}.`,
        "success"
      );

      // Re-apply formatting after rule changes
      await formatSpreadsheet();
      await colorCodeRows();
    } else {
      showMessage("No rules were enabled or applied.", "info");
    }
  } catch (error) {
    console.error("Error applying rules:", error);
    showMessage(`Error applying rules: ${error.message}`, "error");
  }
}

function getDefaultRules(): RulesConfig {
  return {
    timeFormat: {
      enabled: false,
      outputFormat: "HH:MM",
      roundToSixMinutes: true,
    },
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
    needsDetail: {
      enabled: false,
      minWordCount: 3,
    },
    travel: {
      enabled: false,
      keywords: [
        "travel",
        "travelling",
        "drive",
        "driving",
        "airport",
        "flight",
        "hotel",
        "accommodation",
        "train",
        "taxi",
        "uber",
        "journey",
        "commute",
        "transport",
      ],
      caseSensitive: false,
      chargeValue: "N",
      noteText: "NonBillable - Travel",
    },
    nonChargeable: {
      enabled: false,
      caseSensitive: false,
      chargeValue: "N",
      subcategories: {
        clericalAdmin: {
          enabled: false,
          keywords: [
            "filing",
            "admin",
            "administration",
            "clerical",
            "photocopying",
            "scanning",
            "organizing",
            "office",
            "paperwork",
            "housekeeping",
          ],
        },
        audit: {
          enabled: false,
          keywords: [
            "audit",
            "auditing",
            "compliance",
            "review",
            "checking",
            "verification",
            "quality control",
            "monitoring",
          ],
        },
        ownError: {
          enabled: false,
          keywords: [
            "mistake",
            "error",
            "correction",
            "fix",
            "redo",
            "revise",
            "amend",
            "rectify",
            "wrong",
            "incorrect",
          ],
        },
        research: {
          enabled: false,
          keywords: [
            "research",
            "investigating",
            "learning",
            "studying",
            "training",
            "education",
            "reading",
            "background",
            "familiarization",
          ],
        },
      },
    },
    maxDailyHours: {
      enabled: false,
      maxHours: 10,
      chargeValue: "Q",
      noteText: "Max Daily Hours Exceeded",
    },
  };
}

function getCurrentRules(): RulesConfig {
  const excludedNamesText = (document.getElementById("excluded-names") as HTMLInputElement).value;
  const excludedNames = excludedNamesText
    ? excludedNamesText
        .split(",")
        .map((name) => name.trim())
        .filter((name) => name.length > 0)
    : [];

  const customNicknames = getNicknameDatabaseFromForm();

  return {
    timeFormat: {
      enabled:
        (document.getElementById("time-format-enabled") as HTMLInputElement)?.checked || false,
      outputFormat:
        ((document.getElementById("time-format-output") as HTMLSelectElement)?.value as
          | "HH:MM"
          | "XX.YY") || "HH:MM",
      roundToSixMinutes:
        (document.getElementById("time-format-round") as HTMLInputElement)?.checked || false,
    },
    nameStandardisation: {
      enabled: (document.getElementById("name-standardisation-enabled") as HTMLInputElement)
        .checked,
      caseSensitive: false, // Always false for now
      allowPartialMatches: (document.getElementById("partial-matches") as HTMLInputElement).checked,
      useDateMatching: (document.getElementById("date-matching") as HTMLInputElement).checked,
      replaceOnlyFirstOccurrence: (
        document.getElementById("first-occurrence-only") as HTMLInputElement
      ).checked,
      excludedNames: excludedNames,
      minPartialMatchLength: parseInt(
        (document.getElementById("min-partial-match-length") as HTMLInputElement).value || "3"
      ),
      useNicknameDatabase: (document.getElementById("use-nickname-database") as HTMLInputElement)
        .checked,
      customNicknames: customNicknames,
    },
    missingTimeEntries: {
      enabled:
        (document.getElementById("missing-time-entries-enabled") as HTMLInputElement)?.checked ||
        false,
      dateTolerance: parseInt(
        (document.getElementById("date-tolerance") as HTMLInputElement)?.value || "0"
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
    needsDetail: {
      enabled:
        (document.getElementById("needs-detail-enabled") as HTMLInputElement)?.checked || false,
      minWordCount: parseInt(
        (document.getElementById("min-word-count") as HTMLInputElement)?.value || "3"
      ),
    },
    travel: {
      enabled: (document.getElementById("travel-enabled") as HTMLInputElement)?.checked || false,
      keywords: (
        (document.getElementById("travel-keywords") as HTMLInputElement)?.value ||
        "travel,travelling,drive,driving,airport,flight,hotel,accommodation,train,taxi,uber,journey,commute,transport"
      )
        .split(",")
        .map((keyword) => keyword.trim())
        .filter((keyword) => keyword.length > 0),
      caseSensitive:
        (document.getElementById("travel-case-sensitive") as HTMLInputElement)?.checked || false,
      chargeValue:
        (document.getElementById("travel-charge-value") as HTMLInputElement)?.value || "N",
      noteText:
        (document.getElementById("travel-note-text") as HTMLInputElement)?.value ||
        "NonBillable - Travel",
    },
    nonChargeable: {
      enabled:
        (document.getElementById("non-chargeable-enabled") as HTMLInputElement)?.checked || false,
      caseSensitive:
        (document.getElementById("non-chargeable-case-sensitive") as HTMLInputElement)?.checked ||
        false,
      chargeValue:
        (document.getElementById("non-chargeable-charge-value") as HTMLInputElement)?.value || "N",
      subcategories: {
        clericalAdmin: {
          enabled:
            (document.getElementById("clerical-admin-enabled") as HTMLInputElement)?.checked ||
            false,
          keywords: (
            (document.getElementById("clerical-admin-keywords") as HTMLInputElement)?.value ||
            "filing,admin,administration,clerical,photocopying,scanning,organizing,office,paperwork,housekeeping"
          )
            .split(",")
            .map((keyword) => keyword.trim())
            .filter((keyword) => keyword.length > 0),
        },
        audit: {
          enabled: (document.getElementById("audit-enabled") as HTMLInputElement)?.checked || false,
          keywords: (
            (document.getElementById("audit-keywords") as HTMLInputElement)?.value ||
            "audit,auditing,compliance,review,checking,verification,quality control,monitoring"
          )
            .split(",")
            .map((keyword) => keyword.trim())
            .filter((keyword) => keyword.length > 0),
        },
        ownError: {
          enabled:
            (document.getElementById("own-error-enabled") as HTMLInputElement)?.checked || false,
          keywords: (
            (document.getElementById("own-error-keywords") as HTMLInputElement)?.value ||
            "mistake,error,correction,fix,redo,revise,amend,rectify,wrong,incorrect"
          )
            .split(",")
            .map((keyword) => keyword.trim())
            .filter((keyword) => keyword.length > 0),
        },
        research: {
          enabled:
            (document.getElementById("research-enabled") as HTMLInputElement)?.checked || false,
          keywords: (
            (document.getElementById("research-keywords") as HTMLInputElement)?.value ||
            "research,investigating,learning,studying,training,education,reading,background,familiarization"
          )
            .split(",")
            .map((keyword) => keyword.trim())
            .filter((keyword) => keyword.length > 0),
        },
      },
    },
    maxDailyHours: {
      enabled:
        (document.getElementById("max-daily-hours-enabled") as HTMLInputElement)?.checked || false,
      maxHours: parseInt(
        (document.getElementById("max-hours-limit") as HTMLInputElement)?.value || "10"
      ),
      chargeValue:
        (document.getElementById("max-daily-hours-charge") as HTMLInputElement)?.value || "Q",
      noteText:
        (document.getElementById("max-daily-hours-note") as HTMLInputElement)?.value ||
        "Max Daily Hours Exceeded",
    },
  };
}

function loadRulesConfig(rules: RulesConfig) {
  // Load TimeFormat rule settings (with null checks for backward compatibility)
  const timeFormatRule = rules.timeFormat || getDefaultRules().timeFormat;

  const timeFormatEnabledEl = document.getElementById("time-format-enabled") as HTMLInputElement;
  if (timeFormatEnabledEl) {
    timeFormatEnabledEl.checked = timeFormatRule.enabled;
  }

  const timeFormatOutputEl = document.getElementById("time-format-output") as HTMLSelectElement;
  if (timeFormatOutputEl) {
    timeFormatOutputEl.value = timeFormatRule.outputFormat;
  }

  const timeFormatRoundEl = document.getElementById("time-format-round") as HTMLInputElement;
  if (timeFormatRoundEl) {
    timeFormatRoundEl.checked = timeFormatRule.roundToSixMinutes;
  }

  // Show/hide TimeFormat configuration based on enabled state
  const timeFormatConfigDiv = document.getElementById("time-format-content");
  if (timeFormatConfigDiv) {
    timeFormatConfigDiv.style.display = timeFormatRule.enabled ? "block" : "none";
  }

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

  // Load NeedsDetail rule settings (with null checks for backward compatibility)
  const needsDetailRule = rules.needsDetail || getDefaultRules().needsDetail;

  const needsDetailEnabledEl = document.getElementById("needs-detail-enabled") as HTMLInputElement;
  if (needsDetailEnabledEl) {
    needsDetailEnabledEl.checked = needsDetailRule.enabled;
  }

  const minWordCountEl = document.getElementById("min-word-count") as HTMLInputElement;
  if (minWordCountEl) {
    minWordCountEl.value = needsDetailRule.minWordCount.toString();
  }

  // Show/hide NeedsDetail configuration based on enabled state
  const needsDetailConfigDiv = document.getElementById("needs-detail-content");
  if (needsDetailConfigDiv) {
    needsDetailConfigDiv.style.display = needsDetailRule.enabled ? "block" : "none";
  }

  // Load Travel rule settings (with null checks for backward compatibility)
  const travelRule = rules.travel || getDefaultRules().travel;

  const travelEnabledEl = document.getElementById("travel-enabled") as HTMLInputElement;
  if (travelEnabledEl) {
    travelEnabledEl.checked = travelRule.enabled;
  }

  const travelKeywordsEl = document.getElementById("travel-keywords") as HTMLInputElement;
  if (travelKeywordsEl) {
    travelKeywordsEl.value = travelRule.keywords.join(", ");
  }

  const travelCaseSensitiveEl = document.getElementById(
    "travel-case-sensitive"
  ) as HTMLInputElement;
  if (travelCaseSensitiveEl) {
    travelCaseSensitiveEl.checked = travelRule.caseSensitive;
  }

  const travelChargeValueEl = document.getElementById("travel-charge-value") as HTMLInputElement;
  if (travelChargeValueEl) {
    travelChargeValueEl.value = travelRule.chargeValue;
  }

  const travelNoteTextEl = document.getElementById("travel-note-text") as HTMLInputElement;
  if (travelNoteTextEl) {
    travelNoteTextEl.value = travelRule.noteText;
  }

  // Show/hide Travel configuration based on enabled state
  const travelConfigDiv = document.getElementById("travel-content");
  if (travelConfigDiv) {
    travelConfigDiv.style.display = travelRule.enabled ? "block" : "none";
  }

  // Load Non Chargeable rule settings (with null checks for backward compatibility)
  const nonChargeableRule = rules.nonChargeable || getDefaultRules().nonChargeable;

  const nonChargeableEnabledEl = document.getElementById(
    "non-chargeable-enabled"
  ) as HTMLInputElement;
  if (nonChargeableEnabledEl) {
    nonChargeableEnabledEl.checked = nonChargeableRule.enabled;
  }

  const nonChargeableCaseSensitiveEl = document.getElementById(
    "non-chargeable-case-sensitive"
  ) as HTMLInputElement;
  if (nonChargeableCaseSensitiveEl) {
    nonChargeableCaseSensitiveEl.checked = nonChargeableRule.caseSensitive;
  }

  const nonChargeableChargeValueEl = document.getElementById(
    "non-chargeable-charge-value"
  ) as HTMLInputElement;
  if (nonChargeableChargeValueEl) {
    nonChargeableChargeValueEl.value = nonChargeableRule.chargeValue;
  }

  // Load subcategory settings
  const clericalAdminEnabledEl = document.getElementById(
    "clerical-admin-enabled"
  ) as HTMLInputElement;
  if (clericalAdminEnabledEl) {
    clericalAdminEnabledEl.checked = nonChargeableRule.subcategories.clericalAdmin.enabled;
  }
  const clericalAdminKeywordsEl = document.getElementById(
    "clerical-admin-keywords"
  ) as HTMLInputElement;
  if (clericalAdminKeywordsEl) {
    clericalAdminKeywordsEl.value =
      nonChargeableRule.subcategories.clericalAdmin.keywords.join(", ");
  }

  const auditEnabledEl = document.getElementById("audit-enabled") as HTMLInputElement;
  if (auditEnabledEl) {
    auditEnabledEl.checked = nonChargeableRule.subcategories.audit.enabled;
  }
  const auditKeywordsEl = document.getElementById("audit-keywords") as HTMLInputElement;
  if (auditKeywordsEl) {
    auditKeywordsEl.value = nonChargeableRule.subcategories.audit.keywords.join(", ");
  }

  const ownErrorEnabledEl = document.getElementById("own-error-enabled") as HTMLInputElement;
  if (ownErrorEnabledEl) {
    ownErrorEnabledEl.checked = nonChargeableRule.subcategories.ownError.enabled;
  }
  const ownErrorKeywordsEl = document.getElementById("own-error-keywords") as HTMLInputElement;
  if (ownErrorKeywordsEl) {
    ownErrorKeywordsEl.value = nonChargeableRule.subcategories.ownError.keywords.join(", ");
  }

  const researchEnabledEl = document.getElementById("research-enabled") as HTMLInputElement;
  if (researchEnabledEl) {
    researchEnabledEl.checked = nonChargeableRule.subcategories.research.enabled;
  }
  const researchKeywordsEl = document.getElementById("research-keywords") as HTMLInputElement;
  if (researchKeywordsEl) {
    researchKeywordsEl.value = nonChargeableRule.subcategories.research.keywords.join(", ");
  }

  // Show/hide Non Chargeable configuration based on enabled state
  const nonChargeableConfigDiv = document.getElementById("non-chargeable-content");
  if (nonChargeableConfigDiv) {
    nonChargeableConfigDiv.style.display = nonChargeableRule.enabled ? "block" : "none";
  }

  // Load Max Daily Hours rule settings (with null checks for backward compatibility)
  const maxDailyHoursRule = rules.maxDailyHours || getDefaultRules().maxDailyHours;

  const maxDailyHoursEnabledEl = document.getElementById(
    "max-daily-hours-enabled"
  ) as HTMLInputElement;
  if (maxDailyHoursEnabledEl) {
    maxDailyHoursEnabledEl.checked = maxDailyHoursRule.enabled;
  }

  const maxHoursLimitEl = document.getElementById("max-hours-limit") as HTMLInputElement;
  if (maxHoursLimitEl) {
    maxHoursLimitEl.value = maxDailyHoursRule.maxHours.toString();
  }

  const maxDailyHoursChargeEl = document.getElementById(
    "max-daily-hours-charge"
  ) as HTMLInputElement;
  if (maxDailyHoursChargeEl) {
    maxDailyHoursChargeEl.value = maxDailyHoursRule.chargeValue;
  }

  const maxDailyHoursNoteEl = document.getElementById("max-daily-hours-note") as HTMLInputElement;
  if (maxDailyHoursNoteEl) {
    maxDailyHoursNoteEl.value = maxDailyHoursRule.noteText;
  }

  // Show/hide Max Daily Hours configuration based on enabled state
  const maxDailyHoursConfigDiv = document.getElementById("max-daily-hours-content");
  if (maxDailyHoursConfigDiv) {
    maxDailyHoursConfigDiv.style.display = maxDailyHoursRule.enabled ? "block" : "none";
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
    if (currentRules.timeFormat?.enabled) {
      enabledRules.push("TimeFormat");
    }
    if (currentRules.nameStandardisation?.enabled) {
      enabledRules.push("Name Standardisation");
    }
    if (currentRules.missingTimeEntries?.enabled) {
      enabledRules.push("Missing Time Entries");
    }
    if (currentRules.needsDetail?.enabled) {
      enabledRules.push("NeedsDetail");
    }
    if (currentRules.travel?.enabled) {
      enabledRules.push("Travel");
    }
    if (currentRules.nonChargeable?.enabled) {
      enabledRules.push("Non Chargeable");
    }
    if (currentRules.maxDailyHours?.enabled) {
      enabledRules.push("Max Daily Hours");
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

// Max Daily Hours Rule implementation
async function applyMaxDailyHoursRuleWithResult(): Promise<{
  success: boolean;
  updatedRows: number;
  error?: string;
}> {
  try {
    const selectedMatter = (document.getElementById("matter-select") as HTMLSelectElement).value;
    if (!selectedMatter) {
      return { success: false, updatedRows: 0, error: "No matter selected" };
    }

    const profiles = getMatterProfiles();
    const currentProfile = profiles.find((p) => p.name === selectedMatter);

    if (!currentProfile || !currentProfile.rules || !currentProfile.rules.maxDailyHours) {
      return { success: false, updatedRows: 0, error: "Max Daily Hours rule not found in profile" };
    }

    const maxDailyHoursRule = currentProfile.rules.maxDailyHours;
    if (!maxDailyHoursRule.enabled) {
      return { success: false, updatedRows: 0, error: "Max Daily Hours rule is disabled" };
    }

    return await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = worksheet.getUsedRange();
      usedRange.load(["rowCount", "columnCount", "values"]);

      await context.sync();

      if (!usedRange || usedRange.rowCount < 2) {
        return { success: true, updatedRows: 0 };
      }

      const values = usedRange.values;
      const headers = values[0] as string[];

      // Find required columns
      const nameCol = headers.findIndex((h) => h && h.toLowerCase() === "name");
      const dateCol = headers.findIndex((h) => h && h.toLowerCase() === "date");
      const timeCol = headers.findIndex((h) => h && h.toLowerCase() === "amended time");
      const originalTimeCol = headers.findIndex((h) => h && h.toLowerCase() === "original time");
      const chargeCol = headers.findIndex((h) => h && h.toLowerCase() === "charge");
      const notesCol = headers.findIndex((h) => h && h.toLowerCase() === "notes");

      if (nameCol === -1) {
        return { success: false, updatedRows: 0, error: "Name column not found" };
      }
      if (dateCol === -1) {
        return { success: false, updatedRows: 0, error: "Date column not found" };
      }
      if (chargeCol === -1) {
        return { success: false, updatedRows: 0, error: "Charge column not found" };
      }
      if (notesCol === -1) {
        return { success: false, updatedRows: 0, error: "Notes column not found" };
      }

      // Use Amended Time if available, otherwise Original Time, otherwise Time
      let actualTimeCol = timeCol;
      if (actualTimeCol === -1) {
        actualTimeCol = originalTimeCol;
      }
      if (actualTimeCol === -1) {
        actualTimeCol = headers.findIndex((h) => h && h.toLowerCase() === "time");
      }
      if (actualTimeCol === -1) {
        return { success: false, updatedRows: 0, error: "Time column not found" };
      }

      addDebugInfo(
        `Max Daily Hours Rule: Using time column "${headers[actualTimeCol]}" at index ${actualTimeCol}`
      );
      addDebugInfo(`Max Daily Hours Rule: Max hours limit = ${maxDailyHoursRule.maxHours}`);

      // Group entries by fee earner and date
      const dailyHours = new Map<string, number>();
      const rowsToCheck = new Map<string, number[]>();

      // First pass: calculate total hours per fee earner per day
      for (let i = 1; i < usedRange.rowCount; i++) {
        const name = values[i][nameCol];
        const date = values[i][dateCol];
        const timeValue = values[i][actualTimeCol];

        if (!name || !date || !timeValue) continue;

        // Parse the time value
        const hours = parseTimeToHours(timeValue);
        if (hours === null) continue;

        // Create key for fee earner + date
        const key = `${name}|${date}`;

        // Add hours to daily total
        const currentTotal = dailyHours.get(key) || 0;
        dailyHours.set(key, currentTotal + hours);

        // Track which rows belong to this fee earner/date
        if (!rowsToCheck.has(key)) {
          rowsToCheck.set(key, []);
        }
        rowsToCheck.get(key)!.push(i);
      }

      // Second pass: mark rows where daily limit is exceeded
      let updatedRows = 0;

      for (const [key, totalHours] of dailyHours.entries()) {
        if (totalHours > maxDailyHoursRule.maxHours) {
          const [name, date] = key.split("|");
          addDebugInfo(
            `Max Daily Hours exceeded for ${name} on ${date}: ${totalHours.toFixed(2)} hours`
          );

          // Update all rows for this fee earner/date combination
          const rows = rowsToCheck.get(key) || [];
          for (const rowIndex of rows) {
            const chargeCell = usedRange.getCell(rowIndex, chargeCol);
            const notesCell = usedRange.getCell(rowIndex, notesCol);

            // Set charge value
            chargeCell.values = [[maxDailyHoursRule.chargeValue]];

            // Update notes
            const currentNotes = values[rowIndex][notesCol] || "";
            const noteToAdd = maxDailyHoursRule.noteText;

            if (!currentNotes.includes(noteToAdd)) {
              const newNotes = currentNotes ? `${currentNotes}; ${noteToAdd}` : noteToAdd;
              notesCell.values = [[newNotes]];
            }

            updatedRows++;
          }
        }
      }

      await context.sync();

      addDebugInfo(`Max Daily Hours Rule: Updated ${updatedRows} rows`);
      return { success: true, updatedRows };
    });
  } catch (error) {
    console.error("Error in applyMaxDailyHoursRuleWithResult:", error);
    return { success: false, updatedRows: 0, error: error.message };
  }
}

// Helper function to parse time values to decimal hours
function parseTimeToHours(timeValue: any): number | null {
  if (typeof timeValue === "number") {
    return timeValue;
  }

  if (typeof timeValue === "string") {
    // Try to parse HH:MM format
    const timeMatch = timeValue.match(/^(\d+):(\d+)$/);
    if (timeMatch) {
      const hours = parseInt(timeMatch[1]);
      const minutes = parseInt(timeMatch[2]);
      return hours + minutes / 60;
    }

    // Try to parse decimal format
    const decimal = parseFloat(timeValue);
    if (!isNaN(decimal)) {
      return decimal;
    }
  }

  return null;
}

// Undo Name Standardisation functionality
async function undoNameStandardisation() {
  try {
    showMessage("Undo functionality is not yet implemented", "info");
    console.log("undoNameStandardisation called - functionality to be implemented");
  } catch (error) {
    console.error("Error in undoNameStandardisation:", error);
    showMessage("Error during undo operation", "error");
  }
}

// Load nickname database functionality (stub)
function loadNicknameDatabase(customNicknames: any) {
  try {
    console.log("loadNicknameDatabase called with:", customNicknames);
    addDebugInfo(`Loading nickname database with ${Object.keys(customNicknames || {}).length} custom nicknames`);
    // TODO: Implement nickname database loading functionality
  } catch (error) {
    console.error("Error in loadNicknameDatabase:", error);
    addDebugInfo("Error loading nickname database");
  }
}

// Get nickname database from form (stub)
function getNicknameDatabaseFromForm(): any {
  try {
    console.log("getNicknameDatabaseFromForm called");
    addDebugInfo("Getting nickname database from form - returning empty object for now");
    // TODO: Implement nickname database form reading functionality
    // This should read the nickname form data and return an object with nickname mappings
    return {};
  } catch (error) {
    console.error("Error in getNicknameDatabaseFromForm:", error);
    addDebugInfo("Error getting nickname database from form");
    return {};
  }
}
