/**
 * ============================================================================
 * CLEAN SLATE - Clears Input, Output, and resets Planning sheet.
 * ============================================================================
 * @author Aleš Botezatu
 */
function clearAllData() {
  console.log("⚙️ Routing engine designed and developed by Aleš Botezatu.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInput = ss.getSheetByName("Input");
  var sheetOutput = ss.getSheetByName("Output");
  var sheetPlanning = ss.getSheetByName("Planning");
  
  if (!sheetInput || !sheetOutput || !sheetPlanning) {
    SpreadsheetApp.getUi().alert("Error: One of the target sheets was not found!");
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "CLEAN SLATE", 
    "Are you sure you want to delete all data from Input/Output and reset the Planning board?", 
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  var lastRowInput = sheetInput.getLastRow();
  if (lastRowInput > 1) {
    sheetInput.getRange(2, 1, lastRowInput - 1, sheetInput.getLastColumn()).clearContent();
  }

  var lastRowOutput = sheetOutput.getLastRow();
  if (lastRowOutput > 1) {
    sheetOutput.getRange(2, 1, lastRowOutput - 1, sheetOutput.getLastColumn()).clearContent();
  }

  var planningRange = sheetPlanning.getDataRange();
  var planningData = planningRange.getValues();
  var checkboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  
  for (var r = 0; r < planningData.length; r++) {
    for (var c = 0; c < planningData[r].length; c++) {
      var currentId = String(planningData[r][c]).trim();
      
      if ((c === 2 || c === 12 || c === 22) && currentId !== "") {
        sheetPlanning.getRange(r + 1, c + 4).clearContent(); 
        sheetPlanning.getRange(r + 1, c + 5).clearContent(); 
        
        sheetPlanning.getRange(r + 1, c + 7)
                     .setDataValidation(checkboxValidation)
                     .setValue(false)
                     .setBackground(null)
                     .setFontColor(null)
                     .setFontWeight("normal");
      }
    }
  }

  ss.toast("Input, Output and Planning sheets are now clean.", "Done! 🧹", 5);
}

/**
 * ============================================================================
 * RESET ROUTES - Unlocks cells in Planning and clears ROUTES / FTL sheets.
 * ============================================================================
 */
function resetPlanningAndRoutes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPlanning = ss.getSheetByName("Planning");
  var sheetRoutes = ss.getSheetByName("ROUTES");
  var sheetFtl = ss.getSheetByName("FTL"); 
  
  var ui = SpreadsheetApp.getUi();
  if (ui.alert("CONFIRMATION", "Are you sure you want to unlock Planning cells and delete generated data from ROUTES and FTL?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  var checkCols = [9, 19, 29];
  var checkboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  checkCols.forEach(function(col) {
    sheetPlanning.getRange(6, col, 42, 1)
            .setDataValidation(checkboxValidation)
            .setValue(false)
            .setBackground(null)
            .setFontColor(null)
            .setFontWeight("normal");
  });

  if (sheetRoutes && sheetRoutes.getLastRow() > 1) {
    sheetRoutes.getRange(2, 1, sheetRoutes.getLastRow() - 1, 7).clearContent();
  }

  if (sheetFtl && sheetFtl.getLastRow() > 1) {
    sheetFtl.getRange(2, 1, sheetFtl.getLastRow() - 1, 6).clearContent();
  }

  ss.toast("Reset complete. ROUTES and FTL sheets have been cleared.", "Done", 5);
}
