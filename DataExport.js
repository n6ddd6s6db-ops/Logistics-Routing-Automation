/**
 * ============================================================================
 * DATA EXPORT (From Input to Planning + FTL Extraction)
 * ============================================================================
 * @author Aleš Botezatu
 */
function exportDataToPlanning() {
  console.log("⚙️ Routing engine designed and developed by Aleš Botezatu.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInput = ss.getSheetByName("Input");
  var sheetPlanning = ss.getSheetByName("Planning");
  var sheetFtl = ss.getSheetByName("FTL"); 
  var sheetConditions = ss.getSheetByName("CONDITIONS");
  var sheetCarriers = ss.getSheetByName("CARRIERS");
  
  if (!sheetInput || !sheetPlanning || !sheetFtl || !sheetConditions || !sheetCarriers) {
    SpreadsheetApp.getUi().alert("Error: Missing one of the required sheets (Input, Planning, FTL, CONDITIONS, CARRIERS).");
    return;
  }

  // --- LOAD CONDITIONS AND CARRIERS ---
  var conditionsData = sheetConditions.getDataRange().getValues();
  var maxWeight = parseFloat(String(conditionsData[0][1]).replace(/[^\d.,]/g, "").replace(",", ".")) || 23500;
  var maxPallets = parseFloat(String(conditionsData[1][1]).replace(/[^\d.,]/g, "").replace(",", ".")) || 33;
  var minFtlPallets = parseFloat(String(conditionsData[2][1]).replace(/[^\d.,]/g, "").replace(",", ".")) || 26;

  var carrierData = sheetCarriers.getDataRange().getValues();
  
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  var formattedDate = Utilities.formatDate(tomorrow, ss.getSpreadsheetTimeZone(), "dd.MM.yyyy");

  // --- COLLECT AND SUM DATA FROM INPUT SHEET ---
  var inputValues = sheetInput.getDataRange().getValues();
  var dataSums = {};

  for (var i = 1; i < inputValues.length; i++) {
    var customerId = String(inputValues[i][0]).trim(); 
    if (!customerId || customerId === "") continue;

    var pallets = parseFloat(inputValues[i][1]) || 0; 
    var weight = parseFloat(inputValues[i][2]) || 0;   

    if (!dataSums[customerId]) {
      dataSums[customerId] = { pallets: 0, weight: 0, originalPallets: 0 };
    }
    dataSums[customerId].pallets += pallets;
    dataSums[customerId].weight += weight;
    dataSums[customerId].originalPallets += pallets; 
  }

  // --- EXTRACT FULL TRUCKLOADS (FTL) ---
  var ftlRows = [];
  for (var id in dataSums) {
    var remainingPallets = dataSums[id].pallets;
    var remainingWeight = dataSums[id].weight;

    while (remainingPallets >= minFtlPallets || remainingWeight >= maxWeight) {
      var palletRatio = maxPallets / remainingPallets;
      var weightRatio = maxWeight / remainingWeight;
      var limitingRatio = Math.min(palletRatio, weightRatio, 1);

      var loadPallets = remainingPallets * limitingRatio;
      var loadWeight = remainingWeight * limitingRatio;

      var palletsStr = (Math.round(loadPallets * 100) / 100).toFixed(2).replace(".", ",");
      var weightStr = (Math.round(loadWeight * 100) / 100).toFixed(2).replace(".", ",");

      var carrierName = "Not_Found";
      var carrierNumber = "Not_Found";
      for (var col = 0; col < carrierData[0].length; col++) {
        for (var row = 2; row < carrierData.length; row++) {
          if (String(carrierData[row][col]).trim() === id) {
            carrierName = carrierData[0][col];
            carrierNumber = carrierData[1][col];
            break;
          }
        }
        if (carrierName !== "Not_Found") break;
      }

      ftlRows.push([formattedDate, id, palletsStr, weightStr, carrierNumber, carrierName]);

      remainingPallets -= loadPallets;
      remainingWeight -= loadWeight;
    }
    dataSums[id].pallets = remainingPallets;
    dataSums[id].weight = remainingWeight;
  }

  if (ftlRows.length > 0) {
    sheetFtl.getRange(sheetFtl.getLastRow() + 1, 1, ftlRows.length, 6).setValues(ftlRows);
  }

  // --- WRITE REMAINING LTL DATA TO PLANNING SHEET ---
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

        if (dataSums[currentId]) {
           var rPallets = dataSums[currentId].pallets;
           var rWeight = dataSums[currentId].weight;
           
           if (rPallets > 0.1) {
              sheetPlanning.getRange(r + 1, c + 4).setValue(Number(rPallets.toFixed(2)));
              sheetPlanning.getRange(r + 1, c + 5).setValue(Number(rWeight.toFixed(2)));
           } else if (dataSums[currentId].originalPallets > 0) {
              sheetPlanning.getRange(r + 1, c + 7)
                           .setDataValidation(null)
                           .setValue("CONFIRMED")
                           .setBackground("#ea9999")
                           .setFontColor("#ffffff")
                           .setFontWeight("bold");
           }
        }
      }
    }
  }

  SpreadsheetApp.getUi().alert("Export successfully completed!\nFTLs were automatically extracted. Only remaining LTLs were written to the Planning sheet.");
}
