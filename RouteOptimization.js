/**
 * ============================================================================
 * AUTOMATED ROUTE OPTIMIZATION 
 * (FTL checks, Best-Fit Algorithm, Map distances, Constraints & Constraints)
 * ============================================================================
 */
function generateOptimizedRoutes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPlanning = ss.getSheetByName("Planning");
  var sheetRoutes = ss.getSheetByName("ROUTES");
  var sheetCarriers = ss.getSheetByName("CARRIERS");
  var sheetAddresses = ss.getSheetByName("Addresses"); 
  var sheetForbidden = ss.getSheetByName("Forbidden_Links");
  var sheetFtl = ss.getSheetByName("FTL"); 
  var sheetConditions = ss.getSheetByName("CONDITIONS"); 
  
  var sheetTimesTueThu = ss.getSheetByName("LOADING_TIMES_TUE_THU");
  var sheetTimesRest = ss.getSheetByName("LOADING_TIMES_REST");
  
  if (!sheetPlanning || !sheetRoutes || !sheetCarriers || !sheetAddresses || !sheetForbidden || !sheetFtl || !sheetConditions) {
    SpreadsheetApp.getUi().alert("Error: A required core sheet is missing.");
    return;
  }
  if (!sheetTimesTueThu || !sheetTimesRest) {
    SpreadsheetApp.getUi().alert("Error: Loading times sheets not found.");
    return;
  }

  // --- LOAD CONDITIONS ---
  var conditionsData = sheetConditions.getDataRange().getValues();
  var maxWeight = parseFloat(String(conditionsData[0][1]).replace(/[^\d.,]/g, "").replace(",", ".")) || 23500;
  var maxPallets = parseFloat(String(conditionsData[1][1]).replace(/[^\d.,]/g, "").replace(",", ".")) || 33;
  var minFtlPallets = parseFloat(String(conditionsData[2][1]).replace(/[^\d.,]/g, "").replace(",", ".")) || 26;
  var carrierRule = String(conditionsData[3][1]).trim().toUpperCase();
  var minPalletsPrimaryTarget = 5; 
  
  var ltlConstraints = {};
  for (var p = 5; p < conditionsData.length; p++) {
    var id1 = String(conditionsData[p][0]).trim();
    var limit1 = parseFloat(conditionsData[p][2]);
    if (id1 !== "" && !isNaN(limit1)) ltlConstraints[id1] = limit1;
    
    var id2 = String(conditionsData[p][4]).trim();
    var limit2 = parseFloat(conditionsData[p][6]);
    if (id2 !== "" && !isNaN(limit2)) ltlConstraints[id2] = limit2;
  }
  
  var originAddress = "Modletice 91, 251 01 Modletice, Česko"; 
  var routeCache = {};

  // --- OPTIMIZATION: Load Carriers into Dictionary ---
  var carrierData = sheetCarriers.getDataRange().getValues();
  var carrierDict = {};
  for (var col = 0; col < carrierData[0].length; col++) {
    for (var row = 2; row < carrierData.length; row++) {
      var carrierId = String(carrierData[row][col]).trim();
      if (carrierId !== "") {
        carrierDict[carrierId] = { name: carrierData[0][col], number: carrierData[1][col] };
      }
    }
  }

  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  var formattedDate = Utilities.formatDate(tomorrow, ss.getSpreadsheetTimeZone(), "dd.MM.yyyy");
  var dayOfWeek = tomorrow.getDay(); 
  
  var activeTimesSheet = (dayOfWeek === 2 || dayOfWeek === 4) ? sheetTimesTueThu : sheetTimesRest;
  var timesData = activeTimesSheet.getDataRange().getDisplayValues();

  var addressData = sheetAddresses.getDataRange().getValues();
  var addressDict = {};
  for (var i = 1; i < addressData.length; i++) { 
    var addrId = String(addressData[i][0]).trim(); 
    var street = String(addressData[i][2]).trim();
    var zip = String(addressData[i][3]).trim();
    var city = String(addressData[i][4]).trim();
    
    var availableCorridors = [];
    for (var colCorridor = 5; colCorridor <= 7; colCorridor++) {
      if (addressData[i][colCorridor]) {
        var k = String(addressData[i][colCorridor]).trim();
        if (k !== "") availableCorridors.push(k);
      }
    }
    if (availableCorridors.length === 0) availableCorridors = ["NO_CORRIDOR"];

    if (addrId && street) {
      addressDict[addrId] = { address: street + ", " + zip + " " + city + ", Česko", corridors: availableCorridors };
    }
  }

  var forbiddenData = sheetForbidden.getDataRange().getValues();
  var forbiddenDict = {};
  for (var z = 2; z < forbiddenData.length; z++) { 
    var rowIds = [];
    for (var col = 0; col < 3; col++) {
      var cellValue = String(forbiddenData[z][col]).trim();
      if (cellValue !== "") {
        var splitVals = cellValue.split("/");
        for (var s = 0; s < splitVals.length; s++) {
          var cleanId = splitVals[s].trim();
          if (cleanId !== "" && !isNaN(cleanId)) rowIds.push(cleanId);
        }
      }
    }
    for (var a = 0; a < rowIds.length; a++) {
      var idA = rowIds[a];
      if (!forbiddenDict[idA]) forbiddenDict[idA] = [];
      for (var b = 0; b < rowIds.length; b++) {
        if (a === b) continue;
        var idB = rowIds[b];
        if (forbiddenDict[idA].indexOf(idB) === -1) forbiddenDict[idA].push(idB);
      }
    }
  }

  var divisions = [
    {idCol: 3, kmCol: 5, palCol: 6, weightCol: 7, checkCol: 9},   
    {idCol: 13, kmCol: 15, palCol: 16, weightCol: 17, checkCol: 19}, 
    {idCol: 23, kmCol: 25, palCol: 26, weightCol: 27, checkCol: 29}  
  ];

  var planningValues = sheetPlanning.getRange(1, 1, 47, 30).getValues();
  
  // FAILSAFE: Double-check for manual FTL inputs
  divisions.forEach(function(div) {
    for (var r = 5; r < 47; r++) {
      var id = String(planningValues[r][div.idCol - 1]).trim();
      var pallets = parseFloat(planningValues[r][div.palCol - 1]) || 0;
      var weight = parseFloat(planningValues[r][div.weightCol - 1]) || 0;
      var cellStatus = planningValues[r][div.checkCol - 1]; 
      
      if (id !== "" && cellStatus !== "CONFIRMED" && (pallets >= minFtlPallets || weight >= maxWeight)) {
        var remainingPallets = pallets;
        var remainingWeight = weight;

        while (remainingPallets >= minFtlPallets || remainingWeight >= maxWeight) {
          var palletRatio = maxPallets / remainingPallets;
          var weightRatio = maxWeight / remainingWeight;
          var limitRatio = Math.min(palletRatio, weightRatio, 1);

          var autoPallets = remainingPallets * limitRatio;
          var autoWeight = remainingWeight * limitRatio;

          var palletsStr = (Math.round(autoPallets * 100) / 100).toFixed(2).replace(".", ",");
          var weightStr = (Math.round(autoWeight * 100) / 100).toFixed(2).replace(".", ",");

          var carrierName = carrierDict[id] ? carrierDict[id].name : "Not_Found";
          var carrierNumber = carrierDict[id] ? carrierDict[id].number : "Not_Found";

          sheetFtl.appendRow([formattedDate, id, palletsStr, weightStr, carrierNumber, carrierName]);

          remainingPallets -= autoPallets;
          remainingWeight -= autoWeight;
        }

        planningValues[r][div.palCol - 1] = remainingPallets;
        planningValues[r][div.weightCol - 1] = remainingWeight;

        if (remainingPallets < 0.1) {
          sheetPlanning.getRange(r + 1, div.palCol).setValue("");
          sheetPlanning.getRange(r + 1, div.weightCol).setValue("");
          sheetPlanning.getRange(r + 1, div.checkCol)
                  .setDataValidation(null).setValue("CONFIRMED")
                  .setBackground("#ea9999").setFontColor("#ffffff").setFontWeight("bold");
          planningValues[r][div.checkCol - 1] = "CONFIRMED"; 
        } else {
          sheetPlanning.getRange(r + 1, div.palCol).setValue(Number(remainingPallets.toFixed(2)));
          sheetPlanning.getRange(r + 1, div.weightCol).setValue(Number(remainingWeight.toFixed(2)));
        }
      }
    }
  });

  var branches = [];

  divisions.forEach(function(div) {
    for (var r = 5; r < 47; r++) {
      var id = String(planningValues[r][div.idCol - 1]).trim();
      var pallets = parseFloat(planningValues[r][div.palCol - 1]) || 0;
      var weight = parseFloat(planningValues[r][div.weightCol - 1]) || 0;
      var km = parseFloat(String(planningValues[r][div.kmCol - 1]).replace(",", ".").replace(/[^\d.]/g, "")) || 0;
      var cellStatus = planningValues[r][div.checkCol - 1]; 
      
      if (id !== "" && pallets > 0 && cellStatus !== "CONFIRMED") {
        var addrObj = addressDict[id];
        var address = addrObj ? addrObj.address : (id + ", Česko"); 
        var corridorList = addrObj ? addrObj.corridors : ["NO_CORRIDOR"];
        var cellA1 = sheetPlanning.getRange(r + 1, div.checkCol).getA1Notation();

        branches.push({
          id: id, km: km, pallets: pallets, weight: weight, row: r + 1, col: div.checkCol, a1: cellA1, address: address, corridors: corridorList
        });
      }
    }
  });

  branches.sort(function(a, b) { return b.km - a.km; });

  var generatedRoutes = [];
  var processed = {}; 

  // --- BEST-FIT ALGORITHM LOOP ---
  for (var i = 0; i < branches.length; i++) {
    var primaryBranch = branches[i];
    if (processed[primaryBranch.id]) continue;
    
    if (primaryBranch.pallets < minPalletsPrimaryTarget) continue;

    var routeOptions = []; 
    routeOptions.push({
      branches: [primaryBranch],
      pallets: primaryBranch.pallets,
      weight: primaryBranch.weight,
      score: (primaryBranch.pallets / maxPallets) + (primaryBranch.weight / maxWeight)
    });

    var activeCorridors = primaryBranch.corridors.slice(); 
    var validCandidates = [];

    for (var j = i + 1; j < branches.length; j++) {
      var candidate = branches[j];
      if (processed[candidate.id]) continue;
      
      var minLtlLimit = ltlConstraints[candidate.id] || 0;
      if (candidate.pallets < minLtlLimit) continue;
      
      var corridorIntersection = activeCorridors.filter(function(c) { return candidate.corridors.indexOf(c) !== -1; });
      if (corridorIntersection.length === 0) continue;
      if (forbiddenDict[primaryBranch.id] && forbiddenDict[primaryBranch.id].indexOf(candidate.id) !== -1) continue;
      
      validCandidates.push({ data: candidate, intersection: corridorIntersection });
    }

    for (var x = 0; x < validCandidates.length; x++) {
      var cand1 = validCandidates[x].data;
      var pallets2 = primaryBranch.pallets + cand1.pallets;
      var weight2 = primaryBranch.weight + cand1.weight;

      if (pallets2 <= maxPallets && weight2 <= maxWeight) {
        routeOptions.push({
          branches: [primaryBranch, cand1],
          pallets: pallets2,
          weight: weight2,
          score: (pallets2 / maxPallets) + (weight2 / maxWeight)
        });

        for (var y = x + 1; y < validCandidates.length; y++) {
          var cand2 = validCandidates[y].data;
          
          if (forbiddenDict[cand1.id] && forbiddenDict[cand1.id].indexOf(cand2.id) !== -1) continue;
          if (forbiddenDict[primaryBranch.id] && forbiddenDict[primaryBranch.id].indexOf(cand2.id) !== -1) continue;
          
          var finalIntersection = validCandidates[x].intersection.filter(function(c) { return cand2.corridors.indexOf(c) !== -1; });
          if (finalIntersection.length === 0) continue;

          var pallets3 = pallets2 + cand2.pallets;
          var weight3 = weight2 + cand2.weight;

          if (pallets3 <= maxPallets && weight3 <= maxWeight) {
            routeOptions.push({
              branches: [primaryBranch, cand1, cand2],
              pallets: pallets3,
              weight: weight3,
              score: (pallets3 / maxPallets) + (weight3 / maxWeight)
            });
          }
        }
      }
    }

    routeOptions.sort(function(a, b) { return b.score - a.score; });

    var bestRouteFound = null;
    var baseDistance = primaryBranch.km;
    var maxAllowedDetour = baseDistance * 1.15; 

    for (var m = 0; m < routeOptions.length; m++) {
      var testVehicle = routeOptions[m];
      
      if (testVehicle.branches.length === 1) {
        bestRouteFound = testVehicle;
        break;
      }

      var testDistanceKm = calculateShortestRoute(originAddress, testVehicle.branches, routeCache);
      if (testDistanceKm !== null && testDistanceKm <= maxAllowedDetour) {
        bestRouteFound = testVehicle;
        break; 
      }
    }

    bestRouteFound.branches.forEach(function(p) { processed[p.id] = true; });
    generatedRoutes.push(bestRouteFound);
  }

  var finalRouteRows = [];
  var cellsToLock = []; 

  generatedRoutes.forEach(function(route) {
    if (route.branches.length > 1) { 
      
      var refBranch;
      
      if (carrierRule === "ALWAYS_FURTHEST_BRANCH") {
        refBranch = route.branches[0]; 
      } 
      else if (carrierRule === "MOST_PALLETS") {
        refBranch = route.branches.reduce(function(prev, current) {
          return (prev.pallets > current.pallets) ? prev : current;
        });
      } 
      else {
        refBranch = route.branches[0]; 
      }

      var refID = refBranch.id; 
      
      var carrierName = carrierDict[refID] ? carrierDict[refID].name : "Not_Found";
      var carrierNumber = carrierDict[refID] ? carrierDict[refID].number : "Not_Found";

      var routeString = route.branches.map(function(p) { return p.id; }).join("/");
      var palletsStr = route.pallets.toFixed(2).replace(".", ",");
      var weightStr = route.weight.toFixed(2).replace(".", ",");
      
      var loadingTime = findTimeInTable(refID, timesData);

      finalRouteRows.push([formattedDate, loadingTime, routeString, palletsStr, weightStr, carrierNumber, carrierName]);
      route.branches.forEach(function(p) { cellsToLock.push(p.a1); });
    }
  });

  if (finalRouteRows.length > 0) {
    sheetRoutes.getRange(sheetRoutes.getLastRow() + 1, 1, finalRouteRows.length, 7).setValues(finalRouteRows);
  }

  if (cellsToLock.length > 0) {
    var rangeList = sheetPlanning.getRangeList(cellsToLock);
    rangeList.setValue("CONFIRMED").setBackground("#ea9999").setFontColor("#ffffff").setFontWeight("bold");
    cellsToLock.forEach(function(a1Notation) { sheetPlanning.getRange(a1Notation).setDataValidation(null); });
  }

  SpreadsheetApp.getUi().alert("Automated optimization complete!\nConditions limits and carrier rules were successfully applied.");
}

/**
 * ============================================================================
 * HELPER FUNCTIONS
 * ============================================================================
 */
function findTimeInTable(searchId, tableData) {
  for (var i = 0; i < tableData.length; i++) {
    if (String(tableData[i][0]).trim() === String(searchId).trim()) return tableData[i][1]; 
  }
  return ""; 
}

function calculateShortestRoute(origin, branchesArray, cacheObj) {
  try {
    var destinationAddress = branchesArray[0].address;
    var waypoints = [];
    var cacheKeyAddresses = [origin];
    
    for (var i = 1; i < branchesArray.length; i++) {
       waypoints.push(branchesArray[i].address);
       cacheKeyAddresses.push(branchesArray[i].id); 
    }
    cacheKeyAddresses.push(branchesArray[0].id); 
    var cacheKey = cacheKeyAddresses.join("_");
    
    if (cacheObj && cacheObj[cacheKey]) return cacheObj[cacheKey];

    var directionFinder = Maps.newDirectionFinder()
      .setOrigin(origin)
      .setDestination(destinationAddress)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .setOptimizeWaypoints(true); 

    for (var w = 0; w < waypoints.length; w++) directionFinder.addWaypoint(waypoints[w]);

    var directions = directionFinder.getDirections();
    if (directions.routes && directions.routes.length > 0) {
      var totalMeters = 0;
      var legs = directions.routes[0].legs;
      for (var l = 0; l < legs.length; l++) totalMeters += legs[l].distance.value;
      
      var kilometers = totalMeters / 1000;
      if (cacheObj) cacheObj[cacheKey] = kilometers;
      return kilometers;
    }
  } catch (e) { Logger.log("Maps API Error: " + e.toString()); }
  return null; 
}
