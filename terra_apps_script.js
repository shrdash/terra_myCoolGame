/**
 * ============================================================================
 * TERRA // OUTPOST 84 — GOOGLE APPS SCRIPT ENGINE (v4.2)
 * A Colony Terraforming Simulator Played Exclusively Inside Google Sheets
 * ============================================================================
 * 
 * HOW TO INSTALL IN GOOGLE SHEETS:
 * 1. Open Google Sheets (create a new blank spreadsheet or import Terra_Colony_Simulator.xlsx).
 * 2. Click "Extensions" -> "Apps Script".
 * 3. Delete any default code in Code.gs and paste this ENTIRE file.
 * 4. Click the Save icon (Ctrl+S / Cmd+S).
 * 5. Return to your Google Sheet and refresh the page.
 * 6. You will see a new menu item called "🚀 Terra Engine" in the top menu bar!
 * 7. Click "🚀 Terra Engine" -> "🔄 Initialize / Reset Simulation" to start.
 * 
 * HOW TO PLAY:
 * - Click sectors on the "Surface_Sector_Grid" sheet and type facility codes:
 *   [SP] = Solar Array (100 RC)
 *   [GD] = Geothermal Drill (250 RC, needs fissure **)
 *   [IH] = Ice Harvester (150 RC, needs glacier ❄❄)
 *   [TF] = Thermal Gas Factory (200 RC)
 *   [AG] = Algae Greenhouse Dome (300 RC)
 *   [AS] = Atmospheric Scrubber (350 RC)
 *   [HD] = Biosphere Habitat Pod (400 RC)
 *   [MS] = Magnetic Shield Pylon (600 RC, needs ridge ^^)
 * - To advance time, click "🚀 Terra Engine" -> "Advance 1 Sol", or tick the checkbox in B2!
 * ============================================================================
 */

// Initialize Custom Menu upon opening spreadsheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Terra Engine')
    .addItem('▶ Advance 1 Sol (Turn)', 'advanceSol')
    .addItem('⏩ Fast-Forward 5 Sols', 'advanceFiveSols')
    .addSeparator()
    .addItem('🎨 Apply Conditional Formatting', 'setupConditionalFormatting')
    .addItem('🔄 Initialize / Reset Simulation', 'initializeSimulation')
    .addItem('📖 How to Play / Mission Briefing', 'showMissionBriefing')
    .addToUi();
}

/**
 * onEdit Trigger: Allows ticking a checkbox in cell B2 to step turns seamlessly!
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() === "Surface_Sector_Grid" && e.range.getA1Notation() === "B2") {
    if (e.value === "TRUE") {
      advanceSol();
      sheet.getRange("B2").setValue(false); // Reset checkbox for next click
    }
  }
}

/**
 * Initializes or resets the entire planetary simulation
 */
function initializeSimulation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Surface Sector Grid
  var gridSheet = ss.getSheetByName("Surface_Sector_Grid");
  if (!gridSheet) gridSheet = ss.insertSheet("Surface_Sector_Grid");
  
  gridSheet.clear();
  
  // Headers & Status
  gridSheet.getRange("A1").setValue("TERRA // OUTPOST 84 — PLANETARY SECTOR RADAR").setFontWeight("bold").setFontSize(14).setFontColor("#0f9d58");
  gridSheet.getRange("A2").setValue("NEXT SOL TRIGGER:");
  gridSheet.getRange("B2").insertCheckboxes();
  gridSheet.getRange("C2").setValue("CURRENT SOL:");
  gridSheet.getRange("D2").setValue(1).setFontWeight("bold");
  gridSheet.getRange("F2").setValue("MISSION OBJECTIVE: Reach 21% O2, +15°C, 1013 hPa, 10k Pop before Sol 100").setFontColor("#555555");
  
  // Grid Columns Header
  var colLabels = [["SECTOR", "X01","X02","X03","X04","X05","X06","X07","X08","X09","X10","X11","X12","X13","X14","X15","X16"]];
  gridSheet.getRange(4, 1, 1, 17).setValues(colLabels).setBackground("#1f2937").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Initial 16x16 Terrain Map
  var mapRows = [
    ["Y01","❄❄","❄❄","❄❄","..","..","^^","^^","..","..","..","..","..","^^","❄❄","❄❄","❄❄"],
    ["Y02","❄❄","❄❄","..","..","..","^^","..","..","..","..","..","..","..","..","❄❄","❄❄"],
    ["Y03","❄❄","..","..","~~","~~","..","..","..","**","..","..","..","..","..","..","❄❄"],
    ["Y04","..","..","~~","~~","~~","~~","..","..","..","..","..","..","^^","^^","..",".."],
    ["Y05","..","..","~~","~~","~~","..","..","..","..","..","..","..","^^","..","..",".."],
    ["Y06","..","^^","..","~~","..","..","..","[SP]","..","..","..","..","..","..","..",".."],
    ["Y07","..","^^","..","..","..","..","[HB]","[HB]","..","..","..","..","..","**","..",".."],
    ["Y08","..","..","..","..","..","..","[HB]","[AG]","..","..","..","..","..","..","..",".."],
    ["Y09","..","..","..","..","..","..","[AG]","[AG]","..","..","..","..","..","..","..",".."],
    ["Y10","..","..","..","..","..","..","..","..","..","..","~~","~~","..","..","..",".."],
    ["Y11","..","..","..","**","..","..","..","..","..","~~","~~","~~","~~","..","..",".."],
    ["Y12","..","^^","..","..","..","..","..","..","..","~~","~~","~~","~~","..","..",".."],
    ["Y13","..","^^","^^","..","..","..","..","..","..","..","~~","~~","..","..","..",".."],
    ["Y14","❄❄","..","..","..","..","..","..","..","**","..","..","..","..","..","..","❄❄"],
    ["Y15","❄❄","❄❄","..","..","..","^^","..","..","..","..","..","^^","..","..","❄❄","❄❄"],
    ["Y16","❄❄","❄❄","❄❄","..","..","^^","^^","..","..","..","..","..","^^","❄❄","❄❄","❄❄"]
  ];
  
  gridSheet.getRange(5, 1, 16, 17).setValues(mapRows).setHorizontalAlignment("center");
  gridSheet.getRange(5, 1, 16, 1).setBackground("#1f2937").setFontColor("#ffffff").setFontWeight("bold");
  
  // Telemetry Dashboard under map
  gridSheet.getRange("A22").setValue("COLONY TELEMETRY").setFontWeight("bold").setBackground("#0f9d58").setFontColor("#ffffff");
  gridSheet.getRange("B22").setValue("POPULATION: 250");
  gridSheet.getRange("D22").setValue("ENERGY: +145 MWh");
  gridSheet.getRange("F22").setValue("CREDITS: 500 RC");
  gridSheet.getRange("H22").setValue("HABITABILITY: 4.5%");
  
  gridSheet.getRange("A23:Q23").merge().setValue("FACILITIES: [SP] Solar Array (+30 MWh) | [GD] Geothermal (+75 MWh, +0.6°C) | [AG] Algae Dome (+0.3% O2) | [TF] Gas Factory (+1.5°C, +12hPa) | [HB] Habitat (+1k Pop)").setFontSize(9).setFontColor("#666666");
  gridSheet.getRange("A24:Q24").merge().setValue("TERRAIN: .. Barren Regolith | ^^ Mountain Ridge | ~~ Dry Basin | ** Thermal Fissure | ❄❄ Polar Ice Glacier | ≈≈ Liquid Ocean | ♣♣ Terraformed Biome").setFontSize(9).setFontColor("#666666");
  
  // 2. Setup Atmosphere Telemetry Sheet
  var atmoSheet = ss.getSheetByName("Atmosphere_Telemetry");
  if (!atmoSheet) atmoSheet = ss.insertSheet("Atmosphere_Telemetry");
  atmoSheet.clear();
  
  atmoSheet.getRange("A1").setValue("TERRA // ATMOSPHERIC TELEMETRY").setFontWeight("bold").setFontSize(14).setFontColor("#0f9d58");
  var atmoHeaders = [["PARAMETER", "CURRENT VALUE", "UNIT", "TARGET RANGE", "STATUS"]];
  atmoSheet.getRange(2, 1, 1, 5).setValues(atmoHeaders).setBackground("#1f2937").setFontColor("#ffffff").setFontWeight("bold");
  
  var atmoData = [
    ["Surface Temperature", -65.0, "°C", "+14.0 to +18.0", "LETHAL PERMAFROST"],
    ["Atmospheric Pressure", 180.0, "hPa", "980.0 to 1025.0", "THIN / VACUUM LETHAL"],
    ["Atmospheric Oxygen", 0.10, "%", "19.5% to 21.5%", "ANOXIC SUFFOCATION"],
    ["Hydrosphere Volume", 65.0, "Mega-Liters", "800.0 to 1500.0", "LOCKED IN GLACIERS"],
    ["Habitability Index", 4.2, "%", "100.0%", "HOSTILE"]
  ];
  atmoSheet.getRange(3, 1, 5, 5).setValues(atmoData);
  
  setupConditionalFormatting();
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Terra Simulation Initialized on Sol 1. Check your radar grid!", "PROJECT AETHEL ONLINE", 5);
}

/**
 * Advances the planetary simulation by 1 Sol
 */
function advanceSol() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gridSheet = ss.getSheetByName("Surface_Sector_Grid");
  var atmoSheet = ss.getSheetByName("Atmosphere_Telemetry");
  
  if (!gridSheet || !atmoSheet) {
    initializeSimulation();
    return;
  }
  
  // Read current Sol
  var sol = Number(gridSheet.getRange("D2").getValue()) || 1;
  sol += 1;
  gridSheet.getRange("D2").setValue(sol);
  
  // Scan 16x16 grid for facilities and terrain
  var gridValues = gridSheet.getRange(5, 2, 16, 16).getValues();
  
  var counts = {
    SP: 0, GD: 0, IH: 0, TF: 0, AG: 0, AS: 0, HB: 0, MS: 0,
    regolith: 0, ice: 0, water: 0, biome: 0, vent: 0, ridge: 0
  };
  
  for (var r = 0; r < 16; r++) {
    for (var c = 0; c < 16; c++) {
      var cell = String(gridValues[r][c]).trim();
      if (cell === "[SP]") counts.SP++;
      else if (cell === "[GD]") counts.GD++;
      else if (cell === "[IH]") counts.IH++;
      else if (cell === "[TF]") counts.TF++;
      else if (cell === "[AG]") counts.AG++;
      else if (cell === "[AS]") counts.AS++;
      else if (cell === "[HB]") counts.HB++;
      else if (cell === "[MS]") counts.MS++;
      else if (cell === "❄❄") counts.ice++;
      else if (cell === "≈≈") counts.water++;
      else if (cell === "♣♣") counts.biome++;
      else if (cell === "**") counts.vent++;
      else if (cell === "^^") counts.ridge++;
      else counts.regolith++;
    }
  }
  
  // Read current telemetry
  var currentTemp = Number(atmoSheet.getRange("B3").getValue()) || -65.0;
  var currentPress = Number(atmoSheet.getRange("B4").getValue()) || 180.0;
  var currentO2 = Number(atmoSheet.getRange("B5").getValue()) || 0.1;
  var currentWater = Number(atmoSheet.getRange("B6").getValue()) || 65.0;
  
  // Read Colony values
  var popMatch = String(gridSheet.getRange("B22").getValue()).match(/\d+/);
  var currentPop = popMatch ? Number(popMatch[0]) : 250;
  
  var creditsMatch = String(gridSheet.getRange("F22").getValue()).match(/\d+/);
  var currentCredits = creditsMatch ? Number(creditsMatch[0]) : 500;
  
  // Power Balance Calculation
  var powerProduced = (counts.SP * 30) + (counts.GD * 75);
  var powerConsumed = (counts.IH * 10) + (counts.TF * 25) + (counts.AG * 15) + (counts.AS * 20) + (counts.HB * 20) + (counts.MS * 50);
  var netPower = powerProduced - powerConsumed;
  
  // Telemetry Delays & Yields (if power is positive)
  var powerEfficiency = (netPower >= 0) ? 1.0 : 0.25; // brownout causes severe drops
  
  // Temperature deltas
  var tempDelta = ((counts.GD * 0.6) + (counts.TF * 1.5)) * powerEfficiency;
  currentTemp = Math.round((currentTemp + tempDelta) * 10) / 10;
  
  // Pressure deltas
  var pressDelta = ((counts.TF * 12.0) + (counts.AS * 15.0)) * powerEfficiency;
  currentPress = Math.round((currentPress + pressDelta) * 10) / 10;
  
  // Oxygen deltas
  var o2Delta = ((counts.AG * 0.35) + (counts.biome * 0.05)) * powerEfficiency;
  currentO2 = Math.round((currentO2 + o2Delta) * 100) / 100;
  
  // Water deltas
  var waterHarvest = (counts.IH * 25) * powerEfficiency;
  var waterConsumed = (counts.AG * 8) + (counts.HB * 5);
  var waterMelt = (currentTemp > 0) ? Math.min(counts.ice * 3, 20) : 0;
  currentWater = Math.max(0, Math.round(currentWater + waterHarvest - waterConsumed + waterMelt));
  
  // Population growth and credits
  var maxHabCap = 250 + (counts.HB * 1000);
  if (netPower >= 0 && currentO2 >= 0.1) {
    if (currentPop < maxHabCap) {
      currentPop = Math.min(maxHabCap, currentPop + 150);
    }
  } else {
    currentPop = Math.max(50, currentPop - 80); // casualties in blackout
  }
  
  var creditYield = (counts.HB * 40) + Math.round(currentPop * 0.1);
  currentCredits += creditYield;
  
  // Habitability Score calculation
  var habTemp = (currentTemp >= -10 && currentTemp <= 30) ? 25 : Math.max(0, 25 - Math.abs(currentTemp - 15));
  var habPress = (currentPress >= 800 && currentPress <= 1200) ? 25 : Math.max(0, 25 * (currentPress / 1000));
  var habO2 = (currentO2 >= 18.0 && currentO2 <= 23.0) ? 30 : Math.max(0, 30 * (currentO2 / 20));
  var habWater = (currentWater >= 500) ? 20 : Math.max(0, 20 * (currentWater / 500));
  var habitability = Math.min(100, Math.round((habTemp + habPress + habO2 + habWater) * 10) / 10);
  
  // Dynamic Grid Evolution: Ice melting and Biome propagation
  var gridChanged = false;
  if (currentTemp > 0) {
    // Thaw random ice cell into ocean
    for (var r = 0; r < 16; r++) {
      for (var c = 0; c < 16; c++) {
        if (gridValues[r][c] === "❄❄" && Math.random() < 0.25) {
          gridValues[r][c] = "≈≈";
          gridChanged = true;
          break;
        }
      }
      if (gridChanged) break;
    }
  }
  
  if (currentO2 > 5.0 && currentWater > 200) {
    // Sprout vegetation on barren regolith near water
    for (var r = 0; r < 16; r++) {
      for (var c = 0; c < 16; c++) {
        if (gridValues[r][c] === ".." && Math.random() < 0.2) {
          gridValues[r][c] = "♣♣";
          gridChanged = true;
          break;
        }
      }
      if (gridChanged) break;
    }
  }
  
  if (gridChanged) {
    gridSheet.getRange(5, 2, 16, 16).setValues(gridValues);
  }
  
  // Write back Telemetry & Colony Status
  atmoSheet.getRange("B3").setValue(currentTemp);
  atmoSheet.getRange("E3").setValue(currentTemp < -20 ? "FREEZING EXTREME" : (currentTemp > 25 ? "HOT" : "BALMY HABITABLE"));
  
  atmoSheet.getRange("B4").setValue(currentPress);
  atmoSheet.getRange("E4").setValue(currentPress < 600 ? "LOW PRESSURE" : (currentPress > 1100 ? "HIGH PRESSURE" : "EARTH NORMAL (1 ATM)"));
  
  atmoSheet.getRange("B5").setValue(currentO2);
  atmoSheet.getRange("E5").setValue(currentO2 < 10.0 ? "HYPOXIC / TOXIC" : (currentO2 > 23.0 ? "HYPEROXIC" : "BREATHABLE AIR"));
  
  atmoSheet.getRange("B6").setValue(currentWater);
  atmoSheet.getRange("B7").setValue(habitability);
  
  gridSheet.getRange("B22").setValue("POPULATION: " + currentPop);
  gridSheet.getRange("D22").setValue("ENERGY: " + (netPower >= 0 ? "+" : "") + netPower + " MWh");
  gridSheet.getRange("F22").setValue("CREDITS: " + currentCredits + " RC");
  gridSheet.getRange("H22").setValue("HABITABILITY: " + habitability + "%");
  
  // Random Incident Roll (D100)
  var roll = Math.floor(Math.random() * 100) + 1;
  var alertMsg = "Sol " + sol + ": Environmental telemetry stable.";
  if (roll <= 15) {
    if (counts.MS === 0) {
      currentPress = Math.max(100, currentPress - 25);
      atmoSheet.getRange("B4").setValue(currentPress);
      alertMsg = "⚠️ SOL " + sol + " ALERT: Coronal Mass Ejection! Solar wind stripped 25 hPa atmosphere. Build [MS] Shield to protect!";
    } else {
      alertMsg = "🛡️ SOL " + sol + ": Solar flare deflected by Magnetic Shield Pylon!";
    }
  } else if (roll >= 85 && roll <= 92) {
    currentCredits += 300;
    gridSheet.getRange("F22").setValue("CREDITS: " + currentCredits + " RC");
    alertMsg = "🔬 SOL " + sol + " BREAKTHROUGH: Colonist R&D breakthrough granted +300 RC!";
  } else if (roll > 92) {
    currentWater += 80;
    atmoSheet.getRange("B6").setValue(currentWater);
    alertMsg = "☄️ SOL " + sol + " IMPACT: Cryo-meteor deposited +80 ML of fresh ice/water in northern basin!";
  }
  
  // Victory Check
  if (currentO2 >= 20.0 && currentTemp >= 14.0 && currentPress >= 980.0 && currentPop >= 5000) {
    SpreadsheetApp.getUi().alert("🎉 VICTORY ACHIEVED ON SOL " + sol + "!\n\nTerra Prime has reached full planetary habitability!\n- Breathable Atmosphere (O2 >= 20%)\n- Liquid Hydrosphere\n- Stable Earth-Normal Climate\n\nProject Aethel is a historic triumph. You may continue cultivating or reset for a new campaign.");
  } else {
    ss.toast(alertMsg, "SOL " + sol + " PROCESSED", 4);
  }
}

/**
 * Fast-forward 5 Sols in sequence
 */
function advanceFiveSols() {
  for (var i = 0; i < 5; i++) {
    advanceSol();
    Utilities.sleep(400);
  }
}

/**
 * Sets up conditional formatting rules for colors on the grid
 */
function setupConditionalFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Surface_Sector_Grid");
  if (!sheet) return;
  
  var range = sheet.getRange("B5:Q20");
  sheet.clearConditionalFormatRules();
  
  var rules = [];
  
  // Regolith ..
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("..").setBackground("#b45309").setFontColor("#fef3c7").setRanges([range]).build());
  // Mountain Ridge ^^
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("^^").setBackground("#4b5563").setFontColor("#f3f4f6").setRanges([range]).build());
  // Dry Basin ~~
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("~~").setBackground("#78350f").setFontColor("#fde68a").setRanges([range]).build());
  // Thermal Fissure **
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("**").setBackground("#dc2626").setFontColor("#ffffff").setRanges([range]).build());
  // Polar Ice ❄❄
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("❄❄").setBackground("#e0f2fe").setFontColor("#0284c7").setRanges([range]).build());
  // Liquid Ocean ≈≈
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("≈≈").setBackground("#2563eb").setFontColor("#ffffff").setRanges([range]).build());
  // Living Biome ♣♣
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("♣♣").setBackground("#059669").setFontColor("#ecfdf5").setRanges([range]).build());
  
  // Facilities
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("[SP]").setBackground("#f59e0b").setFontColor("#000000").setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("[GD]").setBackground("#ea580c").setFontColor("#ffffff").setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("[IH]").setBackground("#38bdf8").setFontColor("#000000").setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("[TF]").setBackground("#9333ea").setFontColor("#ffffff").setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("[AG]").setBackground("#10b981").setFontColor("#ffffff").setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("[AS]").setBackground("#06b6d4").setFontColor("#ffffff").setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("[HB]").setBackground("#6366f1").setFontColor("#ffffff").setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("[MS]").setBackground("#ec4899").setFontColor("#ffffff").setRanges([range]).build());
  
  sheet.setConditionalFormatRules(rules);
  ss.toast("Conditional formatting color rules mapped to 16x16 matrix.", "RADAR CALIBRATED", 3);
}

/**
 * Displays mission briefing modal
 */
function showMissionBriefing() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    "MISSION BRIEFING: PROJECT AETHEL\n\n" +
    "Welcome Commander. You are stationed aboard Outpost 84.\n" +
    "Your objective: Transform the arid, freezing exoplanet Terra Prime into a living second Earth.\n\n" +
    "HOW TO PLAY:\n" +
    "1. Inspect the 16x16 grid on 'Surface_Sector_Grid'.\n" +
    "2. Enter facility codes into cells (e.g. [SP] for Solar Array, [GD] on vents **).\n" +
    "3. Keep Net Energy positive! A brownout reduces terraforming efficiency by 75%.\n" +
    "4. Advance Sols using the 'Terra Engine' menu or by checking cell B2.\n" +
    "5. Target thresholds: Temp >= 14°C, Pressure >= 980 hPa, O2 >= 20%, Pop >= 5,000.\n\n" +
    "Good luck, Commander."
  );
}
