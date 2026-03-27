/**
 * GENERATE PURCHASE ORDER RECOMMENDATIONS
 *
 * Update: Enforces "Cycle Coverage". If an order is triggered, it ensures the quantity
 * is sufficient to last the full 'Ideal Order Frequency' period.
 *
 * Update: The weekly minimum floor now targets:
 *   Safety Stock + following week's demand
 * instead of Safety Stock alone.
 *
 * Update: Soft-cap optimizer for max stock:
 *   1) Keep service floor as highest priority.
 *   2) Try to reduce to a max-compliant quantity when possible.
 *   3) If breach is unavoidable, still recommend and annotate why.
 */
function generatePurchaseOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- CONFIGURATION ---
  const ARRIVAL_BUFFER_WEEKS = 1; // Stock must arrive this many weeks BEFORE the shortage
  const SINGLE_LOCATION_MODE = true; // This planner serves one DC/location; match settings primarily by SKU.
  const SINGLE_LOCATION_NAME = "AMV DC";
  const INV_HEADER_ROW = 6;
  const INV_DATA_START_ROW = 7;
  const CURRENT_WEEK_CELL = "B2";
  const INV_SKU_COL = 1; // Column A
  const INV_LABEL_COL = 3; // Column C

  // --- SHEETS ---
  const settingsSheet = ss.getSheetByName("📋 Product Settings");
  const invSheet = ss.getSheetByName("📦🔮 Inventory Planner - Detailed View");
  let outputSheet = ss.getSheetByName("🛍️ Purchase Order Recommendations");

  let traceSheet = ss.getSheetByName("🐞 Debug Trace");
  if (!traceSheet) {
    traceSheet = ss.insertSheet("🐞 Debug Trace");
  } else {
    traceSheet.clear();
  }
  traceSheet.appendRow(["Location", "SKU", "Week", "Source Col Index", "Start Stock", "Inbound", "Outbound", "Closing (Pre-Order)", "Deficit", "Order Qty", "Strategy Used"]);

  if (!settingsSheet || !invSheet || !outputSheet) {
    SpreadsheetApp.getUi().alert("Error: One or more required sheets are missing.");
    return;
  }

  // --- 1. READ SETTINGS ---
  const settingsSkuMap = new Map();
  const settingsSkuLooseMap = new Map();
  const normalizeToken = (val) => String(val || "").trim().toLowerCase();
  const normalizeCompact = (val) => normalizeToken(val).replace(/[^a-z0-9]/g, "");
  const parseNum = (val) => {
    if (typeof val === "number") return val;
    if (!val) return 0;
    const n = Number(String(val).replace(/[^0-9.-]/g, ""));
    return isNaN(n) ? 0 : n;
  };

  const settingsData = settingsSheet.getRange(3, 1, settingsSheet.getLastRow() - 2, 13).getValues();
  const buildProductConfig = (row, offset) => ({
    skuName: row[offset + 1],
    leadTimeDays: Number(row[offset + 2]) || 0,
    safetyStock: Number(row[offset + 3]) || 0,
    lastWeekPlanned: row[offset + 4],
    orderFreqDays: Number(row[offset + 5]) || 0,
    moq: Number(row[offset + 6]) || 0,
    maxStock: parseNum(row[offset + 7]),
    orderType: String(row[offset + 8]).toLowerCase(),
    caseSize: Number(row[offset + 9]) || 0,
    unitsPerPallet: row[offset + 10]
  });
  const addSettingsVariant = (row, offset, overwrite) => {
    const skuRaw = row[offset];
    if (!skuRaw) return false;
    const sku = normalizeToken(skuRaw);
    const skuLoose = normalizeCompact(skuRaw);
    const productConfig = buildProductConfig(row, offset);

    if (overwrite || !settingsSkuMap.has(sku)) settingsSkuMap.set(sku, productConfig);
    if (overwrite || !settingsSkuLooseMap.has(skuLoose)) settingsSkuLooseMap.set(skuLoose, productConfig);
    return true;
  };

  for (let r = 0; r < settingsData.length; r++) {
    const row = settingsData[r];
    // Supports both layouts:
    //  - single-location: SKU in col A (offset 0)
    //  - legacy layout:   SKU in col B (offset 1)
    if (SINGLE_LOCATION_MODE) {
      const addedPrimary = addSettingsVariant(row, 0, true);
      if (!addedPrimary) addSettingsVariant(row, 1, false);
    } else {
      const addedPrimary = addSettingsVariant(row, 1, true);
      if (!addedPrimary) addSettingsVariant(row, 0, false);
    }
  }

  // --- 2. MAP HEADERS ---
  const invLastRow = invSheet.getLastRow();
  const invLastCol = invSheet.getLastColumn();
  const headerRowVals = invSheet.getRange(INV_HEADER_ROW, 1, 1, invLastCol).getValues()[0];
  const weekColMap = new Map();
  const weekHeaders = [];
  const weekRegex = /^\d{4}-W\d{2}$/;

  for (let c = 0; c < headerRowVals.length; c++) {
    const val = String(headerRowVals[c]).trim();
    if (weekRegex.test(val)) {
      if (!weekColMap.has(val)) {
        weekColMap.set(val, c);
        weekHeaders.push(val);
      }
    }
  }

  if (weekHeaders.length === 0) {
    SpreadsheetApp.getUi().alert("Error: No 'YYYY-WXX' headers found in Row 6.");
    return;
  }

  const weekDates = weekHeaders.map((w) => getDateFromIsoWeek(w));

  let currentWeekStr = invSheet.getRange(CURRENT_WEEK_CELL).getValue();
  if (!currentWeekStr || typeof currentWeekStr !== "string") currentWeekStr = getIsoWeekString(new Date());

  const currentWeekIndex = weekHeaders.indexOf(currentWeekStr);
  if (currentWeekIndex === -1) {
    SpreadsheetApp.getUi().alert(`Error: Current week ${currentWeekStr} not found in headers.`);
    return;
  }

  const simulationStartIndex = currentWeekIndex + 1;
  if (simulationStartIndex >= weekHeaders.length) {
    SpreadsheetApp.getUi().alert("Error: No future weeks found after " + currentWeekStr);
    return;
  }

  const currentDateObj = getDateFromIsoWeek(currentWeekStr);
  const invData = invSheet.getRange(INV_DATA_START_ROW, 1, invLastRow - (INV_DATA_START_ROW - 1), invLastCol).getValues();

  // --- 3. DYNAMIC SCANNING ---
  const poRecommendations = [];
  const traceData = [];
  const scanStats = {
    blocksScanned: 0,
    matchedSettings: 0,
    missingSettings: 0,
    missingPredictedRow: 0,
    triggerChecks: 0,
    orderTriggers: 0
  };
  let currentRowIndex = 0;

  while (currentRowIndex < invData.length) {
    const firstRow = invData[currentRowIndex];
    const nsidRaw = firstRow[INV_SKU_COL - 1];

    if (!nsidRaw) {
      currentRowIndex++;
      continue;
    }

    const nsid = String(nsidRaw).trim();

    // Collect block rows
    const blockRows = [];
    while (currentRowIndex < invData.length) {
      const nextRow = invData[currentRowIndex];
      const nextNsid = String(nextRow[INV_SKU_COL - 1]).trim();

      if (blockRows.length > 0 && nextNsid !== "") {
        if (nextNsid !== nsid) break;
      }
      blockRows.push(nextRow);
      currentRowIndex++;
    }

    scanStats.blocksScanned++;
    const product = settingsSkuMap.get(normalizeToken(nsid)) || settingsSkuLooseMap.get(normalizeCompact(nsid));
    if (!product) {
      scanStats.missingSettings++;
      continue;
    }
    scanStats.matchedSettings++;

    // Find Logic Rows
    let rowPredicted, rowInboundDue, rowInboundRec, rowForecast, rowTransfers, rowAdj, rowClosing;
    const clean = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, "");
    const getLabel = (row) => {
      const raw = row[INV_LABEL_COL - 1];
      return clean(raw);
    };

    blockRows.forEach((row) => {
      const label = getLabel(row);
      if (!label) return;

      const isStartingStock = label.includes("startingstock") && !label.includes("closingstock");
      if (!rowPredicted && (label.includes("predictedstartingstock") || isStartingStock || label.includes("predicted"))) {
        rowPredicted = row;
      } else if (!rowInboundDue && ((label.includes("inbound") && label.includes("due")) || label.includes("dueinbound") || label.includes("openpo") || label.includes("onorder"))) {
        rowInboundDue = row;
      } else if (!rowInboundRec && ((label.includes("inbound") && label.includes("received")) || label.includes("inboundreceived") || label === "received")) {
        rowInboundRec = row;
      } else if (!rowForecast && label.includes("forecast")) {
        rowForecast = row;
      } else if (!rowTransfers && label.includes("transfer")) {
        rowTransfers = row;
      } else if (!rowAdj && label.includes("adjust")) {
        rowAdj = row;
      } else if (!rowClosing && label.includes("closingstock")) {
        rowClosing = row;
      }
    });
    if (!rowPredicted && rowClosing) rowPredicted = rowClosing;

    const zeroRow = new Array(invLastCol).fill(0);
    if (!rowInboundDue) rowInboundDue = zeroRow;
    if (!rowInboundRec) rowInboundRec = zeroRow;
    if (!rowForecast) rowForecast = zeroRow;
    if (!rowTransfers) rowTransfers = zeroRow;
    if (!rowAdj) rowAdj = zeroRow;

    // Initialize Ledger
    const startWeekHeader = weekHeaders[simulationStartIndex];
    const startCol = weekColMap.get(startWeekHeader);
    let rawPred = rowPredicted ? rowPredicted[startCol] : "";
    if (String(rawPred).trim() === "" && rowPredicted) {
      const currentCol = weekColMap.get(currentWeekStr);
      if (currentCol !== undefined) rawPred = rowPredicted[currentCol];
    }
    if (String(rawPred).trim() === "") {
      scanStats.missingPredictedRow++;
      continue;
    }

    let runningInventory = parseNum(rawPred);

    let eolIndex = -1;
    if (product.lastWeekPlanned && product.lastWeekPlanned !== "N/A") {
      eolIndex = weekHeaders.indexOf(product.lastWeekPlanned);
    }

    const leadTimeWeeks = Math.ceil(product.leadTimeDays / 7);
    const earliestArrivalIndex = currentWeekIndex + leadTimeWeeks;
    const freqWeeksToCover = product.orderFreqDays > 7 ? Math.ceil(product.orderFreqDays / 7) : 1;

    const getDemandForWeek = (weekIndex) => {
      if (weekIndex < 0 || weekIndex >= weekHeaders.length) return 0;
      const dCol = weekColMap.get(weekHeaders[weekIndex]);
      const demand = parseNum(rowForecast[dCol]) + parseNum(rowTransfers[dCol]);
      return Math.max(0, demand);
    };

    const getRequiredFloorForWeek = (weekIndex) => {
      if (eolIndex !== -1 && weekIndex >= eolIndex) return 0;
      const safety = product.safetyStock;
      const nextWeekDemand = getDemandForWeek(weekIndex + 1);
      return safety + nextWeekDemand;
    };

    const applyOrderConstraints = (rawQty) => {
      let qty = Math.max(0, rawQty);
      if (product.orderType === "case" && product.caseSize > 0) {
        qty = Math.ceil(qty / product.caseSize) * product.caseSize;
      } else {
        qty = Math.ceil(qty);
      }

      if (qty < product.moq) {
        qty = product.moq;
        if (product.orderType === "case" && product.caseSize > 0) {
          qty = Math.ceil(qty / product.caseSize) * product.caseSize;
        }
      }
      return qty;
    };

    // --- SIMULATION LOOP ---
    for (let w = simulationStartIndex; w < weekHeaders.length; w++) {
      const thisWeekStr = weekHeaders[w];
      const colIndex = weekColMap.get(thisWeekStr);

      const inbound = parseNum(rowInboundDue[colIndex]) + parseNum(rowInboundRec[colIndex]);
      const outbound = parseNum(rowForecast[colIndex]) + parseNum(rowTransfers[colIndex]);
      const adj = parseNum(rowAdj[colIndex]);

      const startStockForWeek = runningInventory;
      let weekClosing = runningInventory + inbound - outbound + adj;
      let qtyToOrder = 0;
      let deficit = 0;
      let strategy = "";

      if (w >= earliestArrivalIndex) {
        scanStats.triggerChecks++;
        // Trigger condition now uses Safety + following week's demand.
        const requiredFloor = getRequiredFloorForWeek(w);

        if (weekClosing < requiredFloor) {
          scanStats.orderTriggers++;
          // 1) Immediate weekly survival floor (with next-week cover)
          deficit = requiredFloor - weekClosing;
          strategy = "Survival + Next Week Cover";

          // 2) Frequency logic: ensure floor is preserved through cycle
          let cycleDeficit = deficit;

          if (freqWeeksToCover > 1 && (eolIndex === -1 || w < eolIndex)) {
            let tempRunning = weekClosing;

            for (let f = 1; f < freqWeeksToCover; f++) {
              if (w + f < weekHeaders.length) {
                const fIndex = weekColMap.get(weekHeaders[w + f]);
                const fIn = parseNum(rowInboundDue[fIndex]) + parseNum(rowInboundRec[fIndex]);
                const fOut = parseNum(rowForecast[fIndex]) + parseNum(rowTransfers[fIndex]);
                const fAdj = parseNum(rowAdj[fIndex]);

                tempRunning = tempRunning + fIn - fOut + fAdj;

                const futureRequiredFloor = getRequiredFloorForWeek(w + f);
                const neededAtFutureWeek = futureRequiredFloor - tempRunning;
                if (neededAtFutureWeek > cycleDeficit) {
                  cycleDeficit = neededAtFutureWeek;
                }
              }
            }

            if (cycleDeficit > deficit) {
              deficit = cycleDeficit;
              strategy = "Frequency Fill + Next Week Cover";
            }
          }

          const hardServiceDeficit = requiredFloor - weekClosing;
          const preferredCoverageDeficit = deficit;
          const qtyForHardService = applyOrderConstraints(hardServiceDeficit);
          const qtyForPreferredCoverage = applyOrderConstraints(preferredCoverageDeficit);

          qtyToOrder = qtyForPreferredCoverage;

          const commentParts = [];
          const maxAllowedQty = product.maxStock > 0 ? product.maxStock - (startStockForWeek + inbound) : Number.POSITIVE_INFINITY;

          if (isFinite(maxAllowedQty) && qtyForPreferredCoverage > maxAllowedQty) {
            if (qtyForHardService <= maxAllowedQty) {
              // Respect max by dialing back from frequency-fill to minimum service floor.
              qtyToOrder = qtyForHardService;
              strategy = `${strategy} (Soft Capped)`;
              commentParts.push("Soft Capped to Max (Frequency Fill Reduced)");
            } else {
              // Max breach is unavoidable while still protecting service floor.
              qtyToOrder = qtyForHardService;
              const breachReasons = [];

              if (hardServiceDeficit > maxAllowedQty) {
                breachReasons.push("Coverage");
              }

              if (product.orderType === "case" && product.caseSize > 0 && hardServiceDeficit <= maxAllowedQty) {
                const caseRounded = Math.ceil(Math.max(0, hardServiceDeficit) / product.caseSize) * product.caseSize;
                if (caseRounded > maxAllowedQty) breachReasons.push("Case Pack");
              }

              if (product.moq > 0) {
                const moqRounded = applyOrderConstraints(product.moq);
                if (moqRounded > maxAllowedQty) breachReasons.push("MOQ");
              }

              const uniqueReasons = [...new Set(breachReasons)];
              if (uniqueReasons.length > 0) {
                commentParts.push(`Max Capacity Breached (Unavoidable: ${uniqueReasons.join(", ")})`);
              } else {
                commentParts.push("Max Capacity Breached");
              }
            }
          }

          // Order Timing & Buffer
          const shortageDate = weekDates[w];
          const targetArrivalDate = new Date(shortageDate);
          targetArrivalDate.setDate(targetArrivalDate.getDate() - ARRIVAL_BUFFER_WEEKS * 7);

          // Arrival cannot be earlier than what lead time allows from "now".
          const earliestFeasibleArrivalDate = new Date(currentDateObj);
          earliestFeasibleArrivalDate.setDate(earliestFeasibleArrivalDate.getDate() + product.leadTimeDays);

          const plannedArrivalDate = new Date(targetArrivalDate);
          if (plannedArrivalDate < earliestFeasibleArrivalDate) {
            plannedArrivalDate.setTime(earliestFeasibleArrivalDate.getTime());
            commentParts.push("Lead Time Constraint (Arrival Shifted Later)");
          }

          const orderDate = new Date(plannedArrivalDate);
          orderDate.setDate(orderDate.getDate() - product.leadTimeDays);
          if (orderDate < currentDateObj) {
            orderDate.setTime(currentDateObj.getTime());
          }

          const orderWeekStr = getIsoWeekString(orderDate);
          const arrivalWeekStr = getIsoWeekString(plannedArrivalDate);

          // Flag max breaches using both "arrival peak" and modeled post-order closing.
          const arrivalPeakStock = startStockForWeek + inbound + qtyToOrder;
          const postOrderClosingStock = weekClosing + qtyToOrder;
          const modeledMaxPoint = Math.max(arrivalPeakStock, postOrderClosingStock);
          if (
            product.maxStock > 0 &&
            modeledMaxPoint > product.maxStock &&
            !commentParts.some((c) => c.startsWith("Max Capacity Breached"))
          ) {
            commentParts.push("Max Capacity Breached");
          }

          poRecommendations.push([
            nsidRaw,
            product.skuName,
            orderWeekStr,
            qtyToOrder,
            arrivalWeekStr,
            commentParts.join("; ")
          ]);

          weekClosing += qtyToOrder;
        }
      }

      // Trace logging for specific SKU to debug
      if (String(nsidRaw).includes("PK-11-00002")) {
        traceData.push([
          SINGLE_LOCATION_NAME,
          nsidRaw,
          thisWeekStr,
          colIndex + 1,
          Math.round(startStockForWeek),
          inbound,
          outbound,
          Math.round(weekClosing),
          Math.round(deficit),
          qtyToOrder,
          strategy
        ]);
      }

      runningInventory = weekClosing;
    }
  }

  // --- WRITE OUTPUT ---
  const lastRow = outputSheet.getLastRow();
  if (lastRow > 1) outputSheet.getRange(2, 1, lastRow - 1, 6).clearContent();

  if (poRecommendations.length > 0) {
    poRecommendations.sort((a, b) => (a[3] < b[3] ? -1 : 1));
    outputSheet.getRange(2, 1, poRecommendations.length, 6).setValues(poRecommendations);
  }

  if (traceData.length > 0) {
    traceSheet.getRange(2, 1, traceData.length, 11).setValues(traceData);
  }

  let completionMessage = `Generated ${poRecommendations.length} POs.`;
  if (poRecommendations.length === 0) {
    completionMessage += `\nDebug summary: blocks=${scanStats.blocksScanned}, matched settings=${scanStats.matchedSettings}, missing settings=${scanStats.missingSettings}, missing predicted row=${scanStats.missingPredictedRow}, trigger checks=${scanStats.triggerChecks}, order triggers=${scanStats.orderTriggers}.`;
  }
  SpreadsheetApp.getUi().alert(completionMessage);
}

// --- HELPER FUNCTIONS ---
function getDateFromIsoWeek(weekStr) {
  if (!weekStr || !weekStr.includes("-W")) return new Date();
  const parts = weekStr.split("-W");
  const year = parseInt(parts[0], 10);
  const week = parseInt(parts[1], 10);
  const simple = new Date(year, 0, 1 + (week - 1) * 7);
  const day = simple.getDay();
  const isoDate = simple;
  if (day <= 4) isoDate.setDate(simple.getDate() - simple.getDay() + 1);
  else isoDate.setDate(simple.getDate() + 8 - simple.getDay());
  return isoDate;
}

function getIsoWeekString(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  date.setUTCDate(date.getUTCDate() + 4 - (date.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return date.getUTCFullYear() + "-W" + (weekNo < 10 ? "0" + weekNo : weekNo);
}
