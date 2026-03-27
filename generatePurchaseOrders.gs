/**
 * GENERATE PURCHASE ORDER RECOMMENDATIONS (AMV)
 *
 * Key behavior:
 * - Trigger floor uses Safety Stock + following week's demand.
 * - Keeps ideal frequency fill when feasible.
 * - Prioritizes avoiding max breach over frequency fill.
 * - If both preferred/service breach max, attempts max-compliant split quantity.
 * - Enforces lead-time-feasible arrival timing.
 * - Writes PO output in A:F:
 *   A SKU, B Name, C Order Week, D Qty, E Arrival Week, F Comments
 * - Writes additional projection debug rows for Box flower M.
 */
function generatePurchaseOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- CONFIGURATION ---
  const ARRIVAL_BUFFER_WEEKS = 1;
  const SINGLE_LOCATION_NAME = "AMV DC";
  const DEBUG_TARGET_SKU_NAME = "box flower m"; // case-insensitive contains

  const INV_HEADER_ROW = 6;
  const INV_DATA_START_ROW = 7;
  const CURRENT_WEEK_CELL = "B2";
  const INV_SKU_COL = 1;   // Column A
  const INV_LABEL_COL = 3; // Column C

  // --- SHEETS ---
  const settingsSheet = ss.getSheetByName("📋 Product Settings");
  const invSheet = ss.getSheetByName("📦🔮 Inventory Planner - Detailed View");
  const outputSheet = ss.getSheetByName("🛍️ Purchase Order Recommendations");

  let traceSheet = ss.getSheetByName("🐞 Debug Trace");
  if (!traceSheet) traceSheet = ss.insertSheet("🐞 Debug Trace");
  else traceSheet.clear();
  traceSheet.appendRow([
    "Location", "SKU", "Week", "Source Col Index", "Start Stock", "Inbound", "Outbound",
    "Closing (Pre-Order)", "Deficit", "Order Qty", "Strategy Used"
  ]);

  let projectionDebugSheet = ss.getSheetByName("🧪 Debug Projection");
  if (!projectionDebugSheet) projectionDebugSheet = ss.insertSheet("🧪 Debug Projection");
  else projectionDebugSheet.clear();
  projectionDebugSheet.appendRow([
    "Location", "SKU", "Name", "Week",
    "Start Stock", "Inbound (incl Rec PO)", "Outbound", "Adjustment",
    "Closing Pre-Order", "Required Floor",
    "Service Deficit", "Frequency Deficit",
    "Qty Service", "Qty Frequency", "Qty Chosen",
    "Arrival Week", "Arrival Stock (No PO)", "Arrival Headroom", "Arrival Max-Compliant Qty",
    "Baseline Peak", "Preferred Peak", "Service Peak", "Chosen Peak",
    "Max Stock", "Strategy", "Comments"
  ]);

  if (!settingsSheet || !invSheet || !outputSheet) {
    SpreadsheetApp.getUi().alert("Error: One or more required sheets are missing.");
    return;
  }

  // --- HELPERS ---
  const normalizeToken = (val) => String(val || "").trim().toLowerCase();
  const normalizeCompact = (val) => normalizeToken(val).replace(/[^a-z0-9]/g, "");
  const clean = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const parseNum = (val) => {
    if (typeof val === "number") return val;
    if (!val) return 0;
    const n = Number(String(val).replace(/[^0-9.-]/g, ""));
    return isNaN(n) ? 0 : n;
  };

  // --- 1) READ SETTINGS (SKU-first for single AMV file) ---
  const settingsSkuMap = new Map();
  const settingsSkuLooseMap = new Map();

  const settingsData = settingsSheet.getRange(3, 1, Math.max(0, settingsSheet.getLastRow() - 2), 13).getValues();
  const layoutA = { sku: 0, name: 1, lead: 2, safety: 3, eol: 4, freq: 5, moq: 6, max: 7, type: 8, caseSize: 9, pallet: 10 };
  const layoutB = { sku: 1, name: 2, lead: 3, safety: 4, eol: 5, freq: 6, moq: 7, max: 8, type: 9, caseSize: 10, pallet: 11 };
  const scoreLayout = (row, layout) => {
    let score = 0;
    const skuRaw = row[layout.sku];
    if (String(skuRaw || "").trim() !== "") score += 3;
    if (String(row[layout.name] || "").trim() !== "") score += 1;
    const lead = row[layout.lead];
    if (typeof lead === "number" || !isNaN(Number(String(lead || "").replace(/[^0-9.-]/g, "")))) score += 1;
    const type = normalizeToken(row[layout.type]);
    if (type === "" || type === "case" || type === "unit") score += 1;
    return score;
  };
  const buildProductConfig = (row, layout) => ({
    skuName: row[layout.name],
    leadTimeDays: Number(row[layout.lead]) || 0,
    safetyStock: Number(row[layout.safety]) || 0,
    lastWeekPlanned: row[layout.eol],
    orderFreqDays: Number(row[layout.freq]) || 0,
    moq: Number(row[layout.moq]) || 0,
    maxStock: parseNum(row[layout.max]),
    orderType: String(row[layout.type] || "").toLowerCase(),
    caseSize: Number(row[layout.caseSize]) || 0,
    unitsPerPallet: row[layout.pallet]
  });
  for (let r = 0; r < settingsData.length; r++) {
    const row = settingsData[r];
    const aScore = scoreLayout(row, layoutA);
    const bScore = scoreLayout(row, layoutB);
    const chosen = bScore > aScore ? layoutB : layoutA;
    const skuRaw = row[chosen.sku];
    if (!skuRaw) continue;
    const sku = normalizeToken(skuRaw);
    const skuLoose = normalizeCompact(skuRaw);
    const productConfig = buildProductConfig(row, chosen);
    if (!settingsSkuMap.has(sku)) settingsSkuMap.set(sku, productConfig);
    if (!settingsSkuLooseMap.has(skuLoose)) settingsSkuLooseMap.set(skuLoose, productConfig);
  }

  // --- 2) MAP WEEK HEADERS ---
  const invLastRow = invSheet.getLastRow();
  const invLastCol = invSheet.getLastColumn();
  const headerRowVals = invSheet.getRange(INV_HEADER_ROW, 1, 1, invLastCol).getValues()[0];
  const weekColMap = new Map();
  const weekHeaders = [];
  const weekRegex = /^\d{4}-W\d{2}$/;

  for (let c = 0; c < headerRowVals.length; c++) {
    const val = String(headerRowVals[c] || "").trim();
    if (weekRegex.test(val) && !weekColMap.has(val)) {
      weekColMap.set(val, c);
      weekHeaders.push(val);
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
  const invData = invSheet.getRange(
    INV_DATA_START_ROW, 1, Math.max(0, invLastRow - (INV_DATA_START_ROW - 1)), invLastCol
  ).getValues();

  // --- 3) SCAN SKU BLOCKS ---
  const poRecommendations = [];
  const traceData = [];
  const projectionDebugData = [];
  const scanStats = {
    blocksScanned: 0, matchedSettings: 0, missingSettings: 0,
    missingPredictedRow: 0, triggerChecks: 0, orderTriggers: 0
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

    // Gather contiguous block for same SKU.
    const blockRows = [];
    while (currentRowIndex < invData.length) {
      const nextRow = invData[currentRowIndex];
      const nextNsid = String(nextRow[INV_SKU_COL - 1] || "").trim();
      if (blockRows.length > 0 && nextNsid !== "" && nextNsid !== nsid) break;
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

    // Identify logic rows from label column C.
    let rowPredicted, rowInboundDue, rowInboundRec, rowForecast, rowTransfers, rowAdj, rowClosing;
    const getLabel = (row) => clean(row[INV_LABEL_COL - 1]);
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

    // Starting stock comes from simulation start week (fallback to current week).
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
      return Math.max(0, parseNum(rowForecast[dCol]) + parseNum(rowTransfers[dCol]));
    };
    const getRequiredFloorForWeek = (weekIndex) => {
      if (eolIndex !== -1 && weekIndex >= eolIndex) return 0;
      return product.safetyStock + getDemandForWeek(weekIndex + 1);
    };

    const applyOrderConstraints = (rawQty) => {
      let qty = Math.max(0, rawQty);
      if (product.orderType === "case" && product.caseSize > 0) qty = Math.ceil(qty / product.caseSize) * product.caseSize;
      else qty = Math.ceil(qty);
      if (qty < product.moq) {
        qty = product.moq;
        if (product.orderType === "case" && product.caseSize > 0) qty = Math.ceil(qty / product.caseSize) * product.caseSize;
      }
      return qty;
    };
    const getMaxCompliantQty = (maxAllowedQty) => {
      if (!isFinite(maxAllowedQty) || maxAllowedQty <= 0) return 0;
      let qty = 0;
      if (product.orderType === "case" && product.caseSize > 0) qty = Math.floor(maxAllowedQty / product.caseSize) * product.caseSize;
      else qty = Math.floor(maxAllowedQty);
      if (qty <= 0) return 0;
      if (product.moq > 0 && qty < product.moq) return 0;
      return qty;
    };

    // Carries future recommended arrivals into simulation.
    const recommendedInbound = new Array(weekHeaders.length).fill(0);
    const simulatePeak = (startWeekIdx, preOrderClosing, orderQty, arrivalIdx, horizonWeeks) => {
      // If planned arrival is this week or already in the past, stock is effectively injected now.
      const injectAtStart = arrivalIdx !== -1 && arrivalIdx <= startWeekIdx;
      let temp = preOrderClosing + (injectAtStart ? orderQty : 0);
      let peak = temp;
      const maxF = Math.max(1, horizonWeeks);
      for (let f = 1; f < maxF; f++) {
        const idx = startWeekIdx + f;
        if (idx >= weekHeaders.length) break;
        const col = weekColMap.get(weekHeaders[idx]);
        const inQty =
          parseNum(rowInboundDue[col]) +
          parseNum(rowInboundRec[col]) +
          parseNum(recommendedInbound[idx]) +
          (arrivalIdx === idx ? orderQty : 0);
        const outQty = parseNum(rowForecast[col]) + parseNum(rowTransfers[col]);
        const aQty = parseNum(rowAdj[col]);
        temp = temp + inQty - outQty + aQty;
        if (temp > peak) peak = temp;
      }
      return peak;
    };

    const skuNameNorm = normalizeToken(product.skuName);
    const skuCodeNorm = normalizeToken(nsidRaw);
    const isDebugSku = skuNameNorm.includes(DEBUG_TARGET_SKU_NAME) || skuCodeNorm.includes(normalizeToken(DEBUG_TARGET_SKU_NAME));

    // --- WEEKLY SIMULATION ---
    for (let w = simulationStartIndex; w < weekHeaders.length; w++) {
      const thisWeekStr = weekHeaders[w];
      const colIndex = weekColMap.get(thisWeekStr);

      const inbound = parseNum(rowInboundDue[colIndex]) + parseNum(rowInboundRec[colIndex]) + parseNum(recommendedInbound[w]);
      const outbound = parseNum(rowForecast[colIndex]) + parseNum(rowTransfers[colIndex]);
      const adj = parseNum(rowAdj[colIndex]);

      const startStockForWeek = runningInventory;
      let weekClosing = runningInventory + inbound - outbound + adj;
      let qtyToOrder = 0;
      let deficit = 0;
      let strategy = "";
      let requiredFloorForLog = getRequiredFloorForWeek(w);
      let hardServiceDeficitForLog = 0;
      let preferredDeficitForLog = 0;
      let qtyServiceForLog = 0;
      let qtyPreferredForLog = 0;
      let arrivalWeekForLog = "";
      let baselinePeakForLog = 0;
      let preferredPeakForLog = 0;
      let servicePeakForLog = 0;
      let chosenPeakForLog = 0;
      let arrivalStockNoPoForLog = 0;
      let arrivalHeadroomForLog = 0;
      let arrivalMaxCompliantQtyForLog = 0;
      let commentsForLog = "";

      if (w >= earliestArrivalIndex) {
        scanStats.triggerChecks++;
        const requiredFloor = getRequiredFloorForWeek(w);
        requiredFloorForLog = requiredFloor;

        if (weekClosing < requiredFloor) {
          scanStats.orderTriggers++;
          deficit = requiredFloor - weekClosing;
          strategy = "Survival + Next Week Cover";

          let cycleDeficit = deficit;
          if (freqWeeksToCover > 1 && (eolIndex === -1 || w < eolIndex)) {
            let tempRunning = weekClosing;
            for (let f = 1; f < freqWeeksToCover; f++) {
              if (w + f >= weekHeaders.length) break;
              const fCol = weekColMap.get(weekHeaders[w + f]);
              const fIn = parseNum(rowInboundDue[fCol]) + parseNum(rowInboundRec[fCol]);
              const fOut = parseNum(rowForecast[fCol]) + parseNum(rowTransfers[fCol]);
              const fAdj = parseNum(rowAdj[fCol]);
              tempRunning = tempRunning + fIn - fOut + fAdj;
              const futureFloor = getRequiredFloorForWeek(w + f);
              const needed = futureFloor - tempRunning;
              if (needed > cycleDeficit) cycleDeficit = needed;
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
          hardServiceDeficitForLog = hardServiceDeficit;
          preferredDeficitForLog = preferredCoverageDeficit;
          qtyServiceForLog = qtyForHardService;
          qtyPreferredForLog = qtyForPreferredCoverage;
          qtyToOrder = qtyForPreferredCoverage;

          const commentParts = [];

          // Timing
          const shortageDate = weekDates[w];
          const targetArrivalDate = new Date(shortageDate);
          targetArrivalDate.setDate(targetArrivalDate.getDate() - ARRIVAL_BUFFER_WEEKS * 7);
          const earliestFeasibleArrivalDate = new Date(currentDateObj);
          earliestFeasibleArrivalDate.setDate(earliestFeasibleArrivalDate.getDate() + product.leadTimeDays);
          const plannedArrivalDate = new Date(targetArrivalDate);
          if (plannedArrivalDate < earliestFeasibleArrivalDate) {
            plannedArrivalDate.setTime(earliestFeasibleArrivalDate.getTime());
            commentParts.push("Lead Time Constraint (Arrival Shifted Later)");
          }
          const orderDate = new Date(plannedArrivalDate);
          orderDate.setDate(orderDate.getDate() - product.leadTimeDays);
          if (orderDate < currentDateObj) orderDate.setTime(currentDateObj.getTime());

          const orderWeekStr = getIsoWeekString(orderDate);
          const arrivalWeekStr = getIsoWeekString(plannedArrivalDate);
          const plannedArrivalIndex = weekHeaders.indexOf(arrivalWeekStr);
          arrivalWeekForLog = arrivalWeekStr;

          const simulateStockAtWeekNoPo = (startWeekIdx, preOrderClosing, targetWeekIdx) => {
            let temp = preOrderClosing;
            if (targetWeekIdx <= startWeekIdx) return temp;
            for (let f = 1; f <= targetWeekIdx - startWeekIdx; f++) {
              const idx = startWeekIdx + f;
              if (idx >= weekHeaders.length) break;
              const col = weekColMap.get(weekHeaders[idx]);
              const inQty =
                parseNum(rowInboundDue[col]) +
                parseNum(rowInboundRec[col]) +
                parseNum(recommendedInbound[idx]);
              const outQty = parseNum(rowForecast[col]) + parseNum(rowTransfers[col]);
              const aQty = parseNum(rowAdj[col]);
              temp = temp + inQty - outQty + aQty;
            }
            return temp;
          };
          const arrivalStockNoPo =
            plannedArrivalIndex >= 0 ? simulateStockAtWeekNoPo(w, weekClosing, plannedArrivalIndex) : weekClosing;
          const arrivalHeadroom =
            product.maxStock > 0 ? product.maxStock - arrivalStockNoPo : Number.POSITIVE_INFINITY;
          const arrivalMaxCompliantQty = getMaxCompliantQty(arrivalHeadroom);
          arrivalStockNoPoForLog = arrivalStockNoPo;
          arrivalHeadroomForLog = arrivalHeadroom;
          arrivalMaxCompliantQtyForLog = arrivalMaxCompliantQty;

          // Max-priority decisioning
          const horizonWeeks = weekHeaders.length - w;
          const baselinePeak = simulatePeak(w, weekClosing, 0, -1, horizonWeeks);
          const preferredPeak = simulatePeak(w, weekClosing, qtyForPreferredCoverage, plannedArrivalIndex, horizonWeeks);
          const servicePeak = simulatePeak(w, weekClosing, qtyForHardService, plannedArrivalIndex, horizonWeeks);
          baselinePeakForLog = baselinePeak;
          preferredPeakForLog = preferredPeak;
          servicePeakForLog = servicePeak;

          if (product.maxStock > 0 && preferredPeak > product.maxStock) {
            if (servicePeak <= product.maxStock) {
              qtyToOrder = qtyForHardService;
              strategy = `${strategy} (Soft Capped)`;
              commentParts.push("Soft Capped to Max (Frequency Fill Reduced)");
            } else {
              // Never sacrifice service floor. If service itself breaches max, keep service qty and annotate.
              qtyToOrder = qtyForHardService;
              const breachReasons = [];
              if (baselinePeak > product.maxStock) breachReasons.push("Baseline Over Max");
              if (hardServiceDeficit > arrivalHeadroom) breachReasons.push("Coverage");
              if (servicePeak > product.maxStock && hardServiceDeficit <= arrivalHeadroom) breachReasons.push("Projected Peak");
              if (product.orderType === "case" && product.caseSize > 0 && hardServiceDeficit <= arrivalHeadroom) {
                const caseRounded = Math.ceil(Math.max(0, hardServiceDeficit) / product.caseSize) * product.caseSize;
                if (caseRounded > arrivalHeadroom) breachReasons.push("Case Pack");
              }
              if (product.moq > 0) {
                const moqRounded = applyOrderConstraints(product.moq);
                if (moqRounded > arrivalHeadroom) breachReasons.push("MOQ");
              }
              const uniqueReasons = [...new Set(breachReasons)];
              if (uniqueReasons.length > 0) commentParts.push(`Max Capacity Breached (Unavoidable: ${uniqueReasons.join(", ")})`);
              else commentParts.push("Max Capacity Breached");
            }
          }

          const chosenPeak = simulatePeak(w, weekClosing, qtyToOrder, plannedArrivalIndex, horizonWeeks);
          chosenPeakForLog = chosenPeak;
          if (
            product.maxStock > 0 &&
            chosenPeak > product.maxStock &&
            !commentParts.some((c) => c.startsWith("Max Capacity Breached"))
          ) {
            commentParts.push("Max Capacity Breached");
          }

          commentsForLog = commentParts.join("; ");

          poRecommendations.push([
            nsidRaw,
            product.skuName,
            orderWeekStr,
            qtyToOrder,
            arrivalWeekStr,
            commentsForLog
          ]);

          // Apply order at arrival timing in simulation ledger.
          if (plannedArrivalIndex === -1) {
            commentParts.push("Arrival Beyond Planning Horizon");
          } else if (plannedArrivalIndex <= w) {
            weekClosing += qtyToOrder;
          } else {
            recommendedInbound[plannedArrivalIndex] += qtyToOrder;
          }
        }
      }

      if (isDebugSku) {
        projectionDebugData.push([
          SINGLE_LOCATION_NAME,
          nsidRaw,
          product.skuName,
          thisWeekStr,
          Math.round(startStockForWeek),
          inbound,
          outbound,
          adj,
          Math.round(weekClosing),
          Math.round(requiredFloorForLog),
          Math.round(hardServiceDeficitForLog),
          Math.round(preferredDeficitForLog),
          Math.round(qtyServiceForLog),
          Math.round(qtyPreferredForLog),
          Math.round(qtyToOrder),
          arrivalWeekForLog,
          Math.round(arrivalStockNoPoForLog),
          Math.round(arrivalHeadroomForLog),
          Math.round(arrivalMaxCompliantQtyForLog),
          Math.round(baselinePeakForLog),
          Math.round(preferredPeakForLog),
          Math.round(servicePeakForLog),
          Math.round(chosenPeakForLog),
          product.maxStock,
          strategy,
          commentsForLog
        ]);
      }

      // Existing targeted trace
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

  // --- 4) WRITE OUTPUTS ---
  const lastRow = outputSheet.getLastRow();
  if (lastRow > 1) outputSheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  if (poRecommendations.length > 0) {
    poRecommendations.sort((a, b) => (a[2] < b[2] ? -1 : 1)); // by order week
    outputSheet.getRange(2, 1, poRecommendations.length, 6).setValues(poRecommendations);
  }

  if (traceData.length > 0) {
    traceSheet.getRange(2, 1, traceData.length, 11).setValues(traceData);
  }
  if (projectionDebugData.length > 0) {
    projectionDebugSheet.getRange(2, 1, projectionDebugData.length, 26).setValues(projectionDebugData);
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
