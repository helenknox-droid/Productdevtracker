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
 *
 * Truck-fill extension:
 * - Uses Product Settings columns:
 *   M: Truck Fill Required (checkbox)
 *   N: Supplier (group key)
 *   O: Truck Pallet Capacity
 * - For supplier-arrival-week groups that are not a full-truck multiple,
 *   pulls quantities forward from later arrival weeks for the same supplier.
 * - Keeps pack rounding and attempts to avoid leaving donor lines below MOQ.
 */
function generatePurchaseOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- CONFIGURATION ---
  const ARRIVAL_BUFFER_WEEKS = 1;
  const SINGLE_LOCATION_NAME = "AMV DC";
  const DEBUG_TARGET_SKU_NAME = "box flower m"; // case-insensitive contains
  const TRUCK_FILL_EPSILON = 1e-6;

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
    "Start Stock", "Inbound Due", "Inbound Received", "Recommended Inbound Applied", "Recommended Applied Same Week", "Inbound (incl Rec PO)", "Outbound", "Adjustment",
    "Closing Pre-Order", "Closing Post-Order", "Required Floor",
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
  const parseBool = (val) => {
    if (val === true) return true;
    const t = normalizeToken(val);
    return t === "true" || t === "yes" || t === "y" || t === "1" || t === "checked";
  };
  const appendComment = (rec, text) => {
    if (!text) return;
    rec.comments = rec.comments ? `${rec.comments}; ${text}` : text;
  };

  // --- 1) READ SETTINGS (SKU-first for single AMV file) ---
  const settingsSkuMap = new Map();
  const settingsSkuLooseMap = new Map();

  // Keep this wide enough to include M/N/O truck-fill columns.
  const settingsRowsCount = Math.max(0, settingsSheet.getLastRow() - 2);
  const settingsColsCount = Math.max(15, settingsSheet.getLastColumn());
  const settingsData =
    settingsRowsCount > 0
      ? settingsSheet.getRange(3, 1, settingsRowsCount, settingsColsCount).getValues()
      : [];

  const buildProductConfig = (row, offset) => {
    // Some AMV variants place an extra numeric field before Max Stock.
    // Detect order type position and infer max as the column immediately before it.
    const typeCandidates = [offset + 8, offset + 9, offset + 10, offset + 11];
    let typeIdx = offset + 8;
    for (let i = 0; i < typeCandidates.length; i++) {
      const idx = typeCandidates[i];
      const typeVal = normalizeToken(row[idx]);
      if (typeVal === "unit" || typeVal === "case") {
        typeIdx = idx;
        break;
      }
    }
    const maxIdx = Math.max(offset + 7, typeIdx - 1);
    const caseIdx = typeIdx + 1;
    const palletIdx = typeIdx + 2;

    // New truck-fill fields in M/N/O (offset-relative).
    const truckFillRequiredIdx = offset + 12;
    const supplierIdx = offset + 13;
    const truckPalletCapacityIdx = offset + 14;

    return {
      skuName: row[offset + 1],
      leadTimeDays: Number(row[offset + 2]) || 0,
      safetyStock: Number(row[offset + 3]) || 0,
      lastWeekPlanned: row[offset + 4],
      orderFreqDays: Number(row[offset + 5]) || 0,
      moq: Number(row[offset + 6]) || 0,
      maxStock: parseNum(row[maxIdx]),
      orderType: String(row[typeIdx] || "").toLowerCase(),
      caseSize: Number(row[caseIdx]) || 0,
      unitsPerPallet: parseNum(row[palletIdx]),
      truckFillRequired: parseBool(row[truckFillRequiredIdx]),
      supplier: String(row[supplierIdx] || "").trim(),
      truckPalletCapacity: parseNum(row[truckPalletCapacityIdx])
    };
  };
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
    // Stable AMV mapping: prefer SKU in col A, fallback to col B.
    const addedPrimary = addSettingsVariant(row, 0, true);
    if (!addedPrimary) addSettingsVariant(row, 1, false);
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
  const weekIndexMap = new Map(weekHeaders.map((w, i) => [w, i]));

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
  let poRecommendations = [];
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
    const labelColCandidates = [...new Set([INV_LABEL_COL - 1, 3])].filter((i) => i >= 0 && i <= 3);
    const getLabel = (row) => {
      for (let i = 0; i < labelColCandidates.length; i++) {
        const raw = row[labelColCandidates[i]];
        if (String(raw || "").trim() !== "") return clean(raw);
      }
      return "";
    };
    blockRows.forEach((row) => {
      const label = getLabel(row);
      if (!label) return;
      const isStartingStock = label.includes("startingstock") && !label.includes("closingstock");
      if (!rowPredicted && (label.includes("predictedstartingstock") || isStartingStock || label.includes("predicted"))) {
        rowPredicted = row;
      } else if (!rowInboundDue && (
        label.includes("inbounddue") ||
        label.includes("dueinbound") ||
        label.includes("openpo") ||
        label.includes("onorder") ||
        label.includes("podue") ||
        label.includes("inboundpo") ||
        (label.includes("inbound") && label.includes("due"))
      )) {
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

      const inboundDueVal = parseNum(rowInboundDue[colIndex]);
      const inboundRecVal = parseNum(rowInboundRec[colIndex]);
      const recInboundVal = parseNum(recommendedInbound[w]);
      const inbound = inboundDueVal + inboundRecVal + recInboundVal;
      const outbound = parseNum(rowForecast[colIndex]) + parseNum(rowTransfers[colIndex]);
      const adj = parseNum(rowAdj[colIndex]);

      const startStockForWeek = runningInventory;
      let weekClosing = runningInventory + inbound - outbound + adj;
      const closingPreOrderForLog = weekClosing;
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
      let closingPostOrderForLog = weekClosing;
      let inboundDueForLog = inboundDueVal;
      let inboundRecForLog = inboundRecVal;
      let recInboundForLog = recInboundVal;
      let recAppliedSameWeekForLog = 0;
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
          const preferredArrivalStock =
            plannedArrivalIndex >= 0 ? arrivalStockNoPo + qtyForPreferredCoverage : weekClosing + qtyForPreferredCoverage;
          const serviceArrivalStock =
            plannedArrivalIndex >= 0 ? arrivalStockNoPo + qtyForHardService : weekClosing + qtyForHardService;
          baselinePeakForLog = baselinePeak;
          preferredPeakForLog = preferredPeak;
          servicePeakForLog = servicePeak;

          if (product.maxStock > 0) {
            const preferredBreachesMax = preferredArrivalStock > product.maxStock || preferredPeak > product.maxStock;
            const serviceBreachesMax = serviceArrivalStock > product.maxStock || servicePeak > product.maxStock;

            // Trade-off should be max vs frequency (not service): drop to service when preferred breaches.
            if (qtyForPreferredCoverage > qtyForHardService && preferredBreachesMax) {
              qtyToOrder = qtyForHardService;
              strategy = `${strategy} (Frequency Sacrificed for Max)`;
              commentParts.push("Frequency Sacrificed for Max");
            }

            if (serviceBreachesMax) {
              // Never sacrifice service floor. If service itself breaches max, keep service qty and annotate.
              qtyToOrder = qtyForHardService;
              const breachReasons = [];
              if (baselinePeak > product.maxStock) breachReasons.push("Baseline Over Max");
              if (hardServiceDeficit > arrivalHeadroom) breachReasons.push("Coverage");
              if (serviceArrivalStock > product.maxStock) breachReasons.push("Arrival Stock");
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
          const chosenArrivalStock =
            plannedArrivalIndex >= 0 ? arrivalStockNoPo + qtyToOrder : weekClosing + qtyToOrder;
          chosenPeakForLog = chosenPeak;
          if (product.maxStock > 0 && (chosenPeak > product.maxStock || chosenArrivalStock > product.maxStock)) {
            if (!commentParts.some((c) => c.startsWith("Max Capacity Breached"))) {
              commentParts.push("Max Capacity Breached");
            }
            if ((preferredPeak > product.maxStock || preferredArrivalStock > product.maxStock) && qtyToOrder === qtyForHardService) {
              if (!commentParts.includes("Frequency Sacrificed for Max")) {
                commentParts.push("Frequency Sacrificed for Max");
              }
            }
          }

          commentsForLog = commentParts.join("; ");

          poRecommendations.push({
            sku: nsidRaw,
            name: product.skuName,
            orderWeek: orderWeekStr,
            qty: qtyToOrder,
            arrivalWeek: arrivalWeekStr,
            comments: commentsForLog,
            supplier: product.supplier,
            truckFillRequired: product.truckFillRequired,
            truckPalletCapacity: product.truckPalletCapacity,
            unitsPerPallet: product.unitsPerPallet,
            orderType: product.orderType,
            caseSize: product.caseSize,
            moq: product.moq
          });

          // Apply order at arrival timing in simulation ledger.
          if (plannedArrivalIndex === -1) {
            commentParts.push("Arrival Beyond Planning Horizon");
          } else if (plannedArrivalIndex <= w) {
            weekClosing += qtyToOrder;
            recAppliedSameWeekForLog = qtyToOrder;
          } else {
            recommendedInbound[plannedArrivalIndex] += qtyToOrder;
          }
          recInboundForLog = parseNum(recommendedInbound[w]) + recAppliedSameWeekForLog;
          closingPostOrderForLog = weekClosing;
        }
      }

      if (isDebugSku) {
        projectionDebugData.push([
          SINGLE_LOCATION_NAME,
          nsidRaw,
          product.skuName,
          thisWeekStr,
          Math.round(startStockForWeek),
          inboundDueForLog,
          inboundRecForLog,
          recInboundForLog,
          recAppliedSameWeekForLog,
          inbound,
          outbound,
          adj,
          Math.round(closingPreOrderForLog),
          Math.round(closingPostOrderForLog),
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

  // --- 3.5) TRUCK-FILL TOP-UP ACROSS SUPPLIER/WEEK ---
  const isWholeNumber = (n) => Math.abs(n - Math.round(n)) < TRUCK_FILL_EPSILON;
  const isFullTruckMultiple = (pallets, cap) => {
    if (cap <= 0) return true;
    const ratio = pallets / cap;
    return Math.abs(ratio - Math.round(ratio)) < TRUCK_FILL_EPSILON;
  };
  const roundDownToOrderIncrement = (rec, rawQty) => {
    let qty = Math.floor(Math.max(0, rawQty));
    if (rec.orderType === "case" && rec.caseSize > 0) {
      qty = Math.floor(qty / rec.caseSize) * rec.caseSize;
    }
    return qty;
  };
  const consolidateRecommendations = (rows) => {
    const merged = new Map();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const q = Math.round(parseNum(r.qty));
      if (q <= 0) continue;
      const key = [
        String(r.sku || ""),
        String(r.name || ""),
        String(r.orderWeek || ""),
        String(r.arrivalWeek || ""),
        String(r.supplier || "")
      ].join("||");
      if (!merged.has(key)) {
        merged.set(key, {
          sku: r.sku,
          name: r.name,
          orderWeek: r.orderWeek,
          qty: q,
          arrivalWeek: r.arrivalWeek,
          comments: r.comments || "",
          supplier: r.supplier || "",
          truckFillRequired: !!r.truckFillRequired,
          truckPalletCapacity: parseNum(r.truckPalletCapacity),
          unitsPerPallet: parseNum(r.unitsPerPallet),
          orderType: r.orderType || "",
          caseSize: parseNum(r.caseSize),
          moq: parseNum(r.moq)
        });
      } else {
        const existing = merged.get(key);
        existing.qty += q;
        if (r.comments) {
          const bits = new Set(
            (existing.comments ? existing.comments.split(";").map((s) => s.trim()) : [])
              .concat(r.comments.split(";").map((s) => s.trim()))
              .filter((s) => s)
          );
          existing.comments = Array.from(bits).join("; ");
        }
      }
    }
    return Array.from(merged.values());
  };

  const truckStats = {
    suppliersConsidered: 0,
    suppliersWithMixedCapacity: 0,
    weeksConsidered: 0,
    weeksAdjusted: 0,
    weeksStillPartial: 0,
    linesAdjusted: 0,
    unitsShifted: 0
  };

  // Work on mutable rows.
  const mutableRecommendations = poRecommendations.map((r) => Object.assign({}, r));
  const supplierKeys = new Set();
  for (let i = 0; i < mutableRecommendations.length; i++) {
    const r = mutableRecommendations[i];
    if (!r.truckFillRequired) continue;
    if (parseNum(r.truckPalletCapacity) <= 0) continue;
    if (parseNum(r.unitsPerPallet) <= 0) continue;
    const key = normalizeCompact(r.supplier);
    if (!key) continue;
    supplierKeys.add(key);
  }

  supplierKeys.forEach((supplierKey) => {
    const supplierRows = mutableRecommendations.filter((r) =>
      r.truckFillRequired &&
      normalizeCompact(r.supplier) === supplierKey &&
      parseNum(r.truckPalletCapacity) > 0 &&
      parseNum(r.unitsPerPallet) > 0 &&
      parseNum(r.qty) > 0
    );
    if (supplierRows.length === 0) return;

    truckStats.suppliersConsidered++;
    const capacityValues = [...new Set(supplierRows.map((r) => parseNum(r.truckPalletCapacity)).filter((v) => v > 0))];
    const truckCapacity = capacityValues.length > 0 ? Math.max.apply(null, capacityValues) : 0;
    if (truckCapacity <= 0) return;
    if (capacityValues.length > 1) truckStats.suppliersWithMixedCapacity++;

    const arrivalWeekIndexes = [...new Set(
      supplierRows
        .map((r) => weekIndexMap.has(r.arrivalWeek) ? weekIndexMap.get(r.arrivalWeek) : -1)
        .filter((idx) => idx >= 0)
    )].sort((a, b) => a - b);

    for (let w = 0; w < arrivalWeekIndexes.length; w++) {
      const weekIdx = arrivalWeekIndexes[w];
      truckStats.weeksConsidered++;

      const weekRows = mutableRecommendations.filter((r) =>
        r.truckFillRequired &&
        normalizeCompact(r.supplier) === supplierKey &&
        weekIndexMap.get(r.arrivalWeek) === weekIdx &&
        parseNum(r.qty) > 0 &&
        parseNum(r.unitsPerPallet) > 0
      );
      if (weekRows.length === 0) continue;

      const palletsNow = weekRows.reduce((sum, r) => sum + (parseNum(r.qty) / parseNum(r.unitsPerPallet)), 0);
      if (palletsNow <= TRUCK_FILL_EPSILON || isFullTruckMultiple(palletsNow, truckCapacity)) continue;

      const targetTruckCount = Math.ceil(palletsNow / truckCapacity);
      const targetPallets = Math.max(truckCapacity, targetTruckCount * truckCapacity);
      let palletsNeeded = targetPallets - palletsNow;
      if (palletsNeeded <= TRUCK_FILL_EPSILON) continue;

      const futureDonors = mutableRecommendations
        .filter((r) =>
          r.truckFillRequired &&
          normalizeCompact(r.supplier) === supplierKey &&
          weekIndexMap.has(r.arrivalWeek) &&
          weekIndexMap.get(r.arrivalWeek) > weekIdx &&
          parseNum(r.qty) > 0 &&
          parseNum(r.unitsPerPallet) > 0
        )
        .sort((a, b) => weekIndexMap.get(a.arrivalWeek) - weekIndexMap.get(b.arrivalWeek));

      let shiftedThisWeek = false;
      for (let d = 0; d < futureDonors.length && palletsNeeded > TRUCK_FILL_EPSILON; d++) {
        const donor = futureDonors[d];
        const donorUpp = parseNum(donor.unitsPerPallet);
        if (donorUpp <= 0) continue;

        const neededQtyRaw = palletsNeeded * donorUpp;
        let moveQty = roundDownToOrderIncrement(donor, Math.min(parseNum(donor.qty), neededQtyRaw));
        if (moveQty <= 0) continue;

        // Avoid leaving a donor line below MOQ unless fully consumed.
        if (parseNum(donor.moq) > 0) {
          const remainingIfMoved = parseNum(donor.qty) - moveQty;
          if (remainingIfMoved > 0 && remainingIfMoved < parseNum(donor.moq)) {
            const fullMoveQty = roundDownToOrderIncrement(donor, parseNum(donor.qty));
            if (fullMoveQty > 0) moveQty = fullMoveQty;
          }
        }
        if (moveQty <= 0 || moveQty > parseNum(donor.qty)) continue;

        let receiver = mutableRecommendations.find((r) =>
          r !== donor &&
          normalizeCompact(r.supplier) === supplierKey &&
          normalizeCompact(r.sku) === normalizeCompact(donor.sku) &&
          weekIndexMap.get(r.arrivalWeek) === weekIdx
        );
        if (!receiver) {
          const donorArrivalIdx = weekIndexMap.has(donor.arrivalWeek) ? weekIndexMap.get(donor.arrivalWeek) : -1;
          const donorOrderIdx = weekIndexMap.has(donor.orderWeek) ? weekIndexMap.get(donor.orderWeek) : donorArrivalIdx;
          const leadGapWeeks = donorArrivalIdx >= 0 && donorOrderIdx >= 0 ? Math.max(0, donorArrivalIdx - donorOrderIdx) : 0;
          const newOrderIdx = Math.max(0, weekIdx - leadGapWeeks);
          receiver = Object.assign({}, donor, {
            qty: 0,
            arrivalWeek: weekHeaders[weekIdx],
            orderWeek: weekHeaders[newOrderIdx] || donor.orderWeek
          });
          mutableRecommendations.push(receiver);
        }

        donor.qty = parseNum(donor.qty) - moveQty;
        receiver.qty = parseNum(receiver.qty) + moveQty;
        const movedPallets = moveQty / donorUpp;
        palletsNeeded = Math.max(0, palletsNeeded - movedPallets);
        shiftedThisWeek = true;
        truckStats.linesAdjusted++;
        truckStats.unitsShifted += moveQty;

        appendComment(receiver, `Truck Fill Pull-Forward: +${Math.round(moveQty)} units from ${donor.arrivalWeek}`);
        appendComment(donor, `Truck Fill Pull-Forward: -${Math.round(moveQty)} units to ${receiver.arrivalWeek}`);
      }

      const weekRowsAfter = mutableRecommendations.filter((r) =>
        r.truckFillRequired &&
        normalizeCompact(r.supplier) === supplierKey &&
        weekIndexMap.get(r.arrivalWeek) === weekIdx &&
        parseNum(r.qty) > 0 &&
        parseNum(r.unitsPerPallet) > 0
      );
      const palletsAfter = weekRowsAfter.reduce((sum, r) => sum + (parseNum(r.qty) / parseNum(r.unitsPerPallet)), 0);
      if (shiftedThisWeek) truckStats.weeksAdjusted++;

      if (!isFullTruckMultiple(palletsAfter, truckCapacity)) {
        truckStats.weeksStillPartial++;
        const shortfall = Math.max(0, targetPallets - palletsAfter);
        for (let j = 0; j < weekRowsAfter.length; j++) {
          appendComment(
            weekRowsAfter[j],
            `Truck Fill: partial (${palletsAfter.toFixed(2)} pallets, short ${shortfall.toFixed(2)})`
          );
        }
      } else if (!isWholeNumber(palletsNow / truckCapacity) || shiftedThisWeek) {
        for (let j = 0; j < weekRowsAfter.length; j++) {
          appendComment(
            weekRowsAfter[j],
            `Truck Fill: full truck multiple reached (${palletsAfter.toFixed(2)} pallets)`
          );
        }
      }
    }
  });

  poRecommendations = consolidateRecommendations(mutableRecommendations);

  // --- 4) WRITE OUTPUTS ---
  const lastRow = outputSheet.getLastRow();
  if (lastRow > 1) outputSheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  if (poRecommendations.length > 0) {
    poRecommendations.sort((a, b) => {
      const aOrder = weekIndexMap.has(a.orderWeek) ? weekIndexMap.get(a.orderWeek) : Number.MAX_SAFE_INTEGER;
      const bOrder = weekIndexMap.has(b.orderWeek) ? weekIndexMap.get(b.orderWeek) : Number.MAX_SAFE_INTEGER;
      if (aOrder !== bOrder) return aOrder - bOrder;
      const aArrival = weekIndexMap.has(a.arrivalWeek) ? weekIndexMap.get(a.arrivalWeek) : Number.MAX_SAFE_INTEGER;
      const bArrival = weekIndexMap.has(b.arrivalWeek) ? weekIndexMap.get(b.arrivalWeek) : Number.MAX_SAFE_INTEGER;
      if (aArrival !== bArrival) return aArrival - bArrival;
      return String(a.sku || "").localeCompare(String(b.sku || ""));
    });
    const outputRows = poRecommendations.map((r) => [
      r.sku,
      r.name,
      r.orderWeek,
      Math.round(parseNum(r.qty)),
      r.arrivalWeek,
      r.comments || ""
    ]);
    outputSheet.getRange(2, 1, outputRows.length, 6).setValues(outputRows);
  }

  if (traceData.length > 0) {
    traceSheet.getRange(2, 1, traceData.length, 11).setValues(traceData);
  }
  if (projectionDebugData.length > 0) {
    projectionDebugSheet.getRange(2, 1, projectionDebugData.length, 31).setValues(projectionDebugData);
  }

  let completionMessage = `Generated ${poRecommendations.length} POs.`;
  if (poRecommendations.length === 0) {
    completionMessage += `\nDebug summary: blocks=${scanStats.blocksScanned}, matched settings=${scanStats.matchedSettings}, missing settings=${scanStats.missingSettings}, missing predicted row=${scanStats.missingPredictedRow}, trigger checks=${scanStats.triggerChecks}, order triggers=${scanStats.orderTriggers}.`;
  }
  if (truckStats.suppliersConsidered > 0) {
    completionMessage += `\nTruck fill summary: suppliers=${truckStats.suppliersConsidered}, mixed capacity suppliers=${truckStats.suppliersWithMixedCapacity}, weeks checked=${truckStats.weeksConsidered}, weeks adjusted=${truckStats.weeksAdjusted}, partial weeks remaining=${truckStats.weeksStillPartial}, lines adjusted=${truckStats.linesAdjusted}, units shifted=${Math.round(truckStats.unitsShifted)}.`;
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
