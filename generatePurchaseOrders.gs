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
  const ARRIVAL_BUFFER_WEEKS = 2; // Stock must arrive this many weeks BEFORE the shortage

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
  const settingsMap = new Map();
  const generateKey = (loc, sku) => `${String(loc).trim().toLowerCase()}_${String(sku).trim().toLowerCase()}`;
  const parseNum = (val) => {
    if (typeof val === "number") return val;
    if (!val) return 0;
    const n = Number(String(val).replace(/[^0-9.-]/g, ""));
    return isNaN(n) ? 0 : n;
  };

  const settingsData = settingsSheet.getRange(3, 1, settingsSheet.getLastRow() - 2, 13).getValues();

  for (let r = 0; r < settingsData.length; r++) {
    const row = settingsData[r];
    const loc = row[0];
    const nsid = row[1];
    if (loc && nsid) {
      settingsMap.set(generateKey(loc, nsid), {
        skuName: row[2],
        leadTimeDays: Number(row[3]) || 0,
        safetyStock: Number(row[4]) || 0,
        lastWeekPlanned: row[5],
        orderFreqDays: Number(row[6]) || 0,
        moq: Number(row[7]) || 0,
        maxStock: parseNum(row[8]),
        orderType: String(row[9]).toLowerCase(),
        caseSize: Number(row[10]) || 0,
        unitsPerPallet: row[11]
      });
    }
  }

  // --- 2. MAP HEADERS ---
  const invLastRow = invSheet.getLastRow();
  const invLastCol = invSheet.getLastColumn();
  const headerRowVals = invSheet.getRange(6, 1, 1, invLastCol).getValues()[0];
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

  let currentWeekStr = invSheet.getRange("B2").getValue();
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
  const invData = invSheet.getRange(7, 1, invLastRow - 6, invLastCol).getValues();

  // --- 3. DYNAMIC SCANNING ---
  const poRecommendations = [];
  const traceData = [];
  let currentRowIndex = 0;

  while (currentRowIndex < invData.length) {
    const firstRow = invData[currentRowIndex];
    const locRaw = firstRow[0];
    const nsidRaw = firstRow[1];

    if (!locRaw || !nsidRaw) {
      currentRowIndex++;
      continue;
    }

    const loc = String(locRaw).trim();
    const nsid = String(nsidRaw).trim();
    const key = generateKey(loc, nsid);

    // Collect block rows
    const blockRows = [];
    while (currentRowIndex < invData.length) {
      const nextRow = invData[currentRowIndex];
      const nextLoc = String(nextRow[0]).trim();
      const nextNsid = String(nextRow[1]).trim();

      if (blockRows.length > 0 && nextLoc !== "" && (nextLoc !== loc || nextNsid !== nsid)) {
        break;
      }
      blockRows.push(nextRow);
      currentRowIndex++;
    }

    const product = settingsMap.get(key);
    if (!product) continue;

    // Find Logic Rows
    let rowPredicted, rowInboundDue, rowInboundRec, rowForecast, rowTransfers, rowAdj;
    const clean = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, "");

    blockRows.forEach((row) => {
      const label = clean(row[3]);
      if (label.includes("predicted")) rowPredicted = row;
      else if (label.includes("inbounddue")) rowInboundDue = row;
      else if (label.includes("inboundreceived")) rowInboundRec = row;
      else if (label.includes("forecasted")) rowForecast = row;
      else if (label.includes("transfers")) rowTransfers = row;
      else if (label.includes("adjustment")) rowAdj = row;
    });

    const zeroRow = new Array(invLastCol).fill(0);
    if (!rowInboundDue) rowInboundDue = zeroRow;
    if (!rowInboundRec) rowInboundRec = zeroRow;
    if (!rowForecast) rowForecast = zeroRow;
    if (!rowTransfers) rowTransfers = zeroRow;
    if (!rowAdj) rowAdj = zeroRow;

    // Initialize Ledger
    const startWeekHeader = weekHeaders[simulationStartIndex];
    const startCol = weekColMap.get(startWeekHeader);
    const rawPred = rowPredicted ? rowPredicted[startCol] : "";

    if (String(rawPred).trim() === "") continue;

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
        // Trigger condition now uses Safety + following week's demand.
        const requiredFloor = getRequiredFloorForWeek(w);

        if (weekClosing < requiredFloor) {
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
          const bufferedArrivalDate = new Date(shortageDate);
          bufferedArrivalDate.setDate(bufferedArrivalDate.getDate() - ARRIVAL_BUFFER_WEEKS * 7);

          if (bufferedArrivalDate < currentDateObj) {
            bufferedArrivalDate.setTime(currentDateObj.getTime());
          }

          const orderDate = new Date(bufferedArrivalDate);
          orderDate.setDate(orderDate.getDate() - product.leadTimeDays);

          let orderWeekStr;
          if (orderDate < currentDateObj) orderWeekStr = currentWeekStr;
          else orderWeekStr = getIsoWeekString(orderDate);

          const arrivalWeekStr = getIsoWeekString(bufferedArrivalDate);

          const potentialPeakStock = startStockForWeek + inbound + qtyToOrder;
          if (
            product.maxStock > 0 &&
            potentialPeakStock > product.maxStock &&
            !commentParts.some((c) => c.startsWith("Max Capacity Breached"))
          ) {
            commentParts.push("Max Capacity Breached");
          }

          poRecommendations.push([
            locRaw,
            nsidRaw,
            product.skuName,
            orderWeekStr,
            qtyToOrder,
            arrivalWeekStr,
            commentParts.join("; "),
            product.moq
          ]);

          weekClosing += qtyToOrder;
        }
      }

      // Trace logging for specific SKU to debug
      if (nsidRaw.includes("PK-11-00002")) {
        traceData.push([
          locRaw,
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
  if (lastRow > 1) outputSheet.getRange(2, 1, lastRow - 1, 8).clearContent();

  if (poRecommendations.length > 0) {
    poRecommendations.sort((a, b) => (a[3] < b[3] ? -1 : 1));
    outputSheet.getRange(2, 1, poRecommendations.length, 8).setValues(poRecommendations);
  }

  if (traceData.length > 0) {
    traceSheet.getRange(2, 1, traceData.length, 11).setValues(traceData);
  }

  SpreadsheetApp.getUi().alert(`Generated ${poRecommendations.length} POs.`);
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
