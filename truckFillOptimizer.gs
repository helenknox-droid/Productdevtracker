/**
 * TRUCK FILL OPTIMIZER (SECONDARY REPORT)
 *
 * Purpose:
 * - Read SKU-level PO recommendations and build a truck-fill adjusted plan.
 * - Keep PO recommendation output unchanged; this writes a separate plan.
 *
 * Safety defaults:
 * - Decreases are OFF by default per SKU (to avoid reducing service coverage).
 * - Increases are ON by default per SKU.
 *
 * Source sheet:
 * - 🛍️ Purchase Order Recommendations (A:F expected):
 *   A SKU, B Name, C Order Week, D Qty, E Arrival Week, F Comments
 *
 * Product settings source:
 * - 📋 Product Settings (starting row 3)
 *   L Units per pallet
 *   M Truck Fill Required (checkbox)
 *   N Supplier
 *   O Pallets per truck
 *
 * Settings sheet (optional):
 * - ⚙️ Truck Fill Settings
 *   A Key, B Value
 *   LOOKBACK_WEEKS (default 0)
 *   LOOKAHEAD_WEEKS (default 2)
 *   MIN_TRUCK_FILL_PCT (default 90)
 *   TARGET_TRUCK_FILL_PCT (default 100)
 *   ALLOW_MIXED_ORDER_WEEKS (default TRUE)
 *   MAX_UPLIFT_PCT (default 0.10)
 */

function setupTruckFillSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const settingsSheet = getOrCreateSheet_(ss, "⚙️ Truck Fill Settings");
  const planSheet = getOrCreateSheet_(ss, "🚚 Truck Fill Plan");
  const auditSheet = getOrCreateSheet_(ss, "🧾 Truck Fill Audit");

  ensureHeader_(settingsSheet, ["Key", "Value"]);
  if (settingsSheet.getLastRow() < 2 || String(settingsSheet.getRange(2, 1).getValue()).trim() === "") {
    settingsSheet.getRange(2, 1, 6, 2).setValues([
      ["LOOKBACK_WEEKS", 0],
      ["LOOKAHEAD_WEEKS", 2],
      ["MIN_TRUCK_FILL_PCT", 90],
      ["TARGET_TRUCK_FILL_PCT", 100],
      ["ALLOW_MIXED_ORDER_WEEKS", true],
      ["MAX_UPLIFT_PCT", 0.1]
    ]);
  }

  ensureHeader_(planSheet, [
    "Supplier",
    "Arrival Week",
    "Group Key",
    "SKU",
    "Name",
    "Order Week",
    "Original Qty",
    "Adjusted Qty",
    "Delta Qty",
    "Units Per Pallet",
    "Original Pallets",
    "Adjusted Pallets",
    "Pallets Per Truck",
    "Group Original Pallets",
    "Group Adjusted Pallets",
    "Target Pallets",
    "Fill %",
    "Adjustment Reason",
    "Original Comments"
  ]);

  ensureHeader_(auditSheet, [
    "Supplier",
    "Arrival Week",
    "Group Key",
    "Lines",
    "Pallets Per Truck",
    "Original Pallets",
    "Adjusted Pallets",
    "Target Pallets",
    "Fill %",
    "Status",
    "Warnings"
  ]);
}

/**
 * Convenience entrypoint so users can run `truckFillOptimizer`.
 */
function truckFillOptimizer() {
  buildTruckFillPlan();
}

function buildTruckFillPlan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recSheet = ss.getSheetByName("🛍️ Purchase Order Recommendations");
  const productSettingsSheet = ss.getSheetByName("📋 Product Settings");
  const settingsSheet = ss.getSheetByName("⚙️ Truck Fill Settings");
  const planSheet = ss.getSheetByName("🚚 Truck Fill Plan");
  const auditSheet = ss.getSheetByName("🧾 Truck Fill Audit");

  if (!recSheet || !productSettingsSheet || !planSheet || !auditSheet) {
    SpreadsheetApp.getUi().alert("Missing required sheets. Need 🛍️ Purchase Order Recommendations, 📋 Product Settings, 🚚 Truck Fill Plan, 🧾 Truck Fill Audit.");
    return;
  }

  const settings = readTruckFillSettings_(settingsSheet);
  const truckMap = readTruckFillDataFromProductSettings_(productSettingsSheet, settings);
  const recLines = readRecommendationLines_(recSheet);

  const groups = new Map();
  const unmappedSkus = [];

  for (let i = 0; i < recLines.length; i++) {
    const line = recLines[i];
    const cfg = truckMap.get(line.sku);
    if (!cfg) {
      unmappedSkus.push(line.sku);
      continue;
    }

    const step = cfg.unitsPerPallet > 0 ? cfg.unitsPerPallet : 1;
    const qtyOriginal = Math.max(0, line.qty);
    const minQtyRaw = cfg.allowDecrease ? qtyOriginal * (1 - cfg.maxDecreasePct) : qtyOriginal;
    const maxQtyRaw = cfg.allowIncrease ? qtyOriginal * (1 + cfg.maxIncreasePct) : qtyOriginal;
    const minQty = roundDownToStep_(minQtyRaw, step);
    const maxQty = roundUpToStep_(maxQtyRaw, step);

    const lineObj = {
      ...line,
      supplier: cfg.supplier,
      unitsPerPallet: cfg.unitsPerPallet,
      palletsPerTruck: cfg.palletsPerTruck,
      priority: cfg.priority,
      qtyOriginal,
      qtyAdjusted: qtyOriginal,
      qtyMin: minQty,
      qtyMax: Math.max(maxQty, minQty),
      reason: "No Change"
    };

    const groupKey = `${cfg.supplier}__${line.arrivalWeek}`;
    if (!groups.has(groupKey)) {
      groups.set(groupKey, {
        key: groupKey,
        supplier: cfg.supplier,
        arrivalWeek: line.arrivalWeek,
        palletsPerTruck: cfg.palletsPerTruck,
        lines: [],
        warnings: []
      });
    }
    const g = groups.get(groupKey);
    if (cfg.palletsPerTruck !== g.palletsPerTruck) {
      g.warnings.push(`Mixed palletsPerTruck in group (${g.palletsPerTruck} vs ${cfg.palletsPerTruck})`);
      g.palletsPerTruck = Math.max(g.palletsPerTruck, cfg.palletsPerTruck);
    }
    g.lines.push(lineObj);
  }

  const planRows = [];
  const auditRows = [];

  groups.forEach((group) => {
    const palletsPerTruck = Math.max(1, group.palletsPerTruck || 1);
    const originalPallets = sumPallets_(group.lines, "qtyOriginal");
    let adjustedPallets = sumPallets_(group.lines, "qtyAdjusted");

    // Target nearest full-truck multiple while staying within fill band.
    const minTarget = palletsPerTruck * settings.minFillRatio;
    const baseTrucks = Math.max(1, Math.round(adjustedPallets / palletsPerTruck));
    let targetPallets = baseTrucks * palletsPerTruck * settings.targetFillRatio;
    if (targetPallets < minTarget) targetPallets = minTarget;

    let deltaPallets = targetPallets - adjustedPallets;
    if (Math.abs(deltaPallets) > 0.0001) {
      if (deltaPallets > 0) {
        // Increase first from most flexible lines (higher priority number).
        const incCandidates = group.lines
          .map((l) => ({
            line: l,
            capPallets: qtyToPallets_(Math.max(0, l.qtyMax - l.qtyAdjusted), l.unitsPerPallet)
          }))
          .filter((x) => x.capPallets > 0)
          .sort((a, b) => b.line.priority - a.line.priority);

        for (let i = 0; i < incCandidates.length && deltaPallets > 0.0001; i++) {
          const c = incCandidates[i];
          const addPallets = Math.min(deltaPallets, c.capPallets);
          const addQty = palletsToQty_(addPallets, c.line.unitsPerPallet);
          c.line.qtyAdjusted += addQty;
          c.line.reason = "Increased for Truck Fill";
          deltaPallets -= qtyToPallets_(addQty, c.line.unitsPerPallet);
        }
      } else {
        // Decrease from most flexible lines (higher priority number).
        const decCandidates = group.lines
          .map((l) => ({
            line: l,
            capPallets: qtyToPallets_(Math.max(0, l.qtyAdjusted - l.qtyMin), l.unitsPerPallet)
          }))
          .filter((x) => x.capPallets > 0)
          .sort((a, b) => b.line.priority - a.line.priority);

        let needReduce = Math.abs(deltaPallets);
        for (let i = 0; i < decCandidates.length && needReduce > 0.0001; i++) {
          const c = decCandidates[i];
          const cutPallets = Math.min(needReduce, c.capPallets);
          const cutQty = palletsToQty_(cutPallets, c.line.unitsPerPallet);
          c.line.qtyAdjusted -= cutQty;
          c.line.reason = "Reduced for Truck Fill";
          needReduce -= qtyToPallets_(cutQty, c.line.unitsPerPallet);
        }
      }
    }

    adjustedPallets = sumPallets_(group.lines, "qtyAdjusted");
    const fillPct = adjustedPallets / palletsPerTruck;
    const maxAllowedPallets = palletsPerTruck * settings.maxFillRatio;
    let status = "OK";
    if (adjustedPallets < minTarget - 0.0001) status = "UNDERFILLED";
    if (adjustedPallets > maxAllowedPallets + 0.0001) status = "OVERFILLED";

    auditRows.push([
      group.supplier,
      group.arrivalWeek,
      group.key,
      group.lines.length,
      palletsPerTruck,
      round3_(originalPallets),
      round3_(adjustedPallets),
      round3_(targetPallets),
      round3_(fillPct),
      status,
      group.warnings.join("; ")
    ]);

    for (let i = 0; i < group.lines.length; i++) {
      const l = group.lines[i];
      const origPallets = qtyToPallets_(l.qtyOriginal, l.unitsPerPallet);
      const adjPallets = qtyToPallets_(l.qtyAdjusted, l.unitsPerPallet);
      const deltaQty = l.qtyAdjusted - l.qtyOriginal;
      planRows.push([
        group.supplier,
        group.arrivalWeek,
        group.key,
        l.sku,
        l.name,
        l.orderWeek,
        l.qtyOriginal,
        l.qtyAdjusted,
        deltaQty,
        l.unitsPerPallet,
        round3_(origPallets),
        round3_(adjPallets),
        palletsPerTruck,
        round3_(originalPallets),
        round3_(adjustedPallets),
        round3_(targetPallets),
        round3_(fillPct),
        l.reason,
        l.comments
      ]);
    }
  });

  // Write outputs
  clearDataRows_(planSheet);
  clearDataRows_(auditSheet);

  if (planRows.length > 0) {
    planRows.sort((a, b) => {
      if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1;
      if (a[1] !== b[1]) return a[1] < b[1] ? -1 : 1;
      return a[3] < b[3] ? -1 : 1;
    });
    planSheet.getRange(2, 1, planRows.length, 19).setValues(planRows);
  }

  if (auditRows.length > 0) {
    auditRows.sort((a, b) => {
      if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1;
      return a[1] < b[1] ? -1 : 1;
    });
    auditSheet.getRange(2, 1, auditRows.length, 11).setValues(auditRows);
  }

  const warningMsg = unmappedSkus.length > 0
    ? `\nUnmapped SKUs in 📋 Product Settings (M/N/O/L): ${[...new Set(unmappedSkus)].join(", ")}`
    : "";
  SpreadsheetApp.getUi().alert(
    `Truck fill plan created.\nLines: ${planRows.length}\nGroups: ${auditRows.length}${warningMsg}`
  );
}

// ---------- Internal helpers ----------

function readRecommendationLines_(recSheet) {
  const lastRow = recSheet.getLastRow();
  if (lastRow < 2) return [];
  const values = recSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const sku = String(r[0] || "").trim();
    const name = String(r[1] || "").trim();
    const orderWeek = String(r[2] || "").trim();
    const qty = Number(r[3]) || 0;
    const arrivalWeek = String(r[4] || "").trim();
    const comments = String(r[5] || "").trim();
    if (!sku || qty <= 0 || !arrivalWeek) continue;
    rows.push({ sku, name, orderWeek, qty, arrivalWeek, comments });
  }
  return rows;
}

function readTruckFillDataFromProductSettings_(productSettingsSheet, settings) {
  const map = new Map();
  const lastRow = productSettingsSheet.getLastRow();
  if (lastRow < 2) return map;
  const values = productSettingsSheet.getRange(3, 1, Math.max(0, lastRow - 2), 15).getValues();

  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const sku = String(r[0] || r[1] || "").trim(); // prefer col A, fallback col B
    if (!sku) continue;
    const truckFillRequired = parseBool_(r[12], false); // col M
    if (!truckFillRequired) continue;

    const supplier = String(r[13] || "").trim(); // col N
    const unitsPerPallet = Number(r[11]) || 0; // col L
    const palletsPerTruck = Number(r[14]) || 0; // col O
    if (!supplier || unitsPerPallet <= 0 || palletsPerTruck <= 0) continue;

    map.set(sku, {
      supplier,
      unitsPerPallet,
      palletsPerTruck,
      allowIncrease: true,
      maxIncreasePct: settings.maxUpliftPct,
      allowDecrease: false,
      maxDecreasePct: 0,
      priority: 100
    });
  }
  return map;
}

function readTruckFillSettings_(settingsSheet) {
  const defaults = {
    targetFillRatio: 1.0,
    minFillRatio: 0.9,
    maxFillRatio: 1.05,
    maxUpliftPct: 0.10,
    lookbackWeeks: 0,
    lookaheadWeeks: 2,
    allowMixedOrderWeeks: true
  };
  if (!settingsSheet || settingsSheet.getLastRow() < 2) return defaults;

  const values = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 2).getValues();
  const kv = {};
  for (let i = 0; i < values.length; i++) {
    const key = String(values[i][0] || "").trim().toUpperCase();
    if (!key) continue;
    kv[key] = Number(values[i][1]);
  }

  return {
    // Preferred keys (percent-based)
    targetFillRatio: pctToRatio_(kv.TARGET_TRUCK_FILL_PCT, safeNumber_(kv.TARGET_TRUCK_FILL_RATIO, defaults.targetFillRatio)),
    minFillRatio: pctToRatio_(kv.MIN_TRUCK_FILL_PCT, safeNumber_(kv.MIN_TRUCK_FILL_RATIO, defaults.minFillRatio)),
    // Keep max fill ratio backward-compatible only
    maxFillRatio: safeNumber_(kv.MAX_TRUCK_FILL_RATIO, defaults.maxFillRatio),
    // Accept MAX_UPLIFT_PCT as either percent (10) or ratio (0.10)
    maxUpliftPct: normalizePctOrRatio_(kv.MAX_UPLIFT_PCT, defaults.maxUpliftPct),
    lookbackWeeks: safeInt_(kv.LOOKBACK_WEEKS, defaults.lookbackWeeks),
    lookaheadWeeks: safeInt_(kv.LOOKAHEAD_WEEKS, defaults.lookaheadWeeks),
    allowMixedOrderWeeks: parseBool_(kv.ALLOW_MIXED_ORDER_WEEKS, defaults.allowMixedOrderWeeks)
  };
}

function getOrCreateSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeader_(sheet, header) {
  const row1 = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  const empty = row1.every((v) => String(v || "").trim() === "");
  if (empty) sheet.getRange(1, 1, 1, header.length).setValues([header]);
}

function clearDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1 && lastCol > 0) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
}

function parseBool_(v, fallback) {
  if (typeof v === "boolean") return v;
  const s = String(v || "").trim().toLowerCase();
  if (s === "true" || s === "yes" || s === "1") return true;
  if (s === "false" || s === "no" || s === "0") return false;
  return fallback;
}

function clampPct_(v, fallback) {
  const n = Number(v);
  if (isNaN(n)) return fallback;
  if (n < 0) return 0;
  if (n > 1) return 1;
  return n;
}

function safeNumber_(v, fallback) {
  return typeof v === "number" && !isNaN(v) ? v : fallback;
}

function safeInt_(v, fallback) {
  const n = Number(v);
  return isNaN(n) ? fallback : Math.max(0, Math.floor(n));
}

function pctToRatio_(pctOrUndefined, fallbackRatio) {
  const n = Number(pctOrUndefined);
  if (isNaN(n)) return fallbackRatio;
  return n > 1 ? n / 100 : n;
}

function normalizePctOrRatio_(val, fallbackRatio) {
  const n = Number(val);
  if (isNaN(n)) return fallbackRatio;
  const ratio = n > 1 ? n / 100 : n;
  return clampPct_(ratio, fallbackRatio);
}

function qtyToPallets_(qty, unitsPerPallet) {
  if (!unitsPerPallet || unitsPerPallet <= 0) return 0;
  return qty / unitsPerPallet;
}

function palletsToQty_(pallets, unitsPerPallet) {
  return pallets * unitsPerPallet;
}

function roundDownToStep_(n, step) {
  if (step <= 0) return Math.floor(n);
  return Math.floor(n / step) * step;
}

function roundUpToStep_(n, step) {
  if (step <= 0) return Math.ceil(n);
  return Math.ceil(n / step) * step;
}

function sumPallets_(lines, qtyField) {
  let total = 0;
  for (let i = 0; i < lines.length; i++) {
    total += qtyToPallets_(lines[i][qtyField], lines[i].unitsPerPallet);
  }
  return total;
}

function round3_(n) {
  return Math.round((n + Number.EPSILON) * 1000) / 1000;
}
