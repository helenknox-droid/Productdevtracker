function generateOwnBrandComponentRiskReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sourceSheet = ss.getSheetByName('Own-Brand Stage & Gates');
  const reportSheet = ss.getSheetByName('Own Brand - Component - Risk Report');

  if (!sourceSheet) {
    throw new Error('Source sheet not found: Own-Brand Stage & Gates');
  }

  if (!reportSheet) {
    throw new Error('Report sheet not found: Own Brand - Component - Risk Report');
  }

  const sourceHeaderRow = 5;
  const sourceDataStartRow = 6;
  const reportDataStartRow = 3;

  const sourceValues = sourceSheet.getDataRange().getValues();
  const headers = sourceValues[sourceHeaderRow - 1];
  const headerMap = buildHeaderMap(headers);

  const requiredHeaders = [
    'Abandon SKU?',
    'Reference Number',
    'Status',
    'Current Gate',
    'Detailed Status',
    'Brief Name',
    'NSID (Once Created)',
    'Target Launch Date (pre BQID Link)',
    'Earliest BQID Launch Date',
    'Gate 0 Status',
    'Gate 0 Deadline',
    'Gate 1 Status',
    'Gate 1 Deadline',
    'Gate 2 Status',
    'Gate 2 Deadline',
    'Gate 3 Status',
    'Gate 3 Deadline',
    'Gate 4 Status',
    'Gate 4 Deadline',
    'Gate 5 Status',
    'Gate 5 Deadline',
    'Gate 6 Status',
    'Gate 6 Deadline'
  ];

  validateHeaders(headerMap, requiredHeaders);

  const statusesToInclude = [
    'Launch At Risk 🔥',
    'Status Unknown - Due Date Missing ⚠️'
  ];

  const reportRows = [];

  for (let i = sourceDataStartRow - 1; i < sourceValues.length; i++) {
    const row = sourceValues[i];

    const abandonSku = row[headerMap[normaliseHeader('Abandon SKU?')]];

    if (String(abandonSku).trim().toLowerCase() === 'yes') {
      continue;
    }

    const status = row[headerMap[normaliseHeader('Status')]];

    if (!statusesToInclude.includes(String(status).trim())) {
      continue;
    }

    const referenceNumber = row[headerMap[normaliseHeader('Reference Number')]];

    const nsid = row[headerMap[normaliseHeader('NSID (Once Created)')]];
    const briefName = row[headerMap[normaliseHeader('Brief Name')]];
    const componentName = nsid || briefName;

    if (!String(componentName || '').trim()) {
      continue;
    }

    const launchDate = getLaunchDate(row, headerMap);

    const currentGate = row[headerMap[normaliseHeader('Current Gate')]];

    const statusDescription = row[headerMap[normaliseHeader('Detailed Status')]];

    const expectedWeeksLate = calculateExpectedWeeksLate(row, headerMap);

    reportRows.push([
      referenceNumber,
      componentName,
      launchDate,
      status,
      currentGate,
      statusDescription,
      expectedWeeksLate,
      ''
    ]);
  }

  clearExistingReport(reportSheet, reportDataStartRow);

  if (reportRows.length > 0) {
    reportSheet
      .getRange(reportDataStartRow, 1, reportRows.length, 8)
      .setValues(reportRows);
  }
}

function buildHeaderMap(headers) {
  const map = {};

  headers.forEach((header, index) => {
    const normalised = normaliseHeader(header);

    if (normalised) {
      map[normalised] = index;
    }
  });

  return map;
}

function normaliseHeader(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function validateHeaders(headerMap, requiredHeaders) {
  const missingHeaders = requiredHeaders.filter(header => {
    return headerMap[normaliseHeader(header)] === undefined;
  });

  if (missingHeaders.length > 0) {
    throw new Error(
      'The following required headers are missing from row 5: ' +
      missingHeaders.join(', ')
    );
  }
}

function getLaunchDate(row, headerMap) {
  const earliestBqidLaunchDate = row[
    headerMap[normaliseHeader('Earliest BQID Launch Date')]
  ];

  const targetLaunchDate = row[
    headerMap[normaliseHeader('Target Launch Date (pre BQID Link)')]
  ];

  if (isYearWeek(earliestBqidLaunchDate)) {
    return earliestBqidLaunchDate;
  }

  return targetLaunchDate;
}

function isYearWeek(value) {
  const match = String(value || '').trim().match(/^(\d{4})-W(\d{2})$/);

  if (!match) {
    return false;
  }

  const week = Number(match[2]);
  return week >= 1 && week <= 53;
}

function calculateExpectedWeeksLate(row, headerMap) {
  const completeStatus = 'Complete ✅';

  for (let gateNumber = 0; gateNumber <= 6; gateNumber++) {
    const statusHeader = `Gate ${gateNumber} Status`;
    const deadlineHeader = `Gate ${gateNumber} Deadline`;

    const gateStatus = row[headerMap[normaliseHeader(statusHeader)]];

    if (String(gateStatus).trim() === completeStatus) {
      continue;
    }

    const deadlineValue = row[headerMap[normaliseHeader(deadlineHeader)]];
    return calculateWeeksLateFromYearWeek(deadlineValue);
  }

  return 0;
}

function calculateWeeksLateFromYearWeek(yearWeekValue) {
  const value = String(yearWeekValue || '').trim();

  if (!value) {
    return 0;
  }

  const match = value.match(/^(\d{4})-W(\d{1,2})$/);

  if (!match) {
    return 0;
  }

  const year = Number(match[1]);
  const week = Number(match[2]);

  const deadlineWeekStart = getIsoWeekStartDate(year, week);
  const currentWeekStart = getCurrentIsoWeekStartDate();

  const millisecondsPerWeek = 7 * 24 * 60 * 60 * 1000;
  const weeksLate = Math.floor(
    (currentWeekStart.getTime() - deadlineWeekStart.getTime()) / millisecondsPerWeek
  );

  return Math.max(0, weeksLate);
}

function getIsoWeekStartDate(year, week) {
  const simpleDate = new Date(year, 0, 1 + (week - 1) * 7);
  const dayOfWeek = simpleDate.getDay();
  const isoDayOfWeek = dayOfWeek === 0 ? 7 : dayOfWeek;

  simpleDate.setDate(simpleDate.getDate() + 1 - isoDayOfWeek);
  simpleDate.setHours(0, 0, 0, 0);

  return simpleDate;
}

function getCurrentIsoWeekStartDate() {
  const today = new Date();
  const dayOfWeek = today.getDay();
  const isoDayOfWeek = dayOfWeek === 0 ? 7 : dayOfWeek;

  today.setDate(today.getDate() + 1 - isoDayOfWeek);
  today.setHours(0, 0, 0, 0);

  return today;
}

function clearExistingReport(reportSheet, reportDataStartRow) {
  const lastRow = reportSheet.getLastRow();

  if (lastRow >= reportDataStartRow) {
    reportSheet
      .getRange(reportDataStartRow, 1, lastRow - reportDataStartRow + 1, 8)
      .clearContent();
  }
}
