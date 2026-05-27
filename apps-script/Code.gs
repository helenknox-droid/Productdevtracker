/**
 * Builds the "Own Brand - Upcoming Deadlines" report from "Own-Brand Stage & Gates".
 *
 * Source data starts on row 6, so row 5 is treated as the header row.
 * The report discovers deadline columns dynamically by finding headers that
 * contain the word "deadline".
 */
const REPORT_CONFIG = {
  sourceSheetName: 'Own-Brand Stage & Gates',
  targetSheetName: 'Own Brand - Upcoming Deadlines',
  dataStartRow: 6,
  noLinkedBouquetValue: 'no linked bouquet ids',
  deadlineHeaderContains: 'deadline',
  upcomingWeeks: 4,
  includeCommentsColumn: true,
  columns: {
    referenceNumber: {
      expectedColumn: 'C',
      headers: ['Reference Number', 'Reference No', 'Ref Number', 'Ref No'],
    },
    currentStage: {
      expectedColumn: 'D',
      headers: ['Current Stage'],
    },
    status: {
      expectedColumn: 'F',
      headers: ['Status'],
    },
    launchDatePrimary: {
      expectedColumn: 'I',
      headers: ['Launch Date', 'Launch Week'],
    },
    launchDateFallback: {
      expectedColumn: 'J',
      headers: ['Fallback Launch Date', 'Launch Date Fallback', 'Manual Launch Date'],
    },
    componentName: {
      expectedColumn: 'L',
      headers: ['Brief Name', 'Component Name'],
    },
  },
};

const BASE_REPORT_HEADERS = [
  'Reference Number',
  'Component Name',
  'Launch Date',
  'Status',
  'Current Stage',
  'Stage',
  'Deadline',
];

/**
 * Adds a spreadsheet menu for manually rebuilding the report.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Own Brand Reports')
    .addItem('Build Upcoming Deadlines', 'buildOwnBrandUpcomingDeadlinesReport')
    .addToUi();
}

/**
 * Creates or refreshes the upcoming deadlines report sheet.
 */
function buildOwnBrandUpcomingDeadlinesReport() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(REPORT_CONFIG.sourceSheetName);
  if (!sourceSheet) {
    throw new Error(`Source sheet not found: ${REPORT_CONFIG.sourceSheetName}`);
  }

  const targetSheet = getOrCreateSheet_(spreadsheet, REPORT_CONFIG.targetSheetName);
  const headerValues = getHeaderValues_(sourceSheet, REPORT_CONFIG.dataStartRow - 1);
  const sourceValues = getSourceValues_(sourceSheet, REPORT_CONFIG.dataStartRow);
  const columnMap = buildColumnMap_(headerValues, REPORT_CONFIG);
  const deadlineColumns = findDeadlineColumns_(headerValues, REPORT_CONFIG.deadlineHeaderContains);
  const rows = buildDeadlineRows_(sourceValues, columnMap, deadlineColumns, REPORT_CONFIG);

  writeReport_(targetSheet, rows, REPORT_CONFIG);
}

/**
 * Converts source rows into one output row per qualifying upcoming deadline.
 *
 * @param {Array<Array<*>>} sourceValues
 * @param {Object} columnMap
 * @param {Array<Object>} deadlineColumns
 * @param {Object} config
 * @return {Array<Array<*>>}
 */
function buildDeadlineRows_(sourceValues, columnMap, deadlineColumns, config) {
  const upcomingWindow = getUpcomingWeekWindow_(new Date(), config.upcomingWeeks);
  const reportRows = [];

  sourceValues.forEach((sourceRow) => {
    const referenceNumber = getValueByColumn_(sourceRow, columnMap.referenceNumber);
    if (isBlank_(referenceNumber)) {
      return;
    }

    const launchDate = resolveLaunchDate_(
      getValueByColumn_(sourceRow, columnMap.launchDatePrimary),
      getValueByColumn_(sourceRow, columnMap.launchDateFallback),
      config.noLinkedBouquetValue
    );

    deadlineColumns.forEach((deadlineColumn) => {
      const deadline = getValueByColumn_(sourceRow, deadlineColumn.index);
      const deadlineWeekStart = parseYearWeekStart_(deadline);
      if (!deadlineWeekStart || !isDateInRange_(deadlineWeekStart, upcomingWindow)) {
        return;
      }

      reportRows.push({
        sortDate: deadlineWeekStart,
        values: [
          referenceNumber,
          getValueByColumn_(sourceRow, columnMap.componentName),
          launchDate,
          getValueByColumn_(sourceRow, columnMap.status),
          getValueByColumn_(sourceRow, columnMap.currentStage),
          deadlineColumn.stage,
          deadline,
        ],
      });
    });
  });

  reportRows.sort((a, b) => {
    const dateDifference = a.sortDate.getTime() - b.sortDate.getTime();
    if (dateDifference !== 0) {
      return dateDifference;
    }

    return String(a.values[0]).localeCompare(String(b.values[0]));
  });

  return reportRows.map((row) => row.values);
}

/**
 * Uses the primary launch date unless it contains "no linked bouquet IDs";
 * in that case it uses the fallback launch date.
 */
function resolveLaunchDate_(primaryLaunchDate, fallbackLaunchDate, noLinkedBouquetValue) {
  if (normaliseText_(primaryLaunchDate) === normaliseText_(noLinkedBouquetValue)) {
    return fallbackLaunchDate;
  }

  return primaryLaunchDate;
}

function buildColumnMap_(headerValues, config) {
  return Object.keys(config.columns).reduce((columnMap, columnKey) => {
    columnMap[columnKey] = findColumnIndexByHeader_(headerValues, config.columns[columnKey], columnKey);
    return columnMap;
  }, {});
}

function findColumnIndexByHeader_(headerValues, columnConfig, columnKey) {
  const aliases = columnConfig.headers || [];
  const normalisedAliases = aliases.map(normaliseHeader_);
  const matchingIndexes = [];

  headerValues.forEach((header, index) => {
    if (normalisedAliases.indexOf(normaliseHeader_(header)) !== -1) {
      matchingIndexes.push(index);
    }
  });

  if (matchingIndexes.length === 1) {
    return matchingIndexes[0];
  }

  if (matchingIndexes.length > 1) {
    throw new Error(
      `Multiple columns matched ${columnKey}: ${aliases.join(', ')}. ` +
        `Please make the source headers unique.`
    );
  }

  throw new Error(
    `Could not find a source column for ${columnKey}. Expected one of these headers: ` +
      `${aliases.join(', ')}. The original expected column was ${columnConfig.expectedColumn}. ` +
      `Available headers: ${listAvailableHeaders_(headerValues)}`
  );
}

function findDeadlineColumns_(headerValues, deadlineHeaderContains) {
  const needle = normaliseHeader_(deadlineHeaderContains);
  const deadlineColumns = headerValues
    .map((header, index) => ({
      index,
      header: String(header || '').trim(),
      normalisedHeader: normaliseHeader_(header),
    }))
    .filter((column) => column.normalisedHeader.indexOf(needle) !== -1)
    .map((column) => ({
      index: column.index,
      header: column.header,
      stage: stageFromDeadlineHeader_(column.header),
    }));

  if (deadlineColumns.length === 0) {
    throw new Error(`No deadline columns found. Expected headers containing "${deadlineHeaderContains}".`);
  }

  return deadlineColumns;
}

function stageFromDeadlineHeader_(header) {
  const stage = String(header || '')
    .replace(/deadline/gi, '')
    .replace(/[_\-:]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (/^\d+$/.test(stage)) {
    return `Gate ${stage}`;
  }

  return toTitleCase_(stage || header);
}

function getUpcomingWeekWindow_(today, upcomingWeeks) {
  const start = startOfIsoWeek_(today);
  const end = addDays_(start, upcomingWeeks * 7);
  return { start, end };
}

function parseYearWeekStart_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return startOfIsoWeek_(value);
  }

  const match = String(value || '')
    .trim()
    .match(/^(\d{4})\s*[-/]?\s*W?(\d{1,2})$/i);

  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const week = Number(match[2]);
  if (week < 1 || week > 53) {
    return null;
  }

  return isoWeekStartDate_(year, week);
}

function isoWeekStartDate_(year, week) {
  const fourthOfJanuary = new Date(year, 0, 4);
  const firstIsoWeekStart = startOfIsoWeek_(fourthOfJanuary);
  return addDays_(firstIsoWeekStart, (week - 1) * 7);
}

function startOfIsoWeek_(date) {
  const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = start.getDay() || 7;
  start.setDate(start.getDate() - day + 1);
  start.setHours(0, 0, 0, 0);
  return start;
}

function addDays_(date, days) {
  const result = new Date(date.getTime());
  result.setDate(result.getDate() + days);
  return result;
}

function isDateInRange_(date, range) {
  return date.getTime() >= range.start.getTime() && date.getTime() <= range.end.getTime();
}

function getSourceValues_(sourceSheet, dataStartRow) {
  const lastRow = sourceSheet.getLastRow();
  const lastColumn = sourceSheet.getLastColumn();
  if (lastRow < dataStartRow || lastColumn === 0) {
    return [];
  }

  return sourceSheet
    .getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastColumn)
    .getDisplayValues();
}

function getHeaderValues_(sourceSheet, headerRow) {
  if (headerRow < 1 || sourceSheet.getLastColumn() === 0) {
    return [];
  }

  return sourceSheet.getRange(headerRow, 1, 1, sourceSheet.getLastColumn()).getDisplayValues()[0];
}

function writeReport_(targetSheet, rows, config) {
  const headers = getReportHeaders_(config);
  const output = [headers].concat(rows.map((row) => padRow_(row, headers.length)));

  targetSheet.clearContents();
  targetSheet.getRange(1, 1, output.length, headers.length).setValues(output);
  targetSheet.setFrozenRows(1);
  targetSheet.autoResizeColumns(1, headers.length);
}

function getReportHeaders_(config) {
  return config.includeCommentsColumn ? BASE_REPORT_HEADERS.concat(['Comments']) : BASE_REPORT_HEADERS;
}

function padRow_(row, targetLength) {
  const paddedRow = row.slice();
  while (paddedRow.length < targetLength) {
    paddedRow.push('');
  }

  return paddedRow;
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function getValueByColumn_(row, columnIndex) {
  return row[columnIndex];
}

function isBlank_(value) {
  return value === '' || value === null || typeof value === 'undefined';
}

function normaliseHeader_(value) {
  return normaliseText_(value).replace(/[^a-z0-9]+/g, ' ').trim();
}

function normaliseText_(value) {
  return String(value || '').trim().toLowerCase();
}

function toTitleCase_(value) {
  return String(value || '').replace(/\w\S*/g, (word) => {
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  });
}

function listAvailableHeaders_(headerValues) {
  return headerValues
    .map((header) => String(header || '').trim())
    .filter(Boolean)
    .join(', ');
}
