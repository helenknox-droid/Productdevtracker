/**
 * Builds the "Own Brand - Upcoming Deadlines" report from "Own-Brand Stage & Gates".
 *
 * Source data starts on row 6, so row 5 is treated as the header row.
 * The report discovers gate deadline columns dynamically, for example
 * "Gate 0 Deadline" and "Gate 6 Deadline".
 */
const REPORT_CONFIG = {
  sourceSheetName: 'Own-Brand Stage & Gates',
  targetSheetName: 'Own Brand - Upcoming Deadlines',
  diagnosticsSheetName: 'Own Brand - Deadline Diagnostics',
  dataStartRow: 6,
  targetHeaderRow: 2,
  noLinkedBouquetValue: 'no linked bouquet ids',
  upcomingWeeks: 4,
  includeCommentsColumn: true,
  columns: {
    referenceNumber: {
      expectedColumn: 'C',
      headers: ['Reference Number'],
    },
    currentStage: {
      expectedColumn: 'D',
      headers: ['Current Gate'],
    },
    status: {
      expectedColumn: 'F',
      headers: ['Status'],
    },
    launchDatePrimary: {
      expectedColumn: 'I',
      headers: ['Earliest BQID Launch Date'],
    },
    launchDateFallback: {
      expectedColumn: 'J',
      headers: ['Target Launch Date (pre BQID Link)'],
    },
    componentNameCreated: {
      expectedColumn: 'P',
      headers: ['NS NAME (Once Created)'],
      required: false,
    },
    briefName: {
      expectedColumn: 'L',
      headers: ['Brief Name'],
    },
  },
};

const BASE_REPORT_HEADERS = [
  'Reference Number',
  'Component Name',
  'Launch Date',
  'Status',
  'Current Stage',
  'Upcoming Gate',
  'Deadline',
];

/**
 * Adds a spreadsheet menu for manually rebuilding the report.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Own Brand Reports')
    .addItem('Build Upcoming Deadlines', 'buildOwnBrandUpcomingDeadlinesReport')
    .addItem('Build Deadline Diagnostics', 'buildOwnBrandDeadlineDiagnostics')
    .addToUi();
}

/**
 * Creates or refreshes the upcoming deadlines report sheet.
 */
function buildOwnBrandUpcomingDeadlinesReport() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const context = getReportContext_(spreadsheet, REPORT_CONFIG);
  const targetSheet = getOrCreateSheet_(spreadsheet, REPORT_CONFIG.targetSheetName);
  const rows = buildDeadlineRows_(
    context.sourceValues,
    context.columnMap,
    context.deadlineColumns,
    REPORT_CONFIG
  );

  writeReport_(targetSheet, rows, REPORT_CONFIG);
  if (rows.length === 0) {
    const diagnostics = collectDeadlineDiagnostics_(
      context.sourceValues,
      context.columnMap,
      context.deadlineColumns,
      REPORT_CONFIG
    );
    writeEmptyReportNote_(targetSheet, diagnostics, REPORT_CONFIG);
  }
}

/**
 * Writes a diagnostic sheet showing why deadlines did or did not appear.
 */
function buildOwnBrandDeadlineDiagnostics() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const context = getReportContext_(spreadsheet, REPORT_CONFIG);
  const diagnostics = collectDeadlineDiagnostics_(
    context.sourceValues,
    context.columnMap,
    context.deadlineColumns,
    REPORT_CONFIG
  );
  const diagnosticsSheet = getOrCreateSheet_(spreadsheet, REPORT_CONFIG.diagnosticsSheetName);

  writeDiagnostics_(diagnosticsSheet, diagnostics, context, REPORT_CONFIG);
}

function getReportContext_(spreadsheet, config) {
  const sourceSheet = spreadsheet.getSheetByName(config.sourceSheetName);
  if (!sourceSheet) {
    throw new Error(`Source sheet not found: ${config.sourceSheetName}`);
  }

  const headerValues = getHeaderValues_(sourceSheet, config.dataStartRow - 1);
  const sourceValues = getSourceValues_(sourceSheet, config.dataStartRow);
  const columnMap = buildColumnMap_(headerValues, config);
  const deadlineColumns = findDeadlineColumns_(headerValues);

  return {
    sourceSheet,
    headerValues,
    sourceValues,
    columnMap,
    deadlineColumns,
  };
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
          resolveComponentName_(sourceRow, columnMap),
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

function collectDeadlineDiagnostics_(sourceValues, columnMap, deadlineColumns, config) {
  const upcomingWindow = getUpcomingWeekWindow_(new Date(), config.upcomingWeeks);
  const diagnostics = {
    upcomingWindow,
    sourceRows: sourceValues.length,
    deadlineColumns: deadlineColumns.length,
    rowsWithReference: 0,
    blankReferenceRows: 0,
    deadlineCellsWithValues: 0,
    parseableDeadlineCells: 0,
    matchingDeadlineCells: 0,
    unparsableSamples: [],
    pastSamples: [],
    futureSamples: [],
    matchedSamples: [],
  };

  sourceValues.forEach((sourceRow) => {
    const referenceNumber = getValueByColumn_(sourceRow, columnMap.referenceNumber);
    if (isBlank_(referenceNumber)) {
      diagnostics.blankReferenceRows += 1;
      return;
    }

    diagnostics.rowsWithReference += 1;

    deadlineColumns.forEach((deadlineColumn) => {
      const deadline = getValueByColumn_(sourceRow, deadlineColumn.index);
      if (isBlank_(deadline)) {
        return;
      }

      diagnostics.deadlineCellsWithValues += 1;

      const deadlineWeekStart = parseYearWeekStart_(deadline);
      if (!deadlineWeekStart) {
        addSample_(diagnostics.unparsableSamples, referenceNumber, deadlineColumn.header, deadline);
        return;
      }

      diagnostics.parseableDeadlineCells += 1;

      if (isDateInRange_(deadlineWeekStart, upcomingWindow)) {
        diagnostics.matchingDeadlineCells += 1;
        addSample_(diagnostics.matchedSamples, referenceNumber, deadlineColumn.header, deadline);
      } else if (deadlineWeekStart.getTime() < upcomingWindow.start.getTime()) {
        addSample_(diagnostics.pastSamples, referenceNumber, deadlineColumn.header, deadline);
      } else {
        addSample_(diagnostics.futureSamples, referenceNumber, deadlineColumn.header, deadline);
      }
    });
  });

  return diagnostics;
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

function resolveComponentName_(sourceRow, columnMap) {
  const createdComponentName = getValueByColumn_(sourceRow, columnMap.componentNameCreated);
  if (!isBlank_(createdComponentName)) {
    return createdComponentName;
  }

  return getValueByColumn_(sourceRow, columnMap.briefName);
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

  if (columnConfig.required === false) {
    return null;
  }

  throw new Error(
    `Could not find a source column for ${columnKey}. Expected one of these headers: ` +
      `${aliases.join(', ')}. The original expected column was ${columnConfig.expectedColumn}. ` +
      `Available headers: ${listAvailableHeaders_(headerValues)}`
  );
}

function findDeadlineColumns_(headerValues) {
  const deadlineColumns = headerValues
    .map((header, index) => ({
      index,
      header: String(header || '').trim(),
      normalisedHeader: normaliseHeader_(header),
    }))
    .filter((column) => isGateDeadlineHeader_(column.normalisedHeader))
    .map((column) => ({
      index: column.index,
      header: column.header,
      stage: stageFromDeadlineHeader_(column.header),
    }));

  if (deadlineColumns.length === 0) {
    throw new Error('No gate deadline columns found. Expected headers like "Gate 0 Deadline".');
  }

  return deadlineColumns;
}

function isGateDeadlineHeader_(normalisedHeader) {
  return /^gate \d+ deadline\b/.test(normalisedHeader);
}

function stageFromDeadlineHeader_(header) {
  const stage = String(header || '')
    .replace(/deadline/gi, '')
    .replace(/\b(date|week)\b/gi, '')
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

  const text = String(value || '')
    .trim()
    .replace(/^'/, '');
  const match = matchYearWeek_(text);

  if (!match) {
    return null;
  }

  const year = Number(match.year);
  const week = Number(match.week);
  if (week < 1 || week > 53) {
    return null;
  }

  return isoWeekStartDate_(year, week);
}

function matchYearWeek_(text) {
  const yearFirstPatterns = [
    /^(\d{4})\s*[-/]\s*(?:WEEK|WK|W)?\s*(\d{1,2})$/i,
    /^(\d{4})\s+(?:WEEK|WK|W)?\s*(\d{1,2})$/i,
    /^(\d{4})(?:WK|W)(\d{1,2})$/i,
  ];

  for (let index = 0; index < yearFirstPatterns.length; index += 1) {
    const match = text.match(yearFirstPatterns[index]);
    if (match) {
      return { year: match[1], week: match[2] };
    }
  }

  const weekFirstMatch = text.match(/^(?:WEEK|WK|W)\s*(\d{1,2})\s*[-/, ]+\s*(\d{4})$/i);
  if (weekFirstMatch) {
    return { year: weekFirstMatch[2], week: weekFirstMatch[1] };
  }

  return null;
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
  const headerRow = config.targetHeaderRow || 1;

  clearReportOutput_(targetSheet, headerRow);
  targetSheet.getRange(headerRow, 1, output.length, headers.length).setValues(output);
  targetSheet.setFrozenRows(headerRow);
  targetSheet.autoResizeColumns(1, headers.length);
}

function clearReportOutput_(targetSheet, headerRow) {
  const rowsToClear = targetSheet.getMaxRows() - headerRow + 1;
  if (rowsToClear < 1) {
    return;
  }

  targetSheet.getRange(headerRow, 1, rowsToClear, targetSheet.getMaxColumns()).clearContent();
}

function writeEmptyReportNote_(targetSheet, diagnostics, config) {
  const headers = getReportHeaders_(config);
  const message = [
    `No deadlines matched the next ${config.upcomingWeeks} weeks.`,
    `Window checked: ${formatYearWeek_(diagnostics.upcomingWindow.start)} to ` +
      `${formatYearWeek_(diagnostics.upcomingWindow.end)}.`,
    `Found ${diagnostics.deadlineColumns} deadline columns and ` +
      `${diagnostics.deadlineCellsWithValues} populated deadline cells.`,
    `Run "Own Brand Reports > Build Deadline Diagnostics" for details.`,
  ].join(' ');
  const noteRow = padRow_([message], headers.length);
  const noteRowNumber = (config.targetHeaderRow || 1) + 1;

  targetSheet.getRange(noteRowNumber, 1, 1, headers.length).setValues([noteRow]);
}

function writeDiagnostics_(diagnosticsSheet, diagnostics, context, config) {
  const output = [
    ['Metric', 'Value'],
    ['Source sheet', config.sourceSheetName],
    ['Target sheet', config.targetSheetName],
    ['Rows read from source', diagnostics.sourceRows],
    ['Rows with reference number', diagnostics.rowsWithReference],
    ['Rows skipped because reference number is blank', diagnostics.blankReferenceRows],
    ['Deadline columns found', diagnostics.deadlineColumns],
    ['Deadline headers found', context.deadlineColumns.map((column) => column.header).join(', ')],
    ['Upcoming window start', formatYearWeek_(diagnostics.upcomingWindow.start)],
    ['Upcoming window end', formatYearWeek_(diagnostics.upcomingWindow.end)],
    ['Populated deadline cells checked', diagnostics.deadlineCellsWithValues],
    ['Deadline cells parsed as year-week', diagnostics.parseableDeadlineCells],
    ['Deadline cells inside upcoming window', diagnostics.matchingDeadlineCells],
    ['', ''],
    ['Sample matched deadlines', formatSamples_(diagnostics.matchedSamples)],
    ['Sample deadlines before window', formatSamples_(diagnostics.pastSamples)],
    ['Sample deadlines after window', formatSamples_(diagnostics.futureSamples)],
    ['Sample deadline values that could not be parsed', formatSamples_(diagnostics.unparsableSamples)],
  ];

  diagnosticsSheet.clearContents();
  diagnosticsSheet.getRange(1, 1, output.length, 2).setValues(output);
  diagnosticsSheet.setFrozenRows(1);
  diagnosticsSheet.autoResizeColumns(1, 2);
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

function addSample_(samples, referenceNumber, deadlineHeader, deadlineValue) {
  const maxSamples = 10;
  if (samples.length >= maxSamples) {
    return;
  }

  samples.push(`${referenceNumber} | ${deadlineHeader}: ${deadlineValue}`);
}

function formatSamples_(samples) {
  return samples.length === 0 ? 'None' : samples.join('\n');
}

function formatYearWeek_(date) {
  const isoWeek = getIsoWeekInfo_(date);
  return `${isoWeek.year}-W${String(isoWeek.week).padStart(2, '0')}`;
}

function getIsoWeekInfo_(date) {
  const target = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = target.getDay() || 7;
  target.setDate(target.getDate() + 4 - day);

  const isoYear = target.getFullYear();
  const firstDayOfIsoYear = new Date(isoYear, 0, 1);
  const isoWeek = Math.ceil(((target - firstDayOfIsoYear) / 86400000 + 1) / 7);

  return { year: isoYear, week: isoWeek };
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function getValueByColumn_(row, columnIndex) {
  if (columnIndex === null) {
    return '';
  }

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
