/**
 * Builds the "Own Brand-Upcoming Deadlines" report from "Own Brand Stage and Gates".
 *
 * Setup:
 * 1. Review REPORT_CONFIG, especially status/current stage columns and deadlineStages.
 * 2. Paste this file into the spreadsheet's Apps Script project.
 * 3. Run buildOwnBrandUpcomingDeadlinesReport(), or use the custom menu after reload.
 */
const REPORT_CONFIG = {
  sourceSheetName: 'Own Brand Stage and Gates',
  targetSheetName: 'Own Brand-Upcoming Deadlines',
  dataStartRow: 6,
  noLinkedBouquetValue: 'no linked bouquet ids',
  columns: {
    referenceNumber: 'C',
    launchDatePrimary: 'I',
    launchDateFallback: 'J',
    componentName: 'L',

    // TODO: Replace these with the source columns once confirmed.
    status: '',
    currentStage: '',
  },

  /**
   * TODO: Replace these examples with every source deadline column.
   *
   * If the stage name is in the header row, use:
   *   { stageFromHeaderColumn: 'M', deadlineColumn: 'M' }
   *
   * If the stage name should be fixed, use:
   *   { stage: 'Design Sign-off', deadlineColumn: 'N' }
   */
  deadlineStages: [
    // { stageFromHeaderColumn: 'M', deadlineColumn: 'M' },
    // { stage: 'Design Sign-off', deadlineColumn: 'N' },
  ],
};

const REPORT_HEADERS = [
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
  validateReportConfig_(REPORT_CONFIG);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(REPORT_CONFIG.sourceSheetName);
  if (!sourceSheet) {
    throw new Error(`Source sheet not found: ${REPORT_CONFIG.sourceSheetName}`);
  }

  const targetSheet = getOrCreateSheet_(spreadsheet, REPORT_CONFIG.targetSheetName);
  const sourceValues = getSourceValues_(sourceSheet, REPORT_CONFIG.dataStartRow);
  const headerValues = getHeaderValues_(sourceSheet, REPORT_CONFIG.dataStartRow - 1);
  const rows = buildDeadlineRows_(sourceValues, headerValues, REPORT_CONFIG);

  writeReport_(targetSheet, rows);
}

/**
 * Converts source rows into one output row per populated deadline.
 *
 * @param {Array<Array<*>>} sourceValues
 * @param {Array<*>} headerValues
 * @param {Object} config
 * @return {Array<Array<*>>}
 */
function buildDeadlineRows_(sourceValues, headerValues, config) {
  const columns = normaliseColumnConfig_(config.columns);
  const deadlineStages = config.deadlineStages.map(normaliseDeadlineStage_);
  const reportRows = [];

  sourceValues.forEach((sourceRow) => {
    const referenceNumber = getValueByColumn_(sourceRow, columns.referenceNumber);
    if (isBlank_(referenceNumber)) {
      return;
    }

    const launchDate = resolveLaunchDate_(
      getValueByColumn_(sourceRow, columns.launchDatePrimary),
      getValueByColumn_(sourceRow, columns.launchDateFallback),
      config.noLinkedBouquetValue
    );

    deadlineStages.forEach((deadlineStage) => {
      const deadline = getValueByColumn_(sourceRow, deadlineStage.deadlineColumn);
      if (isBlank_(deadline)) {
        return;
      }

      reportRows.push([
        referenceNumber,
        getOptionalValueByColumn_(sourceRow, columns.componentName),
        launchDate,
        getOptionalValueByColumn_(sourceRow, columns.status),
        getOptionalValueByColumn_(sourceRow, columns.currentStage),
        resolveStageName_(sourceRow, headerValues, deadlineStage),
        deadline,
      ]);
    });
  });

  return reportRows;
}

/**
 * Uses column I unless it contains "no linked bouquet IDs"; then column J is used.
 */
function resolveLaunchDate_(primaryLaunchDate, fallbackLaunchDate, noLinkedBouquetValue) {
  if (normaliseText_(primaryLaunchDate) === normaliseText_(noLinkedBouquetValue)) {
    return fallbackLaunchDate;
  }

  return primaryLaunchDate;
}

function resolveStageName_(sourceRow, headerValues, deadlineStage) {
  if (deadlineStage.stage) {
    return deadlineStage.stage;
  }

  if (deadlineStage.stageColumn) {
    const stageFromRow = getValueByColumn_(sourceRow, deadlineStage.stageColumn);
    if (!isBlank_(stageFromRow)) {
      return stageFromRow;
    }
  }

  if (deadlineStage.stageFromHeaderColumn) {
    return getValueByColumn_(headerValues, deadlineStage.stageFromHeaderColumn);
  }

  return '';
}

function getSourceValues_(sourceSheet, dataStartRow) {
  const lastRow = sourceSheet.getLastRow();
  const lastColumn = sourceSheet.getLastColumn();
  if (lastRow < dataStartRow || lastColumn === 0) {
    return [];
  }

  return sourceSheet
    .getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastColumn)
    .getValues();
}

function getHeaderValues_(sourceSheet, headerRow) {
  if (headerRow < 1 || sourceSheet.getLastColumn() === 0) {
    return [];
  }

  return sourceSheet.getRange(headerRow, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
}

function writeReport_(targetSheet, rows) {
  targetSheet.clearContents();

  const output = [REPORT_HEADERS].concat(rows);
  targetSheet.getRange(1, 1, output.length, REPORT_HEADERS.length).setValues(output);
  targetSheet.setFrozenRows(1);
  targetSheet.autoResizeColumns(1, REPORT_HEADERS.length);
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function validateReportConfig_(config) {
  const requiredColumns = [
    'referenceNumber',
    'launchDatePrimary',
    'launchDateFallback',
    'componentName',
  ];

  requiredColumns.forEach((columnName) => {
    if (!config.columns[columnName]) {
      throw new Error(`Missing required column config: ${columnName}`);
    }
  });

  if (!Array.isArray(config.deadlineStages) || config.deadlineStages.length === 0) {
    throw new Error(
      'REPORT_CONFIG.deadlineStages is empty. Add every stage/deadline column before running the report.'
    );
  }

  config.deadlineStages.forEach((deadlineStage, index) => {
    if (!deadlineStage.deadlineColumn) {
      throw new Error(`Missing deadlineColumn for deadlineStages[${index}]`);
    }
  });
}

function normaliseColumnConfig_(columns) {
  return Object.keys(columns).reduce((normalised, key) => {
    normalised[key] = columns[key] ? columnLetterToIndex_(columns[key]) : null;
    return normalised;
  }, {});
}

function normaliseDeadlineStage_(deadlineStage) {
  return {
    stage: deadlineStage.stage || '',
    stageColumn: deadlineStage.stageColumn ? columnLetterToIndex_(deadlineStage.stageColumn) : null,
    stageFromHeaderColumn: deadlineStage.stageFromHeaderColumn
      ? columnLetterToIndex_(deadlineStage.stageFromHeaderColumn)
      : null,
    deadlineColumn: columnLetterToIndex_(deadlineStage.deadlineColumn),
  };
}

function columnLetterToIndex_(columnLetter) {
  const letters = String(columnLetter).trim().toUpperCase();
  if (!/^[A-Z]+$/.test(letters)) {
    throw new Error(`Invalid column letter: ${columnLetter}`);
  }

  return letters.split('').reduce((index, letter) => {
    return index * 26 + letter.charCodeAt(0) - 64;
  }, 0) - 1;
}

function getOptionalValueByColumn_(row, columnIndex) {
  return columnIndex === null ? '' : getValueByColumn_(row, columnIndex);
}

function getValueByColumn_(row, columnIndex) {
  return row[columnIndex];
}

function isBlank_(value) {
  return value === '' || value === null || typeof value === 'undefined';
}

function normaliseText_(value) {
  return String(value || '').trim().toLowerCase();
}
