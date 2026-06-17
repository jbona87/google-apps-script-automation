```javascript
/**
 * SYSTEM: Google Sheets Automation Command Centre
 * DESCRIPTION:
 *   Advanced Google Apps Script pipeline for extracting, validating,
 *   processing and visualising structured API data.
 *
 * AUTHOR: jbona87
 * VERSION: 2.0.0
 */


/* ==========================================================
   01 // ENGINE CONFIGURATION
   ========================================================== */

const ENGINE_CONFIG = Object.freeze({
  SYSTEM_NAME: "GOOGLE SHEETS AUTOMATION ENGINE",
  VERSION: "2.0.0",

  DASHBOARD_SHEET: "Automation_Command_Centre",
  LOG_SHEET: "System_Log",

  TABLE_HEADER_ROW: 9,
  DATA_START_ROW: 10,

  DATE_FORMAT: "yyyy-MM-dd HH:mm:ss",
  MAX_DASHBOARD_ROWS: 100
});


const ENGINE_PALETTE = Object.freeze({
  BACKGROUND: "#070B14",
  PANEL: "#101827",
  PANEL_ALT: "#131D2E",
  BORDER: "#263449",

  CYAN: "#00F5D4",
  PURPLE: "#8B5CF6",
  GREEN: "#22C55E",
  YELLOW: "#FBBF24",
  RED: "#EF4444",

  TEXT: "#E6EDF3",
  MUTED: "#94A3B8",
  DARK_TEXT: "#06101A"
});


const STATUS_CONFIG = Object.freeze({
  Active: {
    colour: "#00D9FF",
    label: "ACTIVE"
  },

  Optimized: {
    colour: "#8B5CF6",
    label: "OPTIMIZED"
  },

  Completed: {
    colour: "#22C55E",
    label: "COMPLETED"
  },

  Pending: {
    colour: "#FBBF24",
    label: "PENDING"
  },

  Error: {
    colour: "#EF4444",
    label: "ERROR"
  }
});


/* ==========================================================
   02 // CUSTOM SHEET MENU
   ========================================================== */

/**
 * Adds the automation controls whenever the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚡ Automation Engine")
    .addItem(
      "🚀 Run Automation Pipeline",
      "runAutomationPipeline"
    )
    .addSeparator()
    .addItem(
      "📊 Open Command Centre",
      "openAutomationDashboard"
    )
    .addItem(
      "📋 Open System Log",
      "openSystemLog"
    )
    .addSeparator()
    .addItem(
      "🧹 Reset Command Centre",
      "resetAutomationDashboard"
    )
    .addToUi();
}


/* ==========================================================
   03 // MAIN PIPELINE CONTROLLER
   ========================================================== */

/**
 * Main automation entry point.
 */
function runAutomationPipeline() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const documentLock = LockService.getDocumentLock();

  const startedAt = new Date();
  const executionStart = Date.now();

  try {
    if (!documentLock.tryLock(30000)) {
      throw new Error(
        "Another pipeline process is currently running."
      );
    }

    spreadsheet.toast(
      "Initialising data extraction and processing...",
      "⚡ Automation Engine",
      5
    );

    console.log(
      "SYSTEM BOOT // Initialising automation pipeline"
    );

    /*
     * Replace extractSimulatedPayload_() with:
     *
     * const rawPayload = fetchJsonFromApi_(
     *   "https://api.example.com/data",
     *   "YOUR_API_TOKEN"
     * );
     */

    const rawPayload = extractSimulatedPayload_();

    const processedRecords = processDataPayload_(
      rawPayload,
      startedAt
    );

    const dashboard = getOrCreateSheet_(
      spreadsheet,
      ENGINE_CONFIG.DASHBOARD_SHEET
    );

    renderCommandCentre_(
      dashboard,
      processedRecords,
      startedAt
    );

    const executionTime = Date.now() - executionStart;

    appendSystemLog_(
      spreadsheet,
      "SUCCESS",
      "Pipeline completed successfully",
      processedRecords.length,
      executionTime
    );

    spreadsheet.setActiveSheet(dashboard);

    spreadsheet.toast(
      processedRecords.length +
        " data streams processed successfully.",
      "✅ Pipeline Complete",
      6
    );

    console.log(
      "PIPELINE SUCCESS // " +
        processedRecords.length +
        " records processed in " +
        executionTime +
        "ms"
    );

  } catch (error) {
    const executionTime = Date.now() - executionStart;

    appendSystemLog_(
      spreadsheet,
      "ERROR",
      error.message,
      0,
      executionTime
    );

    spreadsheet.toast(
      error.message,
      "❌ Pipeline Error",
      8
    );

    console.error(
      "PIPELINE FAILURE // " + error.stack
    );

    SpreadsheetApp.getUi().alert(
      "Automation Pipeline Error",
      error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } finally {
    documentLock.releaseLock();
  }
}


/* ==========================================================
   04 // DATA EXTRACTION LAYER
   ========================================================== */

/**
 * Demonstration payload.
 *
 * Replace this function with fetchJsonFromApi_()
 * when connecting to a real API.
 */
function extractSimulatedPayload_() {
  return [
    {
      id: "TR-101",
      target: "Advanced AI Overview Audit",
      status: "Active"
    },
    {
      id: "TR-102",
      target: "Link-Building Supply Stream",
      status: "Optimized"
    },
    {
      id: "TR-103",
      target: "Programmatic Keyword Classifier",
      status: "Completed"
    },
    {
      id: "TR-104",
      target: "Technical SEO Monitoring Core",
      status: "Active"
    },
    {
      id: "TR-105",
      target: "Search Visibility Intelligence Report",
      status: "Pending"
    },
    {
      id: "TR-106",
      target: "Automated Backlink Risk Scanner",
      status: "Optimized"
    }
  ];
}


/**
 * Generic JSON API connector.
 *
 * Expected API response:
 *
 * [
 *   {
 *     "id": "TR-101",
 *     "target": "Example process",
 *     "status": "Active"
 *   }
 * ]
 */
function fetchJsonFromApi_(url, apiToken) {
  if (!url) {
    throw new Error(
      "An API endpoint URL is required."
    );
  }

  const requestOptions = {
    method: "get",
    muteHttpExceptions: true,
    contentType: "application/json",
    headers: {}
  };

  if (apiToken) {
    requestOptions.headers.Authorization =
      "Bearer " + apiToken;
  }

  const response = UrlFetchApp.fetch(
    url,
    requestOptions
  );

  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (
    responseCode < 200 ||
    responseCode >= 300
  ) {
    throw new Error(
      "API request failed with HTTP status " +
        responseCode +
        ". Response: " +
        responseBody.substring(0, 300)
    );
  }

  let parsedResponse;

  try {
    parsedResponse = JSON.parse(responseBody);
  } catch (error) {
    throw new Error(
      "The API returned invalid JSON data."
    );
  }

  /*
   * Supports either:
   * [
   *   {...}
   * ]
   *
   * or:
   *
   * {
   *   "data": [
   *     {...}
   *   ]
   * }
   */

  const payload = Array.isArray(parsedResponse)
    ? parsedResponse
    : parsedResponse.data;

  if (!Array.isArray(payload)) {
    throw new Error(
      "The API payload must contain an array of records."
    );
  }

  return payload;
}


/* ==========================================================
   05 // DATA PROCESSING LAYER
   ========================================================== */

/**
 * Validates, cleans and deduplicates incoming data.
 */
function processDataPayload_(
  rawPayload,
  processedAt
) {
  if (!Array.isArray(rawPayload)) {
    throw new Error(
      "The extracted payload is not a valid array."
    );
  }

  const processedRecords = [];
  const processedIds = new Set();

  rawPayload.forEach(function(item, index) {
    if (!item || typeof item !== "object") {
      console.warn(
        "Invalid item ignored at position " + index
      );
      return;
    }

    const transactionId = sanitiseValue_(
      item.id || "TR-" + String(index + 1).padStart(3, "0")
    );

    const targetPayload = sanitiseValue_(
      item.target || "Unnamed Automation Process"
    );

    const systemStatus = normaliseStatus_(
      item.status
    );

    if (processedIds.has(transactionId)) {
      console.warn(
        "Duplicate transaction ignored: " +
          transactionId
      );
      return;
    }

    processedIds.add(transactionId);

    processedRecords.push({
      id: transactionId,
      target: targetPayload,
      status: systemStatus,
      timestamp: new Date(processedAt)
    });
  });

  if (processedRecords.length === 0) {
    throw new Error(
      "No valid records remained after payload validation."
    );
  }

  return processedRecords;
}


/**
 * Cleans a value before writing it to Sheets.
 */
function sanitiseValue_(value) {
  return String(value)
    .replace(/\s+/g, " ")
    .trim();
}


/**
 * Ensures statuses follow the supported system taxonomy.
 */
function normaliseStatus_(status) {
  const incomingStatus = sanitiseValue_(
    status || "Pending"
  ).toLowerCase();

  const supportedStatuses =
    Object.keys(STATUS_CONFIG);

  for (
    let i = 0;
    i < supportedStatuses.length;
    i++
  ) {
    if (
      supportedStatuses[i].toLowerCase() ===
      incomingStatus
    ) {
      return supportedStatuses[i];
    }
  }

  return "Pending";
}


/* ==========================================================
   06 // DASHBOARD RENDERING ENGINE
   ========================================================== */

/**
 * Builds the complete futuristic dashboard.
 */
function renderCommandCentre_(
  sheet,
  records,
  lastUpdated
) {
  prepareDashboardCanvas_(sheet);

  renderSystemHeader_(
    sheet,
    lastUpdated
  );

  renderMetricCards_(
    sheet,
    records
  );

  renderDataStream_(
    sheet,
    records
  );

  const summaryRange = renderStatusSummary_(
    sheet,
    records
  );

  renderStatusChart_(
    sheet,
    summaryRange
  );

  applyDashboardFinishing_(
    sheet,
    records.length
  );
}


/**
 * Clears previous output and prepares the visual canvas.
 */
function prepareDashboardCanvas_(sheet) {
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });

  sheet
    .getRange(
      1,
      1,
      sheet.getMaxRows(),
      Math.min(sheet.getMaxColumns(), 8)
    )
    .breakApart();

  sheet.clear();
  sheet.setConditionalFormatRules([]);

  sheet.setHiddenGridlines(true);
  sheet.setTabColor(
    ENGINE_PALETTE.CYAN
  );

  sheet
    .getRange(
      1,
      1,
      ENGINE_CONFIG.MAX_DASHBOARD_ROWS,
      8
    )
    .setBackground(
      ENGINE_PALETTE.BACKGROUND
    )
    .setFontFamily("Roboto Mono")
    .setFontColor(
      ENGINE_PALETTE.TEXT
    );

  sheet.setFrozenRows(
    ENGINE_CONFIG.TABLE_HEADER_ROW
  );
}


/**
 * Creates the primary system title and status bar.
 */
function renderSystemHeader_(
  sheet,
  lastUpdated
) {
  const titleRange = sheet.getRange(
    "A1:H2"
  );

  titleRange.merge();

  titleRange
    .setValue(
      "⚡ " +
        ENGINE_CONFIG.SYSTEM_NAME
    )
    .setBackground(
      ENGINE_PALETTE.BACKGROUND
    )
    .setFontColor(
      ENGINE_PALETTE.CYAN
    )
    .setFontSize(22)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      ENGINE_PALETTE.CYAN,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );

  const formattedDate =
    Utilities.formatDate(
      lastUpdated,
      Session.getScriptTimeZone(),
      ENGINE_CONFIG.DATE_FORMAT
    );

  const subtitleRange = sheet.getRange(
    "A3:H3"
  );

  subtitleRange.merge();

  subtitleRange
    .setValue(
      "VERSION " +
        ENGINE_CONFIG.VERSION +
        "  //  API DATA EXTRACTION  //  STRUCTURED REPORTING  //  LAST SYNC: " +
        formattedDate
    )
    .setBackground(
      ENGINE_PALETTE.PANEL
    )
    .setFontColor(
      ENGINE_PALETTE.MUTED
    )
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.setRowHeight(1, 35);
  sheet.setRowHeight(2, 35);
  sheet.setRowHeight(3, 28);
}


/**
 * Generates the top-level KPI cards.
 */
function renderMetricCards_(
  sheet,
  records
) {
  const statusCounts =
    calculateStatusCounts_(records);

  const metricCards = [
    {
      columns: "A:B",
      label: "TOTAL DATA STREAMS",
      value: records.length,
      colour: ENGINE_PALETTE.CYAN
    },
    {
      columns: "C:D",
      label: "ACTIVE SYSTEMS",
      value: statusCounts.Active || 0,
      colour: STATUS_CONFIG.Active.colour
    },
    {
      columns: "E:F",
      label: "OPTIMIZED STREAMS",
      value: statusCounts.Optimized || 0,
      colour: STATUS_CONFIG.Optimized.colour
    },
    {
      columns: "G:H",
      label: "COMPLETED TASKS",
      value: statusCounts.Completed || 0,
      colour: STATUS_CONFIG.Completed.colour
    }
  ];

  metricCards.forEach(function(card) {
    const columns = card.columns.split(":");

    const labelRange =
      sheet.getRange(
        columns[0] + "5:" +
        columns[1] + "5"
      );

    const valueRange =
      sheet.getRange(
        columns[0] + "6:" +
        columns[1] + "6"
      );

    labelRange.merge();
    valueRange.merge();

    labelRange
      .setValue(card.label)
      .setBackground(
        ENGINE_PALETTE.PANEL
      )
      .setFontColor(
        ENGINE_PALETTE.MUTED
      )
      .setFontSize(9)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    valueRange
      .setValue(card.value)
      .setBackground(
        ENGINE_PALETTE.PANEL
      )
      .setFontColor(card.colour)
      .setFontSize(20)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    sheet
      .getRange(
        labelRange.getRow(),
        labelRange.getColumn(),
        2,
        2
      )
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        card.colour,
        SpreadsheetApp.BorderStyle.SOLID
      );
  });

  sheet.setRowHeight(5, 25);
  sheet.setRowHeight(6, 38);
}


/**
 * Builds the live data table.
 */
function renderDataStream_(
  sheet,
  records
) {
  const sectionTitle =
    sheet.getRange("A8:D8");

  sectionTitle.merge();

  sectionTitle
    .setValue(
      "LIVE AUTOMATION DATA STREAM"
    )
    .setBackground(
      ENGINE_PALETTE.PANEL
    )
    .setFontColor(
      ENGINE_PALETTE.CYAN
    )
    .setFontSize(11)
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  const headers = [[
    "TRANSACTION ID",
    "TARGET PAYLOAD",
    "SYSTEM STATUS",
    "PROCESSED AT"
  ]];

  const headerRange = sheet.getRange(
    ENGINE_CONFIG.TABLE_HEADER_ROW,
    1,
    1,
    headers[0].length
  );

  headerRange
    .setValues(headers)
    .setBackground(
      ENGINE_PALETTE.PANEL_ALT
    )
    .setFontColor(
      ENGINE_PALETTE.CYAN
    )
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      ENGINE_PALETTE.BORDER,
      SpreadsheetApp.BorderStyle.SOLID
    );

  const rowValues = records.map(
    function(record) {
      return [
        record.id,
        record.target,
        record.status,
        record.timestamp
      ];
    }
  );

  const dataRange = sheet.getRange(
    ENGINE_CONFIG.DATA_START_ROW,
    1,
    rowValues.length,
    4
  );

  dataRange
    .setValues(rowValues)
    .setFontColor(
      ENGINE_PALETTE.TEXT
    )
    .setVerticalAlignment("middle")
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      ENGINE_PALETTE.BORDER,
      SpreadsheetApp.BorderStyle.SOLID
    );

  const rowBackgrounds =
    rowValues.map(function(row, index) {
      const background =
        index % 2 === 0
          ? ENGINE_PALETTE.PANEL
          : ENGINE_PALETTE.PANEL_ALT;

      return [
        background,
        background,
        background,
        background
      ];
    });

  dataRange.setBackgrounds(
    rowBackgrounds
  );

  sheet
    .getRange(
      ENGINE_CONFIG.DATA_START_ROW,
      1,
      rowValues.length,
      1
    )
    .setFontColor(
      ENGINE_PALETTE.CYAN
    )
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  const statusRange = sheet.getRange(
    ENGINE_CONFIG.DATA_START_ROW,
    3,
    rowValues.length,
    1
  );

  statusRange
    .setHorizontalAlignment("center")
    .setFontWeight("bold");

  sheet
    .getRange(
      ENGINE_CONFIG.DATA_START_ROW,
      4,
      rowValues.length,
      1
    )
    .setNumberFormat(
      ENGINE_CONFIG.DATE_FORMAT
    )
    .setHorizontalAlignment("center")
    .setFontColor(
      ENGINE_PALETTE.MUTED
    );

  applyStatusConditionalFormatting_(
    sheet,
    statusRange
  );

  applyStatusValidation_(
    statusRange
  );

  sheet
    .getRange(
      ENGINE_CONFIG.TABLE_HEADER_ROW,
      1,
      rowValues.length + 1,
      4
    )
    .createFilter();
}


/**
 * Creates the status distribution table.
 */
function renderStatusSummary_(
  sheet,
  records
) {
  const sectionTitle =
    sheet.getRange("F8:H8");

  sectionTitle.merge();

  sectionTitle
    .setValue(
      "SYSTEM STATUS DISTRIBUTION"
    )
    .setBackground(
      ENGINE_PALETTE.PANEL
    )
    .setFontColor(
      ENGINE_PALETTE.PURPLE
    )
    .setFontSize(11)
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  const statusCounts =
    calculateStatusCounts_(records);

  const summaryRows = [
    ["STATUS", "COUNT", "SHARE"]
  ];

  Object.keys(STATUS_CONFIG)
    .forEach(function(status) {
      const count =
        statusCounts[status] || 0;

      if (count > 0) {
        summaryRows.push([
          status,
          count,
          count / records.length
        ]);
      }
    });

  const summaryRange = sheet.getRange(
    9,
    6,
    summaryRows.length,
    3
  );

  summaryRange
    .setValues(summaryRows)
    .setBackground(
      ENGINE_PALETTE.PANEL
    )
    .setFontColor(
      ENGINE_PALETTE.TEXT
    )
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      ENGINE_PALETTE.BORDER,
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet
    .getRange(9, 6, 1, 3)
    .setBackground(
      ENGINE_PALETTE.PANEL_ALT
    )
    .setFontColor(
      ENGINE_PALETTE.PURPLE
    )
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  if (summaryRows.length > 1) {
    sheet
      .getRange(
        10,
        7,
        summaryRows.length - 1,
        1
      )
      .setHorizontalAlignment("center")
      .setFontWeight("bold");

    sheet
      .getRange(
        10,
        8,
        summaryRows.length - 1,
        1
      )
      .setNumberFormat("0.0%")
      .setHorizontalAlignment("center")
      .setFontColor(
        ENGINE_PALETTE.MUTED
      );
  }

  return sheet.getRange(
    9,
    6,
    summaryRows.length,
    2
  );
}


/**
 * Adds a futuristic status distribution chart.
 */
function renderStatusChart_(
  sheet,
  summaryRange
) {
  if (summaryRange.getNumRows() <= 1) {
    return;
  }

  const chart = sheet
    .newChart()
    .setChartType(
      Charts.ChartType.PIE
    )
    .addRange(summaryRange)
    .setPosition(
      15,
      6,
      0,
      0
    )
    .setOption(
      "title",
      "AUTOMATION STATUS MATRIX"
    )
    .setOption(
      "pieHole",
      0.58
    )
    .setOption(
      "backgroundColor",
      ENGINE_PALETTE.PANEL
    )
    .setOption(
      "titleTextStyle",
      {
        color: ENGINE_PALETTE.CYAN,
        fontName: "Roboto Mono",
        fontSize: 12,
        bold: true
      }
    )
    .setOption(
      "legend",
      {
        position: "bottom",
        textStyle: {
          color: ENGINE_PALETTE.TEXT,
          fontName: "Roboto Mono",
          fontSize: 10
        }
      }
    )
    .setOption(
      "pieSliceTextStyle",
      {
        color: ENGINE_PALETTE.TEXT
      }
    )
    .setOption(
      "colors",
      [
        STATUS_CONFIG.Active.colour,
        STATUS_CONFIG.Optimized.colour,
        STATUS_CONFIG.Completed.colour,
        STATUS_CONFIG.Pending.colour,
        STATUS_CONFIG.Error.colour
      ]
    )
    .setOption(
      "width",
      480
    )
    .setOption(
      "height",
      280
    )
    .build();

  sheet.insertChart(chart);
}


/* ==========================================================
   07 // DASHBOARD FORMATTING
   ========================================================== */

/**
 * Applies colour rules to each status.
 */
function applyStatusConditionalFormatting_(
  sheet,
  statusRange
) {
  const rules = [];

  Object.keys(STATUS_CONFIG)
    .forEach(function(status) {
      const style =
        STATUS_CONFIG[status];

      const rule =
        SpreadsheetApp
          .newConditionalFormatRule()
          .whenTextEqualTo(status)
          .setBackground(style.colour)
          .setFontColor(
            ENGINE_PALETTE.DARK_TEXT
          )
          .setBold(true)
          .setRanges([statusRange])
          .build();

      rules.push(rule);
    });

  sheet.setConditionalFormatRules(
    rules
  );
}


/**
 * Adds an editable status dropdown.
 */
function applyStatusValidation_(
  statusRange
) {
  const validation =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        Object.keys(STATUS_CONFIG),
        true
      )
      .setAllowInvalid(false)
      .setHelpText(
        "Select a valid automation system status."
      )
      .build();

  statusRange.setDataValidation(
    validation
  );
}


/**
 * Applies final dimensions and display settings.
 */
function applyDashboardFinishing_(
  sheet,
  recordCount
) {
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 330);
  sheet.setColumnWidth(3, 155);
  sheet.setColumnWidth(4, 190);

  sheet.setColumnWidth(5, 35);
  sheet.setColumnWidth(6, 170);
  sheet.setColumnWidth(7, 95);
  sheet.setColumnWidth(8, 100);

  sheet.setRowHeight(
    ENGINE_CONFIG.TABLE_HEADER_ROW,
    34
  );

  if (recordCount > 0) {
    sheet.setRowHeights(
      ENGINE_CONFIG.DATA_START_ROW,
      recordCount,
      32
    );
  }

  sheet
    .getRange(
      ENGINE_CONFIG.DATA_START_ROW,
      2,
      recordCount,
      1
    )
    .setWrap(true);
}


/* ==========================================================
   08 // SYSTEM UTILITIES
   ========================================================== */

/**
 * Calculates the number of records per status.
 */
function calculateStatusCounts_(
  records
) {
  return records.reduce(
    function(counts, record) {
      counts[record.status] =
        (counts[record.status] || 0) + 1;

      return counts;
    },
    {}
  );
}


/**
 * Returns an existing sheet or creates it.
 */
function getOrCreateSheet_(
  spreadsheet,
  sheetName
) {
  return (
    spreadsheet.getSheetByName(sheetName) ||
    spreadsheet.insertSheet(sheetName)
  );
}


/* ==========================================================
   09 // SYSTEM LOGGING
   ========================================================== */

/**
 * Stores pipeline executions in a dedicated log sheet.
 */
function appendSystemLog_(
  spreadsheet,
  level,
  message,
  recordCount,
  executionTime
) {
  const logSheet = getOrCreateSheet_(
    spreadsheet,
    ENGINE_CONFIG.LOG_SHEET
  );

  if (logSheet.getLastRow() === 0) {
    logSheet
      .getRange(1, 1, 1, 5)
      .setValues([[
        "TIMESTAMP",
        "LEVEL",
        "SYSTEM EVENT",
        "RECORDS",
        "EXECUTION TIME"
      ]])
      .setBackground(
        ENGINE_PALETTE.PANEL
      )
      .setFontColor(
        ENGINE_PALETTE.CYAN
      )
      .setFontWeight("bold");

    logSheet.setFrozenRows(1);
    logSheet.setHiddenGridlines(true);

    logSheet.setColumnWidth(1, 190);
    logSheet.setColumnWidth(2, 100);
    logSheet.setColumnWidth(3, 420);
    logSheet.setColumnWidth(4, 100);
    logSheet.setColumnWidth(5, 140);
  }

  const nextRow =
    logSheet.getLastRow() + 1;

  logSheet
    .getRange(nextRow, 1, 1, 5)
    .setValues([[
      new Date(),
      level,
      message,
      recordCount,
      executionTime + " ms"
    ]])
    .setBackground(
      ENGINE_PALETTE.BACKGROUND
    )
    .setFontColor(
      ENGINE_PALETTE.TEXT
    );

  logSheet
    .getRange(nextRow, 1)
    .setNumberFormat(
      ENGINE_CONFIG.DATE_FORMAT
    );

  const levelColour =
    level === "SUCCESS"
      ? ENGINE_PALETTE.GREEN
      : ENGINE_PALETTE.RED;

  logSheet
    .getRange(nextRow, 2)
    .setBackground(levelColour)
    .setFontColor(
      ENGINE_PALETTE.DARK_TEXT
    )
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}


/* ==========================================================
   10 // MENU ACTIONS
   ========================================================== */

function openAutomationDashboard() {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const dashboard =
    spreadsheet.getSheetByName(
      ENGINE_CONFIG.DASHBOARD_SHEET
    );

  if (!dashboard) {
    SpreadsheetApp.getUi().alert(
      "The command centre has not been generated yet. Run the automation pipeline first."
    );

    return;
  }

  dashboard.showSheet();
  spreadsheet.setActiveSheet(dashboard);
}


function openSystemLog() {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const logSheet =
    spreadsheet.getSheetByName(
      ENGINE_CONFIG.LOG_SHEET
    );

  if (!logSheet) {
    SpreadsheetApp.getUi().alert(
      "No system logs are currently available."
    );

    return;
  }

  logSheet.showSheet();
  spreadsheet.setActiveSheet(logSheet);
}


function resetAutomationDashboard() {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const dashboard =
    spreadsheet.getSheetByName(
      ENGINE_CONFIG.DASHBOARD_SHEET
    );

  if (!dashboard) {
    spreadsheet.toast(
      "No command centre exists to reset.",
      "Automation Engine",
      4
    );

    return;
  }

  if (dashboard.getFilter()) {
    dashboard.getFilter().remove();
  }

  dashboard.getCharts().forEach(
    function(chart) {
      dashboard.removeChart(chart);
    }
  );

  dashboard.clear();
  dashboard.setConditionalFormatRules([]);

  spreadsheet.toast(
    "Command centre reset successfully.",
    "🧹 System Reset",
    4
  );
}
```
