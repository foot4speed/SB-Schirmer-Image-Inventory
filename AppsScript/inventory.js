/**
 * SB Schirmer Image Inventory Engine
 * Phase 1: Inventory Writer
 *
 * Scans the configured Google Drive folder and subfolders.
 * Writes one spreadsheet row per image.
 *
 * This script does not move, rename, delete, edit, duplicate,
 * or change permissions on any original Drive files.
 */

function getPhase1InventoryHeaders_() {
  return [
    "Filename",
    "Recipe",
    "Meal",
    "Ingredients",
    "Uses Nicer Slicer",
    "Quality",
    "Orientation",
    "Drive File ID",
    "Drive Link",
    "Source Folder",
    "Analysis Status",
    "Confidence",
    "Notes",
    "Date Analyzed"
  ];
}

/**
 * Creates or validates the approved Phase 1 inventory sheet.
 * Existing nonblank headers are never silently rearranged or deleted.
 */
function setupInventorySheet() {
  var spreadsheet = SpreadsheetApp.openById(
    SB_CONFIG.SPREADSHEET_ID
  );

  var sheet = spreadsheet.getSheetByName(
    SB_CONFIG.INVENTORY_SHEET_NAME
  );

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      SB_CONFIG.INVENTORY_SHEET_NAME
    );
  }

  var requiredHeaders = getPhase1InventoryHeaders_();

  if (
    sheet.getLastRow() === 0 ||
    sheet.getLastColumn() === 0
  ) {
    sheet
      .getRange(1, 1, 1, requiredHeaders.length)
      .setValues([requiredHeaders]);
  } else {
    var existingHeaders = sheet
      .getRange(
        1,
        1,
        1,
        Math.max(
          sheet.getLastColumn(),
          requiredHeaders.length
        )
      )
      .getValues()[0]
      .slice(0, requiredHeaders.length)
      .map(function(value) {
        return String(value || "").trim();
      });

    var headersMatch =
      existingHeaders.length === requiredHeaders.length;

    if (headersMatch) {
      for (
        var headerIndex = 0;
        headerIndex < requiredHeaders.length;
        headerIndex++
      ) {
        if (
          existingHeaders[headerIndex] !==
          requiredHeaders[headerIndex]
        ) {
          headersMatch = false;
          break;
        }
      }
    }

    if (!headersMatch) {
      throw new Error(
        "The Image Inventory header row does not match the approved Phase 1 schema. " +
        "No existing columns were changed. Review the sheet before continuing."
      );
    }
  }

  sheet.setFrozenRows(1);

  sheet
    .getRange(1, 1, 1, requiredHeaders.length)
    .setFontWeight("bold");

  if (!sheet.getFilter()) {
    sheet
      .getRange(
        1,
        1,
        Math.max(sheet.getLastRow(), 1),
        requiredHeaders.length
      )
      .createFilter();
  }

  var widths = [
    190, 170, 120, 220, 130, 110, 110,
    190, 220, 220, 140, 110, 240, 150
  ];

  for (
    var column = 1;
    column <= widths.length;
    column++
  ) {
    sheet.setColumnWidth(
      column,
      widths[column - 1]
    );
  }

  applyPhase1Dropdowns_(sheet);

  SpreadsheetApp.flush();

  console.log(
    "Image Inventory sheet setup complete."
  );

  return sheet;
}

function applyPhase1Dropdowns_(sheet) {
  var firstDataRow = 2;
  var validationRows = Math.max(
    sheet.getMaxRows() - 1,
    1
  );

  setDropdownValidation_(
    sheet,
    firstDataRow,
    3,
    validationRows,
    [
      "Breakfast",
      "Lunch",
      "Dinner",
      "Snacks",
      "Hors d'oeuvres",
      "Multiple",
      "Unknown"
    ]
  );

  setDropdownValidation_(
    sheet,
    firstDataRow,
    5,
    validationRows,
    ["Yes", "Likely", "No", "Unknown"]
  );

  setDropdownValidation_(
    sheet,
    firstDataRow,
    6,
    validationRows,
    ["Excellent", "Good", "Usable", "Poor"]
  );

  setDropdownValidation_(
    sheet,
    firstDataRow,
    11,
    validationRows,
    [
      "Complete",
      "Review Required",
      "Error",
      "Skipped",
      "Pending Analysis",
      "Analyzing"
    ]
  );

  setDropdownValidation_(
    sheet,
    firstDataRow,
    12,
    validationRows,
    ["High", "Medium", "Low"]
  );
}

function setDropdownValidation_(
  sheet,
  startRow,
  column,
  numberOfRows,
  values
) {
  var rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();

  sheet
    .getRange(
      startRow,
      column,
      numberOfRows,
      1
    )
    .setDataValidation(rule);
}

/**
 * Production inventory entry point.
 * Scans for all new images.
 */
function writeSupportedAssetInventory() {
  return writeSupportedAssetInventoryWithLimit_(
    null
  );
}

/**
 * Shared inventory writer used by production and tests.
 * maxImages:
 *   null = no test limit
 *   positive integer = stop after that many new images
 */
function writeSupportedAssetInventoryWithLimit_(
  maxImages
) {
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    console.log(
      "Another inventory scan is already running."
    );

    return {
      addedCount: 0,
      firstRow: null
    };
  }

  try {
    var sheet = setupInventorySheet();

    var headerMap =
      getInventoryWriterHeaderMap_(sheet);

    validateInventoryWriterHeaders_(
      headerMap
    );

    var existingFileIds =
      getExistingInventoryFileIds_(
        sheet,
        headerMap
      );

    var rootFolder =
      DriveApp.getFolderById(
        SB_CONFIG.ROOT_FOLDER_ID
      );

    var scanState = {
      newRows: [],
      maxImages: normalizeImageLimit_(
        maxImages
      ),
      foundCount: 0
    };

    scanInventoryFolder_(
      rootFolder,
      rootFolder.getName(),
      headerMap,
      existingFileIds,
      scanState
    );

    if (scanState.newRows.length === 0) {
      console.log(
        "Inventory scan complete. No new images were found."
      );

      return {
        addedCount: 0,
        firstRow: null
      };
    }

    var firstNewRow =
      sheet.getLastRow() + 1;

    sheet
      .getRange(
        firstNewRow,
        1,
        scanState.newRows.length,
        getPhase1InventoryHeaders_().length
      )
      .setValues(scanState.newRows);

    SpreadsheetApp.flush();

    console.log(
      "Inventory scan complete. New images added: " +
      scanState.newRows.length
    );

    return {
      addedCount: scanState.newRows.length,
      firstRow: firstNewRow
    };
  } finally {
    lock.releaseLock();
  }
}

function normalizeImageLimit_(maxImages) {
  if (
    maxImages === null ||
    maxImages === undefined ||
    maxImages === ""
  ) {
    return null;
  }

  var numericLimit = Number(maxImages);

  if (
    !isFinite(numericLimit) ||
    numericLimit <= 0
  ) {
    throw new Error(
      "Image limit must be a positive number."
    );
  }

  return Math.floor(numericLimit);
}

/**
 * Recursively scans folders until the optional test limit is reached.
 */
function scanInventoryFolder_(
  folder,
  folderPath,
  headerMap,
  existingFileIds,
  scanState
) {
  if (inventoryLimitReached_(scanState)) {
    return;
  }

  var files = folder.getFiles();

  while (files.hasNext()) {
    if (inventoryLimitReached_(scanState)) {
      return;
    }

    var file = files.next();

    var mimeType = String(
      file.getMimeType() || ""
    );

    if (mimeType.indexOf("image/") !== 0) {
      continue;
    }

    var fileId = file.getId();

    if (existingFileIds[fileId]) {
      continue;
    }

    var row = createBlankInventoryRow_();

    setInventoryRowValue_(
      row,
      headerMap,
      "Filename",
      file.getName()
    );

    setInventoryRowValue_(
      row,
      headerMap,
      "Drive File ID",
      fileId
    );

    setInventoryRowValue_(
      row,
      headerMap,
      "Drive Link",
      file.getUrl()
    );

    setInventoryRowValue_(
      row,
      headerMap,
      "Source Folder",
      folderPath
    );

    setInventoryRowValue_(
      row,
      headerMap,
      "Analysis Status",
      "Pending Analysis"
    );

    scanState.newRows.push(row);
    scanState.foundCount++;
    existingFileIds[fileId] = true;
  }

  var subfolders = folder.getFolders();

  while (subfolders.hasNext()) {
    if (inventoryLimitReached_(scanState)) {
      return;
    }

    var subfolder = subfolders.next();

    scanInventoryFolder_(
      subfolder,
      folderPath +
        " / " +
        subfolder.getName(),
      headerMap,
      existingFileIds,
      scanState
    );
  }
}

function inventoryLimitReached_(scanState) {
  return (
    scanState.maxImages !== null &&
    scanState.foundCount >=
      scanState.maxImages
  );
}

function getInventoryWriterHeaderMap_(sheet) {
  var headers = sheet
    .getRange(
      1,
      1,
      1,
      getPhase1InventoryHeaders_().length
    )
    .getValues()[0];

  var headerMap = {};

  for (
    var i = 0;
    i < headers.length;
    i++
  ) {
    var headerName =
      String(headers[i]).trim();

    if (headerName) {
      headerMap[headerName] = i;
    }
  }

  return headerMap;
}

function validateInventoryWriterHeaders_(
  headerMap
) {
  var requiredHeaders =
    getPhase1InventoryHeaders_();

  for (
    var i = 0;
    i < requiredHeaders.length;
    i++
  ) {
    if (
      headerMap[
        requiredHeaders[i]
      ] === undefined
    ) {
      throw new Error(
        "Required spreadsheet column is missing: " +
        requiredHeaders[i]
      );
    }
  }
}

function getExistingInventoryFileIds_(
  sheet,
  headerMap
) {
  var existingFileIds = {};
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return existingFileIds;
  }

  var fileIdColumn =
    headerMap["Drive File ID"] + 1;

  var fileIds = sheet
    .getRange(
      2,
      fileIdColumn,
      lastRow - 1,
      1
    )
    .getValues();

  for (
    var i = 0;
    i < fileIds.length;
    i++
  ) {
    var fileId =
      String(fileIds[i][0]).trim();

    if (fileId) {
      existingFileIds[fileId] = true;
    }
  }

  return existingFileIds;
}

function createBlankInventoryRow_() {
  var headers = getPhase1InventoryHeaders_();
  var row = [];

  for (
    var i = 0;
    i < headers.length;
    i++
  ) {
    row.push("");
  }

  return row;
}

function setInventoryRowValue_(
  row,
  headerMap,
  headerName,
  value
) {
  if (headerMap[headerName] === undefined) {
    throw new Error(
      "Inventory column is missing: " +
      headerName
    );
  }

  row[
    headerMap[headerName]
  ] = value;
}

/**
 * Existing full workflow entry point retained.
 * Do not use for one-image or five-image tests.
 */
function runInventoryAutomation() {
  console.log(
    "SB Schirmer inventory automation started."
  );

  writeSupportedAssetInventory();

  SpreadsheetApp.flush();

  if (
    typeof analyzePendingImages ===
    "function"
  ) {
    analyzePendingImages();
  } else {
    throw new Error(
      "The analyzePendingImages function was not found."
    );
  }
}
