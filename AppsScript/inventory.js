/**
 * SB Schirmer Image Inventory Engine
 * Phase 1: Inventory Writer
 *
 * Scans the selected Google Drive folder and all subfolders.
 * Writes one spreadsheet row per image.
 *
 * This script does not move, rename, delete, edit, duplicate,
 * or change permissions on any original Drive files.
 */

var INVENTORY_ROOT_FOLDER_ID =
  "1ETQMtYzxlxzuIdLNi1DR_P5-QEwhH1NK";

var INVENTORY_SPREADSHEET_ID =
  "1KCxbVMjtumF9SwA88nM0SEtZjeseptVsA0NfTVt2lTM";

var INVENTORY_SHEET_NAME =
  "Image Inventory";


/**
 * Main function shown in the Apps Script function menu.
 */
function writeSupportedAssetInventory() {
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    console.log(
      "Another inventory scan is already running."
    );
    return;
  }

  try {
    var sheet =
      getInventoryWriterSheet_();

    ensureInventoryWriterHeaders_(sheet);

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
        INVENTORY_ROOT_FOLDER_ID
      );

    var newRows = [];

    scanInventoryFolder_(
      rootFolder,
      rootFolder.getName(),
      headerMap,
      existingFileIds,
      newRows
    );

    if (newRows.length === 0) {
      console.log(
        "Inventory scan complete. No new images were found."
      );
      return;
    }

    var firstNewRow =
      sheet.getLastRow() + 1;

    sheet
      .getRange(
        firstNewRow,
        1,
        newRows.length,
        sheet.getLastColumn()
      )
      .setValues(newRows);

    SpreadsheetApp.flush();

    console.log(
      "Inventory scan complete. New images added: " +
      newRows.length
    );
  } finally {
    lock.releaseLock();
  }
}


/**
 * Recursively scans one folder.
 */
function scanInventoryFolder_(
  folder,
  folderPath,
  headerMap,
  existingFileIds,
  newRows
) {
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();

    var mimeType = String(
      file.getMimeType() || ""
    );

    if (
      mimeType.indexOf("image/") !== 0
    ) {
      continue;
    }

    var fileId = file.getId();

    if (existingFileIds[fileId]) {
      continue;
    }

    var row =
      createBlankInventoryRow_(
        headerMap
      );

    setInventoryRowValue_(
      row,
      headerMap,
      "Filename",
      file.getName()
    );

    setInventoryRowValue_(
      row,
      headerMap,
      "Asset Type",
      "Image"
    );

    setInventoryRowValue_(
      row,
      headerMap,
      "MIME Type",
      mimeType
    );

    setInventoryRowValue_(
      row,
      headerMap,
      "File Size",
      file.getSize()
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
      "Status",
      "Pending Analysis"
    );

    newRows.push(row);
    existingFileIds[fileId] = true;
  }

  var subfolders =
    folder.getFolders();

  while (subfolders.hasNext()) {
    var subfolder =
      subfolders.next();

    scanInventoryFolder_(
      subfolder,
      folderPath +
        " / " +
        subfolder.getName(),
      headerMap,
      existingFileIds,
      newRows
    );
  }
}


/**
 * Returns the inventory sheet.
 */
function getInventoryWriterSheet_() {
  var spreadsheet =
    SpreadsheetApp.openById(
      INVENTORY_SPREADSHEET_ID
    );

  var sheet =
    spreadsheet.getSheetByName(
      INVENTORY_SHEET_NAME
    );

  if (!sheet) {
    throw new Error(
      'Sheet "' +
      INVENTORY_SHEET_NAME +
      '" was not found.'
    );
  }

  return sheet;
}


/**
 * Adds the required inventory headers when missing.
 */
function ensureInventoryWriterHeaders_(
  sheet
) {
  var requiredHeaders = [
    "Filename",
    "Asset Type",
    "MIME Type",
    "File Size",
    "Drive File ID",
    "Drive Link",
    "Source Folder",
    "Status"
  ];

  var lastColumn =
    sheet.getLastColumn();

  if (lastColumn === 0) {
    sheet
      .getRange(
        1,
        1,
        1,
        requiredHeaders.length
      )
      .setValues([requiredHeaders]);

    sheet
      .getRange(
        1,
        1,
        1,
        requiredHeaders.length
      )
      .setFontWeight("bold");

    return;
  }

  var existingHeaders =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0];

  var existingHeaderNames = {};

  for (
    var i = 0;
    i < existingHeaders.length;
    i++
  ) {
    var headerName =
      String(
        existingHeaders[i]
      ).trim();

    if (headerName) {
      existingHeaderNames[
        headerName
      ] = true;
    }
  }

  var missingHeaders = [];

  for (
    var j = 0;
    j < requiredHeaders.length;
    j++
  ) {
    if (
      !existingHeaderNames[
        requiredHeaders[j]
      ]
    ) {
      missingHeaders.push(
        requiredHeaders[j]
      );
    }
  }

  if (missingHeaders.length > 0) {
    sheet
      .getRange(
        1,
        lastColumn + 1,
        1,
        missingHeaders.length
      )
      .setValues([missingHeaders]);

    sheet
      .getRange(
        1,
        lastColumn + 1,
        1,
        missingHeaders.length
      )
      .setFontWeight("bold");
  }
}


/**
 * Creates a header-name-to-column map.
 *
 * Values are zero-based array positions.
 */
function getInventoryWriterHeaderMap_(
  sheet
) {
  var headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
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


/**
 * Confirms that the necessary columns exist.
 */
function validateInventoryWriterHeaders_(
  headerMap
) {
  var requiredHeaders = [
    "Filename",
    "Asset Type",
    "MIME Type",
    "File Size",
    "Drive File ID",
    "Drive Link",
    "Source Folder",
    "Status"
  ];

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


/**
 * Reads existing Drive File IDs so files are not duplicated.
 */
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

  var fileIds =
    sheet
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
      existingFileIds[fileId] =
        true;
    }
  }

  return existingFileIds;
}


/**
 * Creates a blank row matching the sheet's full width.
 */
function createBlankInventoryRow_(
  headerMap
) {
  var largestColumnIndex = 0;

  for (
    var headerName in headerMap
  ) {
    if (
      headerMap.hasOwnProperty(
        headerName
      ) &&
      headerMap[headerName] >
        largestColumnIndex
    ) {
      largestColumnIndex =
        headerMap[headerName];
    }
  }

  var row = [];

  for (
    var i = 0;
    i <= largestColumnIndex;
    i++
  ) {
    row.push("");
  }

  return row;
}


/**
 * Places one value in the correct row position.
 */
function setInventoryRowValue_(
  row,
  headerMap,
  headerName,
  value
) {
  if (
    headerMap[headerName] === undefined
  ) {
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
 * Runs the complete Phase 1 workflow.
 *
 * 1. Inventories new Drive images.
 * 2. Starts analysis of Pending Analysis rows.
 * 3. Analysis automatically continues in batches until finished.
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