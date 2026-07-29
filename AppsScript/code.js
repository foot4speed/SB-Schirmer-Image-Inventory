function writeSupportedAssetInventory() {
  var folderId = "1ETQMtYzxlxzuIdLNi1DR_P5-QEwhH1NK";
  var spreadsheetId = "1KCxbVMjtumF9SwA88nM0SEtZjeseptVsA0NfTVt2lTM";
  var sheetName = "Image Inventory";

  var rootFolder = DriveApp.getFolderById(folderId);
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  var headers = [
    "Filename",
    "Asset Type",
    "MIME Type",
    "File Size",
    "Drive File ID",
    "Drive Link",
    "Source Folder",
    "Status"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  var existingIds = {};

  if (sheet.getLastRow() >= 2) {
    var idValues = sheet
      .getRange(2, 5, sheet.getLastRow() - 1, 1)
      .getValues();

    for (var i = 0; i < idValues.length; i++) {
      var existingId = String(idValues[i][0]).trim();

      if (existingId) {
        existingIds[existingId] = true;
      }
    }
  }

  var rows = [];

  scanFolderRecursively(rootFolder, rootFolder.getName(), existingIds, rows);

  if (rows.length > 0) {
    var startRow = sheet.getLastRow() + 1;

    sheet
      .getRange(startRow, 1, rows.length, headers.length)
      .setValues(rows);

    console.log("New supported assets added: " + rows.length);
      } else {
    console.log("No new supported assets found.");
  }

  console.log("SUPPORTED ASSET INVENTORY COMPLETED");
}


function scanFolderRecursively(folder, folderPath, existingIds, rows) {
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    var mimeType = file.getMimeType();
    var fileId = file.getId();

    var isImage = mimeType.indexOf("image/") === 0;
    var isVideo = mimeType.indexOf("video/") === 0;

    if (!isImage && !isVideo) {
      continue;
    }

    if (existingIds[fileId]) {
      continue;
    }

    rows.push([
      file.getName(),
      isImage ? "Image" : "Video",
      mimeType,
      file.getSize(),
      fileId,
      file.getUrl(),
      folderPath,
      "Pending Analysis"
    ]);

    existingIds[fileId] = true;
  }

  var subfolders = folder.getFolders();

  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    var subfolderPath = folderPath + " / " + subfolder.getName();

    scanFolderRecursively(
      subfolder,
      subfolderPath,
      existingIds,
      rows
    );
  }
}

function scheduleNextAnalysisBatch() {
  // Remove duplicates before creating a new trigger.
  deleteAnalysisTriggers();

  ScriptApp.newTrigger("analyzePendingImages")
    .timeBased()
    .after(60 * 1000) // Run again in approximately one minute
    .create();
}


function deleteAnalysisTriggers() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "analyzePendingImages") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
function listFilesFoundInFolder() {
  var folderId = "1ETQMtYzxlxzuIdLNi1DR_P5-QEwhH1NK";
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();

  var count = 0;

  console.log("Scanning folder: " + folder.getName());
  console.log("Folder URL: " + folder.getUrl());

  while (files.hasNext()) {
    var file = files.next();
    count++;

    console.log(
      count +
      " | " +
      file.getName() +
      " | " +
      file.getMimeType() +
      " | " +
      file.getId()
    );
  }

  console.log("Total files found directly in folder: " + count);
}

    