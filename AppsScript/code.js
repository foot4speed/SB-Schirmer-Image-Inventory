/**
 * SB Schirmer Image Inventory Engine
 * Shared automation and diagnostic utilities.
 *
 * The original prototype inventory scanner was removed during the
 * configuration refactor. The active scanner is in inventory.js.
 */

/**
 * Schedules the next image-analysis batch.
 */
function scheduleNextAnalysisBatch() {
  deleteAnalysisTriggers();

  ScriptApp.newTrigger("analyzePendingImages")
    .timeBased()
    .after(60 * 1000)
    .create();
}


/**
 * Removes triggers that run the image-analysis batch.
 */
function deleteAnalysisTriggers() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    if (
      triggers[i].getHandlerFunction() ===
      "analyzePendingImages"
    ) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


/**
 * Diagnostic helper: lists files directly inside the configured root folder.
 */
function listFilesFoundInFolder() {
  var folder = DriveApp.getFolderById(
    SB_CONFIG.ROOT_FOLDER_ID
  );

  var files = folder.getFiles();
  var count = 0;

  console.log(
    "Scanning folder: " + folder.getName()
  );
  console.log(
    "Folder URL: " + folder.getUrl()
  );

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

  console.log(
    "Total files found directly in folder: " +
    count
  );
}
