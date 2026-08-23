/**
 * SB Schirmer Image Inventory Engine
 * Phase 1 test harness
 *
 * These tests automatically discover new images in the configured
 * Google Drive folder. No manual file ID is required.
 */

function testSingleImage() {
  return runLimitedInventoryTest_(1);
}

function testFiveImages() {
  return runLimitedInventoryTest_(5);
}

function runLimitedInventoryTest_(imageLimit) {
  console.log(
    "Starting limited inventory test for up to " +
    imageLimit +
    " image(s)."
  );

  deleteAnalysisTriggers();
  setupInventorySheet();

  var inventoryResult =
    writeSupportedAssetInventoryWithLimit_(imageLimit);

  if (
    !inventoryResult ||
    inventoryResult.addedCount === 0
  ) {
    console.log(
      "No new eligible images were found for this test."
    );

    return {
      inventoried: 0,
      analyzed: 0
    };
  }

  var analyzedCount =
    analyzeSpecificInventoryRows_(
      inventoryResult.firstRow,
      inventoryResult.addedCount
    );

  console.log(
    "Limited inventory test finished. Inventoried: " +
    inventoryResult.addedCount +
    "; analyzed: " +
    analyzedCount
  );

  return {
    inventoried: inventoryResult.addedCount,
    analyzed: analyzedCount
  };
}

function testInsufficientQuotaSafeStop() {
  var sheet = getInventorySheet_();
  var headerMap = getPhase1HeaderMap_(sheet);

  validateAnalysisHeaders_(headerMap);

  var lastRow = sheet.getLastRow();
  var testRow = null;

  for (var sheetRow = 2; sheetRow <= lastRow; sheetRow++) {
    var status = String(
      sheet
        .getRange(
          sheetRow,
          headerMap["Analysis Status"] + 1
        )
        .getValue()
    ).trim();

    if (status === "Pending Analysis") {
      testRow = sheetRow;
      break;
    }
  }

  if (!testRow) {
    throw new Error(
      "No Pending Analysis row is available for the quota-stop test."
    );
  }

  var originalAnalyzeInventoryRow =
    analyzeInventoryRow_;

  try {
    analyzeInventoryRow_ = function() {
      throw new Error(
        'OpenAI API returned HTTP 429: {"error":{"type":"insufficient_quota"}}'
      );
    };

    analyzeInventoryRowSafely_(
      sheet,
      testRow,
      headerMap
    );
  } finally {
    analyzeInventoryRow_ =
      originalAnalyzeInventoryRow;
  }
}
