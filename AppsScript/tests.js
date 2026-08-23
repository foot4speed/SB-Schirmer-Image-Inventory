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

    if (status === "Complete") {
      testRow = sheetRow;
      break;
    }
  }

  if (!testRow) {
    throw new Error(
      "No Complete row is available for the quota-stop test."
    );
  }

  var originalStatus = sheet
    .getRange(
      testRow,
      headerMap["Analysis Status"] + 1
    )
    .getValue();

  var originalConfidence = sheet
    .getRange(
      testRow,
      headerMap["Confidence"] + 1
    )
    .getValue();

  var originalNotes = sheet
    .getRange(
      testRow,
      headerMap["Notes"] + 1
    )
    .getValue();

  var originalDateAnalyzed = sheet
    .getRange(
      testRow,
      headerMap["Date Analyzed"] + 1
    )
    .getValue();

  var originalAnalyzeInventoryRow =
    analyzeInventoryRow_;

  try {
    setCellByHeader_(
      sheet,
      testRow,
      headerMap,
      "Analysis Status",
      "Pending Analysis"
    );

    analyzeInventoryRow_ = function() {
      throw new Error(
        'OpenAI API returned HTTP 429: {"error":{"type":"insufficient_quota"}}'
      );
    };

    try {
      analyzeInventoryRowSafely_(
        sheet,
        testRow,
        headerMap
      );
    } catch (error) {
      var resultingStatus = String(
        sheet
          .getRange(
            testRow,
            headerMap["Analysis Status"] + 1
          )
          .getValue()
      ).trim();

      console.log(
        "Quota safe-stop test row status: " +
        resultingStatus
      );

      if (resultingStatus !== "Pending Analysis") {
        throw new Error(
          "Quota safe-stop test failed. Expected Pending Analysis."
        );
      }

      console.log(
        "Quota safe-stop test passed."
      );
    }
  } finally {
    analyzeInventoryRow_ =
      originalAnalyzeInventoryRow;

    setCellByHeader_(
      sheet,
      testRow,
      headerMap,
      "Analysis Status",
      originalStatus
    );

    setCellByHeader_(
      sheet,
      testRow,
      headerMap,
      "Confidence",
      originalConfidence
    );

    setCellByHeader_(
      sheet,
      testRow,
      headerMap,
      "Notes",
      originalNotes
    );

    setCellByHeader_(
      sheet,
      testRow,
      headerMap,
      "Date Analyzed",
      originalDateAnalyzed
    );

    SpreadsheetApp.flush();
  }
}
