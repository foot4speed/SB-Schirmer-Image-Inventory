/**
 * SB Schirmer Image Inventory Engine
 * Phase 1: Image analysis
 */
function recoverStaleAnalyzingRows_() {
  var sheet = getInventorySheet_();
  var headerMap = getPhase1HeaderMap_(sheet);

  validateAnalysisHeaders_(headerMap);

  var lastRow = sheet.getLastRow();
  var recoveredCount = 0;

  for (var sheetRow = 2; sheetRow <= lastRow; sheetRow++) {
    var status = String(
      sheet
        .getRange(
          sheetRow,
          headerMap["Analysis Status"] + 1
        )
        .getValue()
    ).trim();

    if (status !== "Analyzing") {
      continue;
    }

    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Analysis Status",
      "Pending Analysis"
    );

    recoveredCount++;
  }

  SpreadsheetApp.flush();

  console.log(
    "Stale analyzing rows recovered: " +
    recoveredCount
  );

  return recoveredCount;
}

function recoverStaleAnalyzingRows() {
  return recoverStaleAnalyzingRows_();
}

function recoverInsufficientQuotaErrors() {
  var sheet = getInventorySheet_();
  var headerMap = getPhase1HeaderMap_(sheet);

  validateAnalysisHeaders_(headerMap);

  var lastRow = sheet.getLastRow();
  var recoveredCount = 0;

  for (var sheetRow = 2; sheetRow <= lastRow; sheetRow++) {
    var status = String(
      sheet
        .getRange(
          sheetRow,
          headerMap["Analysis Status"] + 1
        )
        .getValue()
    ).trim();

    if (status !== "Error") {
      continue;
    }

    var notes = String(
      sheet
        .getRange(
          sheetRow,
          headerMap["Notes"] + 1
        )
        .getValue()
    ).trim();

    if (notes.indexOf("insufficient_quota") === -1) {
      continue;
    }

    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Analysis Status",
      "Pending Analysis"
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Confidence",
      ""
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Notes",
      ""
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Date Analyzed",
      ""
    );

    recoveredCount++;
  }

  SpreadsheetApp.flush();

  console.log(
    "Insufficient quota errors recovered: " +
    recoveredCount
  );

  return recoveredCount;
}

function reclassifyReviewRequiredRows() {
  var sheet = getInventorySheet_();
  var headerMap = getPhase1HeaderMap_(sheet);

  validateAnalysisHeaders_(headerMap);

  var lastRow = sheet.getLastRow();
  var reclassifiedCount = 0;

  for (var sheetRow = 2; sheetRow <= lastRow; sheetRow++) {
    var status = String(
      sheet
        .getRange(
          sheetRow,
          headerMap["Analysis Status"] + 1
        )
        .getValue()
    ).trim();

    if (status !== "Review Required") {
      continue;
    }

    var confidence = String(
      sheet
        .getRange(
          sheetRow,
          headerMap["Confidence"] + 1
        )
        .getValue()
    ).trim();

    if (confidence === "Low") {
      continue;
    }

    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Analysis Status",
      "Complete"
    );

    reclassifiedCount++;
  }

  SpreadsheetApp.flush();

  console.log(
    "Review Required rows reclassified to Complete: " +
    reclassifiedCount
  );

  return reclassifiedCount;
}


function diagnoseAnalysisErrors() {
  var sheet = getInventorySheet_();
  var headerMap = getPhase1HeaderMap_(sheet);

  validateAnalysisHeaders_(headerMap);

  var lastRow = sheet.getLastRow();

  var counts = {
    totalRows: 0,
    complete: 0,
    reviewRequired: 0,
    error: 0,
    skipped: 0,
    pendingAnalysis: 0,
    analyzing: 0,
    other: 0
  };

  var errorGroups = {};

  for (var sheetRow = 2; sheetRow <= lastRow; sheetRow++) {
    counts.totalRows++;

    var status = String(
      sheet
        .getRange(
          sheetRow,
          headerMap["Analysis Status"] + 1
        )
        .getValue()
    ).trim();

    if (status === "Complete") {
      counts.complete++;
    } else if (status === "Review Required") {
      counts.reviewRequired++;
    } else if (status === "Error") {
      counts.error++;

      var notes = String(
        sheet
          .getRange(
            sheetRow,
            headerMap["Notes"] + 1
          )
          .getValue()
      ).trim();

      var group = notes || "(blank error notes)";

      if (group.length > 120) {
        group = group.substring(0, 120);
      }

      errorGroups[group] =
        (errorGroups[group] || 0) + 1;
    } else if (status === "Skipped") {
      counts.skipped++;
    } else if (status === "Pending Analysis") {
      counts.pendingAnalysis++;
    } else if (status === "Analyzing") {
      counts.analyzing++;
    } else {
      counts.other++;
    }
  }

  console.log("Total inventory rows: " + counts.totalRows);
  console.log("Complete: " + counts.complete);
  console.log("Review Required: " + counts.reviewRequired);
  console.log("Error: " + counts.error);
  console.log("Skipped: " + counts.skipped);
  console.log("Pending Analysis: " + counts.pendingAnalysis);
  console.log("Analyzing: " + counts.analyzing);
  console.log("Other: " + counts.other);

  for (var key in errorGroups) {
    if (errorGroups.hasOwnProperty(key)) {
      console.log(
        "ERROR GROUP (" +
        errorGroups[key] +
        "): " +
        key
      );
    }
  }

  return {
    counts: counts,
    errorGroups: errorGroups
  };
}

function diagnoseReviewRequiredReasons() {
  var sheet = getInventorySheet_();
  var headerMap = getPhase1HeaderMap_(sheet);

  validateAnalysisHeaders_(headerMap);

  var lastRow = sheet.getLastRow();

  var counts = {
    totalReviewRequired: 0,
    lowConfidenceOnly: 0,
    unknownRecipeOnly: 0,
    both: 0
  };

  var qualityCounts = {
    Excellent: 0,
    Good: 0,
    Usable: 0,
    Poor: 0,
    Other: 0
  };

  for (var sheetRow = 2; sheetRow <= lastRow; sheetRow++) {
    var status = String(
      sheet
        .getRange(
          sheetRow,
          headerMap["Analysis Status"] + 1
        )
        .getValue()
    ).trim();

    if (status !== "Review Required") {
      continue;
    }

    counts.totalReviewRequired++;

    var confidence = String(
      sheet
        .getRange(
          sheetRow,
          headerMap["Confidence"] + 1
        )
        .getValue()
    ).trim();

    var recipe = String(
      sheet
        .getRange(
          sheetRow,
          headerMap["Recipe"] + 1
        )
        .getValue()
    ).trim();

    var quality = String(
      sheet
        .getRange(
          sheetRow,
          headerMap["Quality"] + 1
        )
        .getValue()
    ).trim();

    var isLowConfidence =
      confidence === "Low";

    var isUnknownRecipe =
      recipe === "Unknown";

    if (
      isLowConfidence &&
      isUnknownRecipe
    ) {
      counts.both++;
    } else if (isLowConfidence) {
      counts.lowConfidenceOnly++;
    } else if (isUnknownRecipe) {
      counts.unknownRecipeOnly++;
    }

    if (
      qualityCounts[quality] !==
      undefined
    ) {
      qualityCounts[quality]++;
    } else {
      qualityCounts.Other++;
    }
  }

  console.log(
    "Review Required total: " +
    counts.totalReviewRequired
  );

  console.log(
    "Low confidence only: " +
    counts.lowConfidenceOnly
  );

  console.log(
    "Unknown recipe only: " +
    counts.unknownRecipeOnly
  );

  console.log(
    "Both Low confidence and Unknown recipe: " +
    counts.both
  );

  console.log(
    "Review Required quality - Excellent: " +
    qualityCounts.Excellent
  );

  console.log(
    "Review Required quality - Good: " +
    qualityCounts.Good
  );

  console.log(
    "Review Required quality - Usable: " +
    qualityCounts.Usable
  );

  console.log(
    "Review Required quality - Poor: " +
    qualityCounts.Poor
  );

  console.log(
    "Review Required quality - Other: " +
    qualityCounts.Other
  );

  return {
    counts: counts,
    qualityCounts: qualityCounts
  };
}

function analyzePendingImages() {
  return analyzePendingImagesWithOptions_(
    SB_CONFIG.ANALYSIS_BATCH_SIZE,
    true
  );
}

function analyzePendingImagesWithOptions_(
  maxRows,
  allowContinuation
) {
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    console.log(
      "Another analysis batch is already running."
    );
    return 0;
  }

  try {
    var sheet = getInventorySheet_();
    var headerMap = getPhase1HeaderMap_(sheet);

    validateAnalysisHeaders_(headerMap);

    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      console.log(
        "There are no inventory rows to analyze."
      );

      if (allowContinuation) {
        deleteAnalysisTriggers();
      }

      return 0;
    }

    var processedCount = 0;
    var pendingRowsRemain = false;

    for (
      var sheetRow = 2;
      sheetRow <= lastRow;
      sheetRow++
    ) {
      var status = String(
        sheet
          .getRange(
            sheetRow,
            headerMap["Analysis Status"] + 1
          )
          .getValue()
      ).trim();

      if (status !== "Pending Analysis") {
        continue;
      }

      if (processedCount >= maxRows) {
        pendingRowsRemain = true;
        break;
      }

      analyzeInventoryRowSafely_(
        sheet,
        sheetRow,
        headerMap
      );

      processedCount++;
    }

    SpreadsheetApp.flush();

    console.log(
      "Images processed in this batch: " +
      processedCount
    );

    if (allowContinuation) {
      if (!pendingRowsRemain) {
        pendingRowsRemain =
          hasPendingAnalysisRows_(sheet);
      }

      if (pendingRowsRemain) {
        console.log(
          "More pending images remain. Scheduling next batch."
        );
        scheduleNextAnalysisBatch();
      } else {
        console.log(
          "No pending images remain."
        );
        deleteAnalysisTriggers();
      }
    }

    return processedCount;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Analyzes only the rows inserted by a limited test.
 * No continuation trigger is created.
 */
function analyzeSpecificInventoryRows_(
  firstRow,
  rowCount
) {
  if (!firstRow || rowCount <= 0) {
    return 0;
  }

  var lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    console.log(
      "Another analysis batch is already running."
    );
    return 0;
  }

  try {
    var sheet = getInventorySheet_();
    var headerMap = getPhase1HeaderMap_(sheet);

    validateAnalysisHeaders_(headerMap);

    var processedCount = 0;
    var lastRow = Math.min(
      sheet.getLastRow(),
      firstRow + rowCount - 1
    );

    for (
      var sheetRow = firstRow;
      sheetRow <= lastRow;
      sheetRow++
    ) {
      var status = String(
        sheet
          .getRange(
            sheetRow,
            headerMap["Analysis Status"] + 1
          )
          .getValue()
      ).trim();

      if (status !== "Pending Analysis") {
        continue;
      }

      analyzeInventoryRowSafely_(
        sheet,
        sheetRow,
        headerMap
      );

      processedCount++;
    }

    SpreadsheetApp.flush();

    console.log(
      "Test images analyzed: " +
      processedCount
    );

    return processedCount;
  } finally {
    lock.releaseLock();
  }
}

function analyzeInventoryRowSafely_(
  sheet,
  sheetRow,
  headerMap
) {
  try {
    analyzeInventoryRow_(
      sheet,
      sheetRow,
      headerMap
    );
  } catch (error) {
    console.error(
      "Analysis failed for row " +
      sheetRow +
      ": " +
      error.message
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Analysis Status",
      "Error"
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Confidence",
      "Low"
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Notes",
      truncateText_(
        error.message,
        250
      )
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Date Analyzed",
      new Date()
    );
  }
}

function analyzeInventoryRow_(
  sheet,
  sheetRow,
  headerMap
) {
  var filename = String(
    sheet
      .getRange(
        sheetRow,
        headerMap["Filename"] + 1
      )
      .getValue()
  ).trim();

  var fileId = String(
    sheet
      .getRange(
        sheetRow,
        headerMap["Drive File ID"] + 1
      )
      .getValue()
  ).trim();

  if (!fileId) {
    throw new Error(
      "The Drive File ID is missing."
    );
  }

  var file = DriveApp.getFileById(fileId);
  var mimeType = String(
    file.getMimeType() || ""
  );

  if (mimeType.indexOf("image/") !== 0) {
    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Analysis Status",
      "Skipped"
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Confidence",
      "Low"
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Notes",
      "Drive file is not an image."
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headerMap,
      "Date Analyzed",
      new Date()
    );

    return;
  }

  var maxBytes =
    SB_CONFIG.MAX_FILE_SIZE_MB *
    1024 *
    1024;

  if (
    SB_CONFIG.MAX_FILE_SIZE_MB &&
    file.getSize() > maxBytes
  ) {
    throw new Error(
      "Image exceeds the configured maximum file size of " +
      SB_CONFIG.MAX_FILE_SIZE_MB +
      " MB."
    );
  }

  setCellByHeader_(
    sheet,
    sheetRow,
    headerMap,
    "Analysis Status",
    "Analyzing"
  );

  SpreadsheetApp.flush();

  var analysis = callOpenAIImageAnalysis_(
    file.getBlob(),
    filename
  );

  writeAnalysisResults_(
    sheet,
    sheetRow,
    headerMap,
    analysis
  );

  var finalStatus =
  analysis.confidence === "Low"
    ? "Review Required"
    : "Complete";

  setCellByHeader_(
    sheet,
    sheetRow,
    headerMap,
    "Analysis Status",
    finalStatus
  );

  setCellByHeader_(
    sheet,
    sheetRow,
    headerMap,
    "Date Analyzed",
    new Date()
  );

  console.log(
    "Completed analysis: " + filename
  );
}

/**
 * Retries once if the model response cannot be validated.
 */
function callOpenAIImageAnalysis_(
  imageBlob,
  filename
) {
  var lastError = null;

  for (var attempt = 1; attempt <= 2; attempt++) {
    try {
      return callOpenAIImageAnalysisOnce_(
        imageBlob,
        filename
      );
    } catch (error) {
      lastError = error;

      if (attempt < 2) {
        console.log(
          "Image analysis attempt failed; retrying once: " +
          error.message
        );
      }
    }
  }

  throw lastError;
}

function callOpenAIImageAnalysisOnce_(
  imageBlob,
  filename
) {
  var apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("OPENAI_API_KEY");

  if (!apiKey) {
    throw new Error(
      "OPENAI_API_KEY was not found in Script Properties."
    );
  }

  var mimeType =
    imageBlob.getContentType() ||
    "image/jpeg";

  var base64Image =
    Utilities.base64Encode(
      imageBlob.getBytes()
    );

  var dataUrl =
    "data:" +
    mimeType +
    ";base64," +
    base64Image;

  var prompt = [
    "Analyze this image for the SB Schirmer food and product image inventory.",
    "Be conservative. Do not invent details that are not visible or strongly supported.",
    "Use Unknown when the exact recipe or classification cannot be supported.",
    "Ingredients must contain only visible or strongly supported ingredients as a comma-separated string.",
    "Notes must be concise.",
    "Filename: " + filename
  ].join("\n");

  var analysisSchema = {
    type: "object",
    properties: {
      recipe: {
        type: "string"
      },
      meal: {
        type: "string",
        enum: [
          "Breakfast",
          "Lunch",
          "Dinner",
          "Snacks",
          "Hors d'oeuvres",
          "Multiple",
          "Unknown"
        ]
      },
      ingredients: {
        type: "string"
      },
      uses_nicer_slicer: {
        type: "string",
        enum: [
          "Yes",
          "Likely",
          "No",
          "Unknown"
        ]
      },
      quality: {
        type: "string",
        enum: [
          "Excellent",
          "Good",
          "Usable",
          "Poor"
        ]
      },
      orientation: {
        type: "string",
        enum: [
          "Landscape",
          "Portrait",
          "Square"
        ]
      },
      confidence: {
        type: "string",
        enum: [
          "High",
          "Medium",
          "Low"
        ]
      },
      notes: {
        type: "string"
      }
    },
    required: [
      "recipe",
      "meal",
      "ingredients",
      "uses_nicer_slicer",
      "quality",
      "orientation",
      "confidence",
      "notes"
    ],
    additionalProperties: false
  };

  var payload = {
    model: SB_CONFIG.OPENAI_MODEL,
    input: [
      {
        role: "user",
        content: [
          {
            type: "input_text",
            text: prompt
          },
          {
            type: "input_image",
            image_url: dataUrl,
            detail: "high"
          }
        ]
      }
    ],
    text: {
      format: {
        type: "json_schema",
        name: "sb_schirmer_image_analysis",
        strict: true,
        schema: analysisSchema
      }
    },
    max_output_tokens: 4000
  };

  var response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/responses",
    {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  var responseCode =
    response.getResponseCode();

  var responseText =
    response.getContentText();

  if (
    responseCode < 200 ||
    responseCode >= 300
  ) {
    throw new Error(
      "OpenAI API returned HTTP " +
      responseCode +
      ": " +
      truncateText_(responseText, 500)
    );
  }

  var responseJson;

  try {
    responseJson = JSON.parse(responseText);
  } catch (error) {
    throw new Error(
      "The OpenAI API response was not valid JSON."
    );
  }

  if (responseJson.status === "incomplete") {
    var incompleteReason = responseJson.incomplete_details && responseJson.incomplete_details.reason ? responseJson.incomplete_details.reason : "unknown reason";
    throw new Error("OpenAI response was incomplete: " + incompleteReason);
  }

  var outputText =
    extractResponseOutputText_(
      responseJson
    );

  if (!outputText) {
    throw new Error(
      "OpenAI returned no analysis text."
    );
  }

  return parseAndValidateAnalysisJson_(
    outputText
  );
}

function extractResponseOutputText_(
  responseJson
) {
  if (
    responseJson.output_text &&
    typeof responseJson.output_text ===
      "string"
  ) {
    return responseJson.output_text.trim();
  }

  var collectedText = [];
  var output = responseJson.output || [];

  for (
    var outputIndex = 0;
    outputIndex < output.length;
    outputIndex++
  ) {
    var item = output[outputIndex];
    var content = item.content || [];

    for (
      var contentIndex = 0;
      contentIndex < content.length;
      contentIndex++
    ) {
      var contentItem = content[contentIndex];

      if (
        contentItem.type ===
          "output_text" &&
        contentItem.text
      ) {
        collectedText.push(
          contentItem.text
        );
      }
    }
  }

  return collectedText.join("\n").trim();
}

function parseAndValidateAnalysisJson_(
  text
) {
  var parsed;

  try {
    parsed = JSON.parse(
      String(text).trim()
    );
  } catch (error) {
    throw new Error(
      "Could not parse the image analysis JSON. Output preview: " + truncateText_(String(text).trim(), 300)
    );
  }

  var result = {
    recipe: safeText_(
      parsed.recipe,
      "Unknown"
    ),
    meal: safeText_(
      parsed.meal,
      "Unknown"
    ),
    ingredients: safeText_(
      parsed.ingredients,
      "Unknown"
    ),
    uses_nicer_slicer: safeText_(
      parsed.uses_nicer_slicer,
      "Unknown"
    ),
    quality: safeText_(
      parsed.quality,
      "Poor"
    ),
    orientation: safeText_(
      parsed.orientation,
      "Unknown"
    ),
    confidence: safeText_(
      parsed.confidence,
      "Low"
    ),
    notes: safeText_(
      parsed.notes,
      ""
    )
  };

  validateAllowedValue_(
    "meal",
    result.meal,
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

  validateAllowedValue_(
    "uses_nicer_slicer",
    result.uses_nicer_slicer,
    ["Yes", "Likely", "No", "Unknown"]
  );

  validateAllowedValue_(
    "quality",
    result.quality,
    ["Excellent", "Good", "Usable", "Poor"]
  );

  validateAllowedValue_(
    "orientation",
    result.orientation,
    ["Landscape", "Portrait", "Square"]
  );

  validateAllowedValue_(
    "confidence",
    result.confidence,
    ["High", "Medium", "Low"]
  );

  return result;
}

function validateAllowedValue_(
  fieldName,
  value,
  allowedValues
) {
  if (allowedValues.indexOf(value) === -1) {
    throw new Error(
      "AI response contained an invalid " +
      fieldName +
      " value: " +
      value
    );
  }
}

function writeAnalysisResults_(
  sheet,
  sheetRow,
  headerMap,
  analysis
) {
  setCellByHeader_(
    sheet,
    sheetRow,
    headerMap,
    "Recipe",
    analysis.recipe
  );

  setCellByHeader_(
    sheet,
    sheetRow,
    headerMap,
    "Meal",
    analysis.meal
  );

  setCellByHeader_(
    sheet,
    sheetRow,
    headerMap,
    "Ingredients",
    analysis.ingredients
  );

  setCellByHeader_(
    sheet,
    sheetRow,
    headerMap,
    "Uses Nicer Slicer",
    analysis.uses_nicer_slicer
  );

  setCellByHeader_(
    sheet,
    sheetRow,
    headerMap,
    "Quality",
    analysis.quality
  );

  setCellByHeader_(
    sheet,
    sheetRow,
    headerMap,
    "Orientation",
    analysis.orientation
  );

  setCellByHeader_(
    sheet,
    sheetRow,
    headerMap,
    "Confidence",
    analysis.confidence
  );

  setCellByHeader_(
    sheet,
    sheetRow,
    headerMap,
    "Notes",
    analysis.notes
  );
}

function getInventorySheet_() {
  var spreadsheet =
    SpreadsheetApp.openById(
      SB_CONFIG.SPREADSHEET_ID
    );

  var sheet =
    spreadsheet.getSheetByName(
      SB_CONFIG.INVENTORY_SHEET_NAME
    );

  if (!sheet) {
    throw new Error(
      'Sheet "' +
      SB_CONFIG.INVENTORY_SHEET_NAME +
      '" was not found.'
    );
  }

  return sheet;
}

function getPhase1HeaderMap_(sheet) {
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

function validateAnalysisHeaders_(
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

function hasPendingAnalysisRows_(sheet) {
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return false;
  }

  var headerMap =
    getPhase1HeaderMap_(sheet);

  var statuses = sheet
    .getRange(
      2,
      headerMap["Analysis Status"] + 1,
      lastRow - 1,
      1
    )
    .getValues();

  for (
    var i = 0;
    i < statuses.length;
    i++
  ) {
    if (
      String(statuses[i][0]).trim() ===
      "Pending Analysis"
    ) {
      return true;
    }
  }

  return false;
}

function setCellByHeader_(
  sheet,
  sheetRow,
  headerMap,
  headerName,
  value
) {
  if (headerMap[headerName] === undefined) {
    throw new Error(
      "Analysis column is missing: " +
      headerName
    );
  }

  sheet
    .getRange(
      sheetRow,
      headerMap[headerName] + 1
    )
    .setValue(value);
}

function safeText_(
  value,
  fallback
) {
  if (
    value === null ||
    value === undefined
  ) {
    return fallback;
  }

  var text = String(value).trim();

  return text || fallback;
}

function truncateText_(
  text,
  maximumLength
) {
  var value = String(text || "");

  if (value.length <= maximumLength) {
    return value;
  }

  return value.substring(
    0,
    maximumLength
  ) + "...";
}
