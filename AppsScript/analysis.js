/**
 * SB Schirmer Image Inventory Engine
 * Phase 1: Image analysis
 *
 * Reads rows marked "Pending Analysis", sends each image to OpenAI,
 * writes the analysis into the spreadsheet, and marks the row Complete.
 */

/**
 * Configuration is defined once in Config.js.
 */

/**
 * Processes the next batch of rows marked Pending Analysis.
 */
function analyzePendingImages() {
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    console.log("Another analysis batch is already running.");
    return;
  }

  try {
    var sheet = getInventorySheet_();
    ensureAnalysisHeaders_(sheet);

    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      console.log("There are no inventory rows to analyze.");
      deleteAnalysisTriggers();
      return;
    }

    var lastColumn = sheet.getLastColumn();
    var values = sheet
      .getRange(1, 1, lastRow, lastColumn)
      .getValues();

    var headerMap = createHeaderMap_(values[0]);

    validateRequiredHeaders_(headerMap);

    var processedCount = 0;
    var pendingRowsRemain = false;

    for (var rowIndex = 1; rowIndex < values.length; rowIndex++) {
      var row = values[rowIndex];

      var status = String(
        row[headerMap["Status"]]
      ).trim();

      if (status !== "Pending Analysis") {
        continue;
      }

      if (processedCount >= SB_CONFIG.ANALYSIS_BATCH_SIZE) {
        pendingRowsRemain = true;
        break;
      }

      var sheetRow = rowIndex + 1;

      try {
        analyzeInventoryRow_(
          sheet,
          sheetRow,
          row,
          headerMap
        );
      } catch (error) {
        console.error(
          "Analysis failed for row " +
          sheetRow +
          ": " +
          error.message
        );

        sheet
          .getRange(
            sheetRow,
            headerMap["Status"] + 1
          )
          .setValue(
            "Error: " +
            truncateText_(error.message, 250)
          );
      }

      processedCount++;
    }

    SpreadsheetApp.flush();

    console.log(
      "Images processed in this batch: " +
      processedCount
    );

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
  } finally {
    lock.releaseLock();
  }
}


/**
 * Analyzes one inventory row.
 */
function analyzeInventoryRow_(
  sheet,
  sheetRow,
  row,
  headerMap
) {
  var filename = String(
    row[headerMap["Filename"]]
  ).trim();

  var assetType = String(
    row[headerMap["Asset Type"]]
  ).trim();

  var mimeType = String(
    row[headerMap["MIME Type"]]
  ).trim();

  var fileId = String(
    row[headerMap["Drive File ID"]]
  ).trim();

  if (!fileId) {
    throw new Error(
      "The Drive File ID is missing."
    );
  }

  if (
    assetType !== "Image" ||
    mimeType.indexOf("image/") !== 0
  ) {
    sheet
      .getRange(
        sheetRow,
        headerMap["Status"] + 1
      )
      .setValue("Inventory Only - Not an Image");

    console.log(
      "Skipped non-image file: " + filename
    );

    return;
  }

  sheet
    .getRange(
      sheetRow,
      headerMap["Status"] + 1
    )
    .setValue("Analyzing");

  SpreadsheetApp.flush();

  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();

  var analysis = callOpenAIImageAnalysis_(
    blob,
    filename
  );

  writeAnalysisResults_(
    sheet,
    sheetRow,
    headerMap,
    analysis
  );

  sheet
    .getRange(
      sheetRow,
      headerMap["Status"] + 1
    )
    .setValue("Complete");

  console.log(
    "Completed analysis: " + filename
  );
}


/**
 * Sends an image to the OpenAI Responses API.
 */
function callOpenAIImageAnalysis_(
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
    "",
    "Return only valid JSON. Do not use Markdown or code fences.",
    "",
    "Use this exact structure:",
    "{",
    '  "recipe": "",',
    '  "meal": "",',
    '  "ingredients": "",',
    '  "uses_nicer_slicer": "",',
    '  "quality": "",',
    '  "orientation": "",',
    '  "description": "",',
    '  "confidence": ""',
    "}",
    "",
    "Rules:",
    '- recipe: Identify the likely dish. Use "Unknown" when uncertain.',
    '- meal: Choose Breakfast, Lunch, Dinner, Snack, Appetizer, Dessert, Product, or Unknown.',
    "- ingredients: List only clearly visible or strongly suggested principal ingredients, separated by commas.",
    '- uses_nicer_slicer: Choose Yes, No, Possible, or Unknown.',
    '- quality: Choose Excellent, Good, Fair, Poor, or Unusable.',
    '- orientation: Choose Landscape, Portrait, or Square.',
    "- description: Give a concise factual description of what is visible.",
    '- confidence: Choose High, Medium, or Low.',
    "",
    "Do not invent a precise recipe when the image does not provide enough evidence.",
    "Filename: " + filename
  ].join("\n");

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
    max_output_tokens: 1200
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
      "The OpenAI response was not valid JSON."
    );
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

  return parseAnalysisJson_(outputText);
}


/**
 * Extracts text from a Responses API response.
 */
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
      var contentItem =
        content[contentIndex];

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


/**
 * Converts the model response into an analysis object.
 */
function parseAnalysisJson_(text) {
  var cleanedText = String(text).trim();

  cleanedText = cleanedText
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/i, "")
    .replace(/\s*```$/i, "")
    .trim();

  var firstBrace =
    cleanedText.indexOf("{");

  var lastBrace =
    cleanedText.lastIndexOf("}");

  if (
    firstBrace !== -1 &&
    lastBrace !== -1
  ) {
    cleanedText =
      cleanedText.substring(
        firstBrace,
        lastBrace + 1
      );
  }

  var parsed;

  try {
    parsed = JSON.parse(cleanedText);
  } catch (error) {
    throw new Error(
      "Could not parse the image analysis JSON: " +
      truncateText_(cleanedText, 300)
    );
  }

  return {
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
      ""
    ),

    uses_nicer_slicer: safeText_(
      parsed.uses_nicer_slicer,
      "Unknown"
    ),

    quality: safeText_(
      parsed.quality,
      "Unknown"
    ),

    orientation: safeText_(
      parsed.orientation,
      "Unknown"
    ),

    description: safeText_(
      parsed.description,
      ""
    ),

    confidence: safeText_(
      parsed.confidence,
      "Low"
    )
  };
}


/**
 * Writes analysis results into the matching spreadsheet row.
 */
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
    "Description",
    analysis.description
  );

  setCellByHeader_(
    sheet,
    sheetRow,
    headerMap,
    "Confidence",
    analysis.confidence
  );
}


/**
 * Adds analysis columns if they do not already exist.
 */
function ensureAnalysisHeaders_(sheet) {
  var analysisHeaders = [
    "Recipe",
    "Meal",
    "Ingredients",
    "Uses Nicer Slicer",
    "Quality",
    "Orientation",
    "Description",
    "Confidence"
  ];

  var lastColumn =
    sheet.getLastColumn();

  if (lastColumn === 0) {
    throw new Error(
      "The Image Inventory sheet has no headers."
    );
  }

  var existingHeaders = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0];

  var existingHeaderNames = {};

  for (
    var i = 0;
    i < existingHeaders.length;
    i++
  ) {
    var headerName = String(
      existingHeaders[i]
    ).trim();

    if (headerName) {
      existingHeaderNames[headerName] =
        true;
    }
  }

  var missingHeaders = [];

  for (
    var j = 0;
    j < analysisHeaders.length;
    j++
  ) {
    if (
      !existingHeaderNames[
        analysisHeaders[j]
      ]
    ) {
      missingHeaders.push(
        analysisHeaders[j]
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

    console.log(
      "Added analysis headers: " +
      missingHeaders.join(", ")
    );
  }
}


/**
 * Returns the Image Inventory sheet.
 */
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


/**
 * Creates a map of header names to zero-based array positions.
 */
function createHeaderMap_(headers) {
  var headerMap = {};

  for (
    var i = 0;
    i < headers.length;
    i++
  ) {
    var headerName = String(
      headers[i]
    ).trim();

    if (headerName) {
      headerMap[headerName] = i;
    }
  }

  return headerMap;
}


/**
 * Confirms that the inventory contains the columns required for analysis.
 */
function validateRequiredHeaders_(
  headerMap
) {
  var requiredHeaders = [
    "Filename",
    "Asset Type",
    "MIME Type",
    "Drive File ID",
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
 * Checks whether any Pending Analysis rows remain.
 */
function hasPendingAnalysisRows_(sheet) {
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return false;
  }

  var headers = sheet
    .getRange(
      1,
      1,
      1,
      sheet.getLastColumn()
    )
    .getValues()[0];

  var headerMap =
    createHeaderMap_(headers);

  if (
    headerMap["Status"] === undefined
  ) {
    return false;
  }

  var statuses = sheet
    .getRange(
      2,
      headerMap["Status"] + 1,
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


/**
 * Writes one value using a spreadsheet header name.
 */
function setCellByHeader_(
  sheet,
  sheetRow,
  headerMap,
  headerName,
  value
) {
  if (
    headerMap[headerName] === undefined
  ) {
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


function safeText_(value, fallback) {
  if (
    value === null ||
    value === undefined
  ) {
    return fallback;
  }

  var text = String(value).trim();

  return text || fallback;
}


function truncateText_(text, maximumLength) {
  var value = String(text || "");

  if (value.length <= maximumLength) {
    return value;
  }

  return value.substring(
    0,
    maximumLength
  ) + "...";
}