/**
 * SB Schirmer Asset Management System
 * Shared configuration
 *
 * Do not store API keys or passwords in this file.
 * The OpenAI API key remains in Apps Script Properties.
 */

var SB_CONFIG = Object.freeze({
  ROOT_FOLDER_ID:
    "1ETQMtYzxlxzuIdLNi1DR_P5-QEwhH1NK",

  SPREADSHEET_ID:
    "1KCxbVMjtumF9SwA88nM0SEtZjeseptVsA0NfTVt2lTM",

  INVENTORY_SHEET_NAME:
    "Image Inventory",

  OPENAI_MODEL:
    "gpt-5",

  ANALYSIS_BATCH_SIZE:
    5,

  MAX_FILE_SIZE_MB:
    10,

  MAX_IMAGES_PER_SESSION:
    25,

  MAX_RETRY_ATTEMPTS:
    3,

  ANALYSIS_ENABLED:
    true,

  PROMPT_VERSION:
    "1.0"
});