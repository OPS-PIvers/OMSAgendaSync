/**
 * @fileoverview Main script file for extracting agenda text from Google Slides.
 * This file contains the core logic for processing presentations and interacting with the Google Sheet.
 * It relies on a separate 'Constants.gs' file for configuration.
 */

// --- TRIGGER SUPPORT FUNCTIONS ---
/**
 * Wrapper function designed to be called by a timed trigger.
 * It calls the main extraction function without passing any arguments,
 * which prevents the trigger's event object from causing errors.
 */
function runDailyExtractionTrigger() {
  Logger.log('Hourly extraction trigger initiated.');
  extractTextForCurrentDayAgenda();
}

/**
 * Wrapper function designed to be called by a daily archive trigger.
 * This should run once per day (e.g., at 11:30 PM) to archive the final
 * version of the day's agenda data before starting fresh the next day.
 */
function runDailyArchiveTrigger() {
  Logger.log('Daily archive trigger initiated.');
  archiveCurrentDayDataOnly();
}

/**
 * Checks if the script is running in a context where it can access the user interface.
 * @returns {boolean} True if the UI is available, false otherwise.
 */
function isUiAvailable() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * Gets the date for the Monday of the current week.
 * @returns {Date} A Date object set to the preceding Monday.
 */
function getMondayOfCurrentWeek() {
  const today = new Date();
  const day = today.getDay(); // Sunday = 0, Monday = 1, ..., Saturday = 6
  const diff = today.getDate() - day + (day === 0 ? -6 : 1); // Adjust for Sunday
  return new Date(today.setDate(diff));
}

/**
 * Extracts text from a shape while preserving all individual hyperlinks.
 * Returns either plain text or multiple HYPERLINK formulas joined together.
 * @param {GoogleAppsScript.Slides.TextRange} textRange The TextRange from a shape.
 * @returns {string} The text content with preserved hyperlinks as HYPERLINK formulas.
 */
function extractTextWithAllLinks(textRange) {
  const fullText = textRange.asString().trim();
  if (fullText === '') {
    return 'N/A'; // Return a default value for empty text boxes
  }

  const runs = textRange.getRuns();
  const textParts = [];
  
  for (const run of runs) {
    const runText = run.asString();
    if (runText.trim() === '') continue; // Skip empty runs
    
    const link = run.getTextStyle().getLink();
    if (link && link.getUrl()) {
      // This run has a hyperlink - create a HYPERLINK formula for it
      const url = link.getUrl();
      const linkText = runText.trim();
      textParts.push(`=HYPERLINK("${url}", "${linkText.replace(/"/g, '""')}")`);
    } else {
      // This run has no hyperlink - add as plain text
      const plainText = runText.trim();
      if (plainText) {
        textParts.push(plainText);
      }
    }
  }
  
  if (textParts.length === 0) {
    return fullText; // Fallback to original text if no parts were processed
  }
  
  // Join the parts with newlines for better readability in the spreadsheet
  return textParts.join('\n');
}

/**
 * Legacy function kept for backward compatibility - now calls the new multi-link function
 * @deprecated Use extractTextWithAllLinks instead
 * @param {GoogleAppsScript.Slides.TextRange} textRange The TextRange from a shape.
 * @returns {string} The text content with preserved hyperlinks.
 */
function extractTextAndFirstLink(textRange) {
  return extractTextWithAllLinks(textRange);
}


/**
 * Extracts text from specific text boxes on Google Slide presentations
 * for the current day of the week, after finding the correct slide for the current week.
 * Appends the extracted data, preserving rich text and hyperlinks, to a main data sheet.
 * @param {string} [dayToTest] - Optional. A string representing the day of the week
 * (e.g., "Monday") to run the script for, used for testing purposes. If undefined,
 * the script will use the actual current day.
 */
function extractTextForCurrentDayAgenda(dayToTest) {
  const SPREADSHEET_ID = CONSTANTS.SPREADSHEET_ID;
  const CONFIG_SHEET_NAME = CONSTANTS.CONFIG_SHEET_NAME;
  const DATA_SHEET_NAME = CONSTANTS.DATA_SHEET_NAME;
  const BOX_COORDINATES = CONSTANTS.BOX_COORDINATES;

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

  const configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    const errorMessage = `Error: Configuration sheet '${CONFIG_SHEET_NAME}' not found. Please ensure it exists.`;
    Logger.log(errorMessage);
    if (isUiAvailable()) {
      SpreadsheetApp.getUi().alert(errorMessage);
    }
    return;
  }

  let dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
  if (!dataSheet) {
    const errorMessage = `Error: Data sheet '${DATA_SHEET_NAME}' not found. Please ensure it exists.`;
    Logger.log(errorMessage);
    if (isUiAvailable()) {
      SpreadsheetApp.getUi().alert(errorMessage);
    }
    return;
  }

  const today = new Date();
  
  // Note: Archiving is now handled by a separate daily trigger (runDailyArchiveTrigger)
  // This allows teachers to update their agendas throughout the day with only the 
  // final version being archived at end of day
  Logger.log('Extraction running - archiving handled separately by daily trigger');
  
  dataSheet.clearContents();
  dataSheet.appendRow([
    'Teacher Last Name', 'Class Name', 'Day of Week', 'Turn In', 'Activities',
    'Practice Work', 'Upcoming', 'Grade Level'
  ]);
  const dayOfWeek = dayToTest || Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE');
  Logger.log(`Running extraction for: ${dayOfWeek}`);

  if (!BOX_COORDINATES.hasOwnProperty(dayOfWeek)) {
    const message = dayToTest ?
      `The provided test day '${dayToTest}' has no coordinates defined.` :
      `Today is ${dayOfWeek}. No agenda extraction scheduled for this day.`;
    Logger.log(message);
    if (isUiAvailable()) {
      SpreadsheetApp.getUi().alert(message);
    }
    return;
  }

  const monday = getMondayOfCurrentWeek();
  const formattedMonday = Utilities.formatDate(monday, Session.getScriptTimeZone(), 'M/d/yyyy');
  const weekOfText = `WEEK OF ${formattedMonday}`.toUpperCase();
  const semanaDeText = `SEMANA DE ${formattedMonday}`.toUpperCase();
  Logger.log(`Searching for slides with English text: "${weekOfText}"`);
  Logger.log(`Searching for slides with Spanish text: "${semanaDeText}"`);

  const currentDayBoxes = BOX_COORDINATES[dayOfWeek];
  const upcomingBox = BOX_COORDINATES['Upcoming'];

  const configDataRange = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 4);
  const configValues = configDataRange.getValues();

  if (configValues.length === 0 || configValues[0].every(cell => !cell)) {
    const errorMessage = 'No presentation IDs found in the configuration sheet.';
    Logger.log(errorMessage);
    if (isUiAvailable()) {
      SpreadsheetApp.getUi().alert(errorMessage);
    }
    return;
  }

  Logger.log(`Found ${configValues.length} presentation entries to process.`);

  configValues.forEach(row => {
    const [presentationId, teacherLastName, className, gradeLevel] = row.map(String);

    if (!presentationId.trim()) {
      Logger.log('Skipping empty presentation ID row.');
      return;
    }

    try {
      const presentation = SlidesApp.openById(presentationId.trim());
      const slides = presentation.getSlides();

      if (slides.length === 0) {
        throw new Error("Presentation has no slides.");
      }

      let agendaSlide = null;
      let matchedPattern = '';
      for (const slide of slides) {
        const shapes = slide.getShapes();
        for (const shape of shapes) {
          const shapeText = shape.getText().asString().toUpperCase();
          if (shapeText.includes(weekOfText)) {
            agendaSlide = slide;
            matchedPattern = 'English (WEEK OF)';
            break;
          } else if (shapeText.includes(semanaDeText)) {
            agendaSlide = slide;
            matchedPattern = 'Spanish (SEMANA DE)';
            break;
          }
        }
        if (agendaSlide) {
          break;
        }
      }

      if (!agendaSlide) {
        throw new Error(`Slide not found with either "${weekOfText}" or "${semanaDeText}".`);
      }
      
      Logger.log(`Found agenda slide using ${matchedPattern} pattern for ${teacherLastName.trim()} - ${className.trim()}`);

      const pageElements = agendaSlide.getPageElements();
      let topBoxText = 'N/A', midBoxText = 'N/A', botBoxText = 'N/A', upcomingText = 'N/A';
      const tolerance = CONSTANTS.TOLERANCE;

      const matchesBox = (shape, targetBox, boxType = '') => {
        const xDiff = Math.abs(shape.getLeft() - targetBox.x);
        const yDiff = Math.abs(shape.getTop() - targetBox.y);
        const wDiff = Math.abs(shape.getWidth() - targetBox.width);
        const hDiff = Math.abs(shape.getHeight() - targetBox.height);
        const matches = xDiff < tolerance && yDiff < tolerance && wDiff < tolerance && hDiff < tolerance;
        
        // Debug logging for Tuesday practice work specifically
        if (dayOfWeek === 'Tuesday' && boxType === 'bottom') {
          Logger.log(`=== TUESDAY PRACTICE WORK DEBUG ===`);
          Logger.log(`Shape: (${shape.getLeft()}, ${shape.getTop()}) ${shape.getWidth()}x${shape.getHeight()}`);
          Logger.log(`Target: (${targetBox.x}, ${targetBox.y}) ${targetBox.width}x${targetBox.height}`);
          Logger.log(`Differences: X=${xDiff}, Y=${yDiff}, W=${wDiff}, H=${hDiff} (tolerance=${tolerance})`);
          Logger.log(`Matches: ${matches}`);
        }
        
        return matches;
      };

      pageElements.forEach(element => {
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          const shape = element.asShape();
          const textRange = shape.getText();
          if (textRange.isEmpty()) return;

          const cellValue = extractTextWithAllLinks(textRange);
          const shapeText = textRange.asString().trim();

          // Debug logging for Tuesday shapes with content
          if (dayOfWeek === 'Tuesday' && shapeText !== '' && shapeText !== '...') {
            Logger.log(`=== TUESDAY SHAPE WITH CONTENT ===`);
            Logger.log(`Text: "${shapeText}"`);
            Logger.log(`Position: (${shape.getLeft()}, ${shape.getTop()})`);
            Logger.log(`Size: ${shape.getWidth()}x${shape.getHeight()}`);
          }

          if (matchesBox(shape, currentDayBoxes.top, 'top')) topBoxText = cellValue;
          else if (matchesBox(shape, currentDayBoxes.middle, 'middle')) midBoxText = cellValue;
          else if (matchesBox(shape, currentDayBoxes.bottom, 'bottom')) botBoxText = cellValue;
          else if (matchesBox(shape, upcomingBox, 'upcoming')) upcomingText = cellValue;
        }
      });

      // Debug logging for Tuesday final results
      if (dayOfWeek === 'Tuesday') {
        Logger.log(`=== TUESDAY FINAL RESULTS ===`);
        Logger.log(`Turn In (top): "${topBoxText}"`);
        Logger.log(`Activities (middle): "${midBoxText}"`);
        Logger.log(`Practice Work (bottom): "${botBoxText}"`);
        Logger.log(`Upcoming: "${upcomingText}"`);
      }

      dataSheet.appendRow([
        teacherLastName.trim(), className.trim(), dayOfWeek, topBoxText, midBoxText,
        botBoxText, upcomingText, gradeLevel.trim()
      ]);
      Logger.log(`Processed: ${teacherLastName.trim()} - ${className.trim()} for ${dayOfWeek}`);

    } catch (e) {
      Logger.log(`Error processing presentation ID ${presentationId.trim()} (${teacherLastName.trim()}, ${className.trim()}): ${e.message}`);
      dataSheet.appendRow([
        teacherLastName.trim(), className.trim(), dayOfWeek, 'ERROR', 'ERROR', 'ERROR', 'ERROR',
        gradeLevel.trim(), `Error: ${e.message}`
      ]);
    }
  });

  const completionMessage = 'Data for ' + dayOfWeek + ' has been extracted and compiled into the "' + DATA_SHEET_NAME + '" tab.';
  Logger.log('Hourly extraction complete - data updated in Current_Day_Agendas tab (archiving handled by separate daily trigger)');
  Logger.log(completionMessage);
  
  if (isUiAvailable()) {
    SpreadsheetApp.getUi().alert(
      'Daily Agenda Extraction Complete!',
      completionMessage,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// --- NEW TESTING FUNCTIONS ---
function testForMonday() { extractTextForCurrentDayAgenda('Monday'); }
function testForTuesday() { extractTextForCurrentDayAgenda('Tuesday'); }
function testForWednesday() { extractTextForCurrentDayAgenda('Wednesday'); }
function testForThursday() { extractTextForCurrentDayAgenda('Thursday'); }
function testForFriday() { extractTextForCurrentDayAgenda('Friday'); }

/**
 * Test function to verify the archive-only functionality works correctly.
 * This can be run manually to test the archiving without waiting for the trigger.
 */
function testArchiveOnly() {
  Logger.log('=== Testing Archive Only Function ===');
  archiveCurrentDayDataOnly();
  Logger.log('=== Archive Test Complete ===');
}

/**
 * Test function to verify archive date retrieval specifically for September 1, 2025.
 * This helps troubleshoot the user's specific issue with date detection.
 */
function testArchiveDateRetrieval() {
  Logger.log('=== Testing Archive Date Retrieval for 2025-09-01 ===');
  
  // First, debug the archive sheet formats
  debugArchiveDateFormats('Archive_2025_09');
  
  // Then try to retrieve data for the specific date
  Logger.log('\n=== Attempting to retrieve data for 2025-09-01 ===');
  const result = getArchivedDataForDate('2025-09-01');
  Logger.log(`Result: ${JSON.stringify(result)}`);
  Logger.log(`Found ${result.payload ? result.payload.length : 0} records`);
  
  Logger.log('=== Archive Date Retrieval Test Complete ===');
}

/**
 * Helper function to set up the required triggers programmatically.
 * Run this once to set up both the hourly extraction and daily archive triggers.
 * 
 * IMPORTANT: You need to set up two triggers:
 * 1. Hourly trigger for runDailyExtractionTrigger() - runs every hour during business hours
 * 2. Daily trigger for runDailyArchiveTrigger() - runs once daily at 11:30 PM
 * 
 * These should be set up in the Google Apps Script trigger interface, not through code,
 * for better reliability and management.
 */
function setupTriggersInstructions() {
  const message = `
  To complete the setup, you need to configure two triggers in Google Apps Script:
  
  1. HOURLY EXTRACTION TRIGGER:
     - Function: runDailyExtractionTrigger
     - Event: Time-driven, Hour timer, Every hour
     - Time: During business hours (e.g., 6 AM to 6 PM)
     - This extracts fresh data from slides every hour
  
  2. DAILY ARCHIVE TRIGGER:  
     - Function: runDailyArchiveTrigger
     - Event: Time-driven, Day timer, Every day
     - Time: 11:30 PM (or desired end-of-day time)
     - This archives the final version once per day
  
  This setup ensures:
  - Teachers can update agendas throughout the day
  - Only final versions get archived (no duplicates)
  - Current day view stays updated hourly
  `;
  
  Logger.log(message);
  if (isUiAvailable()) {
    SpreadsheetApp.getUi().alert('Trigger Setup Instructions', message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
// --- END NEW TESTING FUNCTIONS ---


/**
 * Gets or creates an archive sheet for the specified date.
 * Archive sheets are organized by month (e.g., "Archive_2024_01").
 * @param {Date} date The date for which to get/create the archive sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The archive sheet
 */
function getOrCreateArchiveSheet(date) {
  const SPREADSHEET_ID = CONSTANTS.SPREADSHEET_ID;
  const ARCHIVE_SHEET_PREFIX = CONSTANTS.ARCHIVE_SHEET_PREFIX;
  
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const archiveSheetName = `${ARCHIVE_SHEET_PREFIX}${year}_${month}`;
  
  let archiveSheet = spreadsheet.getSheetByName(archiveSheetName);
  
  if (!archiveSheet) {
    archiveSheet = spreadsheet.insertSheet(archiveSheetName);
    
    archiveSheet.appendRow([
      'Date', 'Teacher Last Name', 'Class Name', 'Day of Week', 'Turn In', 
      'Activities', 'Practice Work', 'Upcoming', 'Grade Level'
    ]);
    
    const headerRange = archiveSheet.getRange(1, 1, 1, 9);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f0f0f0');
    
    Logger.log(`Created new archive sheet: ${archiveSheetName}`);
  }
  
  return archiveSheet;
}

/**
 * Archives current day data for end-of-day archiving only.
 * This function includes safety checks to prevent duplicate archiving and
 * should be called by the daily archive trigger.
 */
function archiveCurrentDayDataOnly() {
  const SPREADSHEET_ID = CONSTANTS.SPREADSHEET_ID;
  const DATA_SHEET_NAME = CONSTANTS.DATA_SHEET_NAME;
  
  try {
    const today = new Date();
    const dateString = normalizeDateToString(today);
    
    Logger.log(`Starting daily archive process for ${dateString}`);
    
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
    
    if (!dataSheet) {
      Logger.log(`Warning: Data sheet '${DATA_SHEET_NAME}' not found for archiving.`);
      return;
    }
    
    const dataRange = dataSheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length <= 1) {
      Logger.log('No data to archive (only headers present).');
      return;
    }
    
    // Check if we've already archived data for today to prevent duplicates
    const archiveSheet = getOrCreateArchiveSheet(today);
    const archiveRange = archiveSheet.getDataRange();
    const archiveValues = archiveRange.getValues();
    
    // Check if today's data already exists in archive
    let alreadyArchived = false;
    for (let i = 1; i < archiveValues.length; i++) {
      if (archiveValues[i][0] === dateString) {
        alreadyArchived = true;
        break;
      }
    }
    
    if (alreadyArchived) {
      Logger.log(`Data for ${dateString} has already been archived today. Skipping to prevent duplicates.`);
      return;
    }
    
    // Proceed with archiving
    const formulas = dataRange.getFormulas();
    
    for (let i = 1; i < values.length; i++) {
      const rowValues = values[i];
      const rowFormulas = formulas[i];
      
      const archiveRow = [dateString];
      
      for (let j = 0; j < rowValues.length; j++) {
        if (rowFormulas[j]) {
          archiveRow.push(rowFormulas[j]);
        } else {
          archiveRow.push(rowValues[j]);
        }
      }
      
      archiveSheet.appendRow(archiveRow);
    }
    
    Logger.log(`Successfully archived ${values.length - 1} rows for ${dateString}`);
    
  } catch (e) {
    Logger.log(`Error in daily archive process: ${e.message}`);
  }
}

/**
 * Archives current day data by moving it to the appropriate monthly archive sheet.
 * This function should be called before clearing the current day sheet.
 * @param {Date} date The date of the data being archived
 * @deprecated This function is kept for backward compatibility but should not be used for regular operations
 */
function archiveCurrentDayData(date) {
  const SPREADSHEET_ID = CONSTANTS.SPREADSHEET_ID;
  const DATA_SHEET_NAME = CONSTANTS.DATA_SHEET_NAME;
  
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
    
    if (!dataSheet) {
      Logger.log(`Warning: Data sheet '${DATA_SHEET_NAME}' not found for archiving.`);
      return;
    }
    
    const dataRange = dataSheet.getDataRange();
    const values = dataRange.getValues();
    const formulas = dataRange.getFormulas();
    
    if (values.length <= 1) {
      Logger.log('No data to archive (only headers present).');
      return;
    }
    
    const archiveSheet = getOrCreateArchiveSheet(date);
    const dateString = normalizeDateToString(date);
    
    for (let i = 1; i < values.length; i++) {
      const rowValues = values[i];
      const rowFormulas = formulas[i];
      
      const archiveRow = [dateString];
      
      for (let j = 0; j < rowValues.length; j++) {
        if (rowFormulas[j]) {
          archiveRow.push(rowFormulas[j]);
        } else {
          archiveRow.push(rowValues[j]);
        }
      }
      
      archiveSheet.appendRow(archiveRow);
    }
    
    Logger.log(`Archived ${values.length - 1} rows to archive sheet for ${dateString}`);
    
  } catch (e) {
    Logger.log(`Error archiving data: ${e.message}`);
  }
}

/**
 * Normalizes date values to YYYY-MM-DD string format for consistent comparison.
 * @param {Date|string|number} dateValue The date value to normalize
 * @returns {string|null} The date in YYYY-MM-DD format, or null if invalid
 */
function normalizeDateToString(dateValue) {
  try {
    let date;
    
    if (dateValue instanceof Date) {
      date = dateValue;
    } else if (typeof dateValue === 'string') {
      // Handle various string formats
      if (dateValue.includes('/')) {
        // Handle formats like "9/1/2025" or "09/01/2025"
        date = new Date(dateValue);
      } else if (dateValue.match(/^\d{4}-\d{2}-\d{2}$/)) {
        // Already in YYYY-MM-DD format
        return dateValue;
      } else {
        date = new Date(dateValue);
      }
    } else if (typeof dateValue === 'number') {
      // Excel serial date number
      date = new Date(dateValue);
    } else {
      return null;
    }
    
    if (isNaN(date.getTime())) {
      return null;
    }
    
    // Format as YYYY-MM-DD
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
    
  } catch (e) {
    Logger.log(`Error normalizing date value: ${dateValue}, Error: ${e.message}`);
    return null;
  }
}

/**
 * Retrieves archived agenda data for a specific date.
 * @param {string} dateString The date in 'YYYY-MM-DD' format
 * @returns {Object} An object containing the archived data or error
 */
function getArchivedDataForDate(dateString) {
  const SPREADSHEET_ID = CONSTANTS.SPREADSHEET_ID;
  const ARCHIVE_SHEET_PREFIX = CONSTANTS.ARCHIVE_SHEET_PREFIX;
  
  try {
    // Extract year and month directly from YYYY-MM-DD string to avoid timezone issues
    const dateParts = dateString.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!dateParts) {
      throw new Error(`Invalid date format. Expected YYYY-MM-DD, got: ${dateString}`);
    }
    
    const year = dateParts[1];
    const month = dateParts[2];
    const archiveSheetName = `${ARCHIVE_SHEET_PREFIX}${year}_${month}`;
    
    Logger.log(`Parsing date: ${dateString} -> Year: ${year}, Month: ${month}`);
    Logger.log(`Looking for archive sheet: ${archiveSheetName}`);
    
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const archiveSheet = spreadsheet.getSheetByName(archiveSheetName);
    if (!archiveSheet) {
      Logger.log(`Archive sheet "${archiveSheetName}" not found. Available sheets: ${spreadsheet.getSheets().map(s => s.getName()).join(', ')}`);
      return { payload: [] };
    }
    
    Logger.log(`Successfully found archive sheet: ${archiveSheetName}`);
    
    const range = archiveSheet.getDataRange();
    const values = range.getValues();
    const formulas = range.getFormulas();
    
    if (values.length <= 1) {
      return { payload: [] };
    }
    
    const headers = values[0];
    const data = [];
    const targetDate = normalizeDateToString(dateString);
    
    Logger.log(`Searching for archived data for date: ${dateString} (normalized: ${targetDate})`);
    Logger.log(`Archive sheet ${archiveSheetName} has ${values.length - 1} data rows`);
    
    for (let i = 1; i < values.length; i++) {
      const currentRowValues = values[i];
      const currentRowFormulas = formulas[i];
      const rowDate = normalizeDateToString(currentRowValues[0]);
      
      // Debug logging for first few rows
      if (i <= 3) {
        Logger.log(`Row ${i}: Original date value: ${currentRowValues[0]}, Type: ${typeof currentRowValues[0]}, Normalized: ${rowDate}`);
      }
      
      if (rowDate === targetDate) {
        Logger.log(`Found matching date at row ${i}: ${currentRowValues[0]} -> ${rowDate}`);
        const obj = {};
        
        for (let j = 1; j < headers.length; j++) {
          const cleanedHeader = headers[j].replace(/[^a-zA-Z0-9]/g, '');
          if (currentRowFormulas[j]) {
            obj[cleanedHeader] = currentRowFormulas[j];
          } else {
            obj[cleanedHeader] = currentRowValues[j];
          }
        }
        data.push(obj);
      }
    }
    
    return { payload: data };
    
  } catch (e) {
    Logger.log(`Error retrieving archived data for ${dateString}: ${e.message}`);
    return { error: `Failed to fetch archived data: ${e.message}` };
  }
}

/**
 * Debug function to inspect the format and content of dates in archive sheets.
 * This helps troubleshoot date matching issues.
 * @param {string} archiveSheetName Optional specific archive sheet name to inspect
 */
function debugArchiveDateFormats(archiveSheetName) {
  const SPREADSHEET_ID = CONSTANTS.SPREADSHEET_ID;
  const ARCHIVE_SHEET_PREFIX = CONSTANTS.ARCHIVE_SHEET_PREFIX;
  
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = archiveSheetName ? 
      [spreadsheet.getSheetByName(archiveSheetName)] : 
      spreadsheet.getSheets().filter(sheet => sheet.getName().startsWith(ARCHIVE_SHEET_PREFIX));
    
    Logger.log(`=== DEBUGGING ARCHIVE DATE FORMATS ===`);
    
    sheets.forEach(sheet => {
      if (!sheet) return;
      
      Logger.log(`\nSheet: ${sheet.getName()}`);
      const range = sheet.getDataRange();
      const values = range.getValues();
      
      if (values.length <= 1) {
        Logger.log(`  No data rows found`);
        return;
      }
      
      Logger.log(`  Headers: ${values[0].join(' | ')}`);
      Logger.log(`  Total rows: ${values.length - 1}`);
      
      // Inspect first 5 date values
      for (let i = 1; i < Math.min(6, values.length); i++) {
        const dateValue = values[i][0];
        const normalized = normalizeDateToString(dateValue);
        Logger.log(`  Row ${i}: "${dateValue}" (${typeof dateValue}) -> normalized: "${normalized}"`);
        
        if (dateValue instanceof Date) {
          Logger.log(`    Date details: ${dateValue.toISOString()}, Local: ${dateValue.toLocaleDateString()}`);
        }
      }
      
      // Show unique date values in column A
      const uniqueDates = new Set();
      for (let i = 1; i < values.length; i++) {
        const normalized = normalizeDateToString(values[i][0]);
        if (normalized) uniqueDates.add(normalized);
      }
      Logger.log(`  Unique normalized dates: ${Array.from(uniqueDates).sort().join(', ')}`);
    });
    
  } catch (e) {
    Logger.log(`Error in debugArchiveDateFormats: ${e.message}`);
  }
}

/**
 * Gets a list of all available archive dates.
 * @returns {Array<string>} Array of date strings in 'YYYY-MM-DD' format
 */
function getAvailableArchiveDates() {
  const SPREADSHEET_ID = CONSTANTS.SPREADSHEET_ID;
  const ARCHIVE_SHEET_PREFIX = CONSTANTS.ARCHIVE_SHEET_PREFIX;
  
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = spreadsheet.getSheets();
    const dates = new Set();
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith(ARCHIVE_SHEET_PREFIX)) {
        const range = sheet.getDataRange();
        const values = range.getValues();
        
        for (let i = 1; i < values.length; i++) {
          const normalizedDate = normalizeDateToString(values[i][0]);
          if (normalizedDate) {
            dates.add(normalizedDate);
          }
        }
      }
    });
    
    return Array.from(dates).sort();
    
  } catch (e) {
    Logger.log(`Error getting available archive dates: ${e.message}`);
    return [];
  }
}

/**
 * Creates copies of the master presentation for selected staff members.
 * Reads selected rows from the Staff Directory sheet and creates personalized copies.
 */
function createCopiesForSelectedRows() {
  const SPREADSHEET_ID = CONSTANTS.SPREADSHEET_ID;
  const STAFF_DIRECTORY_SHEET_NAME = CONSTANTS.STAFF_DIRECTORY_SHEET_NAME;
  const MASTER_PRESENTATION_ID = CONSTANTS.MASTER_PRESENTATION_ID;
  const COLUMNS = CONSTANTS.STAFF_DIRECTORY_COLUMNS;

  Logger.log('Starting createCopiesForSelectedRows function');

  if (MASTER_PRESENTATION_ID === 'REPLACE_WITH_MASTER_PRESENTATION_ID') {
    SpreadsheetApp.getUi().alert('Error: Master Presentation ID not configured. Please update MASTER_PRESENTATION_ID in Constants.js');
    return;
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const staffSheet = spreadsheet.getSheetByName(STAFF_DIRECTORY_SHEET_NAME);
  
  if (!staffSheet) {
    SpreadsheetApp.getUi().alert(`Error: Staff Directory sheet '${STAFF_DIRECTORY_SHEET_NAME}' not found.`);
    return;
  }

  // Better selection handling - try multiple methods
  let selectedRange = null;
  try {
    // First try to get selection from the active sheet
    const activeSheet = SpreadsheetApp.getActiveSheet();
    if (activeSheet && activeSheet.getName() === STAFF_DIRECTORY_SHEET_NAME) {
      selectedRange = activeSheet.getActiveRange();
      Logger.log(`Got selection from active sheet: ${selectedRange.getA1Notation()}`);
    } else {
      // Fall back to getting selection from the staff sheet
      const selection = staffSheet.getSelection();
      if (selection && selection.getActiveRange()) {
        selectedRange = selection.getActiveRange();
        Logger.log(`Got selection from staff sheet: ${selectedRange.getA1Notation()}`);
      }
    }
  } catch (e) {
    Logger.log(`Error getting selection: ${e.message}`);
  }
  
  if (!selectedRange) {
    SpreadsheetApp.getUi().alert('Please select one or more rows in the Staff Directory sheet to create copies for. Make sure you are on the Staff Directory sheet when running this function.');
    return;
  }

  Logger.log(`Processing selection: ${selectedRange.getA1Notation()}, rows ${selectedRange.getRow()} to ${selectedRange.getLastRow()}`);

  let processedCount = 0;
  let errorCount = 0;
  const errors = [];

  // Process each row in the selected range
  for (let row = selectedRange.getRow(); row <= selectedRange.getLastRow(); row++) {
    if (row === 1) {
      Logger.log(`Skipping header row ${row}`);
      continue; // Skip header row
    }
    
    Logger.log(`Processing row ${row}`);
    const rowData = staffSheet.getRange(row, 1, 1, 5).getValues()[0];
    Logger.log(`Raw row data: ${JSON.stringify(rowData)}`);
    
    // Trim whitespace and validate
    const firstName = rowData[COLUMNS.FIRST_NAME] ? String(rowData[COLUMNS.FIRST_NAME]).trim() : '';
    const lastName = rowData[COLUMNS.LAST_NAME] ? String(rowData[COLUMNS.LAST_NAME]).trim() : '';
    const email = rowData[COLUMNS.EMAIL] ? String(rowData[COLUMNS.EMAIL]).trim() : '';
    const existingUrl = rowData[COLUMNS.AGENDA_URL] ? String(rowData[COLUMNS.AGENDA_URL]).trim() : '';
    const existingId = rowData[COLUMNS.SLIDE_ID] ? String(rowData[COLUMNS.SLIDE_ID]).trim() : '';

    Logger.log(`Processed data - First: "${firstName}", Last: "${lastName}", Email: "${email}"`);

    if (!firstName || !lastName || !email) {
      const errorMsg = `Row ${row}: Missing required information - First: "${firstName}", Last: "${lastName}", Email: "${email}"`;
      Logger.log(errorMsg);
      errors.push(errorMsg);
      errorCount++;
      continue;
    }

    if (existingUrl && existingId) {
      Logger.log(`${firstName} ${lastName} already has a copy, asking user for confirmation`);
      const response = SpreadsheetApp.getUi().alert(
        `${firstName} ${lastName} already has a copy`,
        `${firstName} ${lastName} already has an agenda copy. Do you want to create a new one?`,
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      if (response !== SpreadsheetApp.getUi().Button.YES) {
        Logger.log(`User chose not to recreate copy for ${firstName} ${lastName}`);
        continue;
      }
    }

    try {
      Logger.log(`Creating copy for ${firstName} ${lastName} (${email})`);
      const result = createPersonalizedCopyForTeacher(firstName, lastName, email);
      if (result.success) {
        updateStaffDirectoryRow(staffSheet, row, result.presentationUrl, result.presentationId);
        processedCount++;
        Logger.log(`Successfully created copy for ${firstName} ${lastName}: ${result.presentationUrl}`);
      } else {
        const errorMsg = `${firstName} ${lastName}: ${result.error}`;
        Logger.log(`Failed to create copy: ${errorMsg}`);
        errors.push(errorMsg);
        errorCount++;
      }
    } catch (e) {
      const errorMsg = `${firstName} ${lastName}: ${e.message}`;
      Logger.log(`Exception creating copy: ${errorMsg}`);
      errors.push(errorMsg);
      errorCount++;
    }
  }

  let message = `Operation completed!\nCopies created: ${processedCount}`;
  if (errorCount > 0) {
    message += `\nErrors: ${errorCount}`;
    if (errors.length > 0) {
      message += `\n\nError details:\n${errors.join('\n')}`;
    }
  }
  
  Logger.log(`Final result: ${message}`);
  SpreadsheetApp.getUi().alert('Staff Directory Copy Creation', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Creates a personalized copy of the master presentation for a specific teacher.
 * @param {string} firstName Teacher's first name
 * @param {string} lastName Teacher's last name  
 * @param {string} email Teacher's email address
 * @returns {Object} Result object with success status and details
 */
function createPersonalizedCopyForTeacher(firstName, lastName, email) {
  const MASTER_PRESENTATION_ID = CONSTANTS.MASTER_PRESENTATION_ID;
  
  try {
    // Use DriveApp to copy the file (correct method for Google Apps Script)
    const copyName = `${lastName} - OMS Agenda 25-26`;
    
    Logger.log(`Creating copy using DriveApp: ${copyName}`);
    const masterFile = DriveApp.getFileById(MASTER_PRESENTATION_ID);
    const copiedFile = masterFile.makeCopy(copyName);
    const presentationId = copiedFile.getId();
    const presentationUrl = `https://docs.google.com/presentation/d/${presentationId}/edit`;
    
    Logger.log(`Copy created with ID: ${presentationId}`);
    
    // Open the copied presentation using SlidesApp for text replacement
    try {
      const copiedPresentation = SlidesApp.openById(presentationId);
      
      // Replace [TEACHER NAME] with teacher's last name in all caps
      const replacementCount = copiedPresentation.replaceAllText('[TEACHER NAME]', lastName.toUpperCase());
      Logger.log(`Replaced ${replacementCount} instances of [TEACHER NAME] with ${lastName.toUpperCase()}`);
    } catch (e) {
      Logger.log(`Warning: Could not replace text in presentation: ${e.message}`);
    }
    
    // Share with teacher as editor using the Drive file
    try {
      copiedFile.addEditor(email);
      Logger.log(`Shared presentation with ${email} as editor`);
    } catch (e) {
      Logger.log(`Warning: Could not share with ${email}: ${e.message}`);
    }
    
    return {
      success: true,
      presentationId: presentationId,
      presentationUrl: presentationUrl,
      copyName: copyName
    };
    
  } catch (e) {
    Logger.log(`Error in createPersonalizedCopyForTeacher: ${e.message}`);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Updates a row in the Staff Directory sheet with presentation URL and ID.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Staff Directory sheet
 * @param {number} row The row number to update
 * @param {string} presentationUrl The URL of the created presentation
 * @param {string} presentationId The ID of the created presentation
 */
function updateStaffDirectoryRow(sheet, row, presentationUrl, presentationId) {
  const COLUMNS = CONSTANTS.STAFF_DIRECTORY_COLUMNS;
  
  sheet.getRange(row, COLUMNS.AGENDA_URL + 1).setValue(presentationUrl);
  sheet.getRange(row, COLUMNS.SLIDE_ID + 1).setValue(presentationId);
}

/**
 * Creates a custom menu in the Google Sheet UI for manual script execution.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Agenda Tools')
    .addItem('Run Daily Agenda Extraction Now', 'extractTextForCurrentDayAgenda')
    .addSeparator()
    .addItem('Archive Current Day Data Now', 'archiveCurrentDayDataOnly')
    .addItem('Test Archive Function', 'testArchiveOnly')
    .addItem('Test Archive Date Retrieval', 'testArchiveDateRetrieval')
    .addItem('Debug Archive Date Formats', 'debugArchiveDateFormats')
    .addSeparator()
    .addItem('View Trigger Setup Instructions', 'setupTriggersInstructions');

  const testSubMenu = ui.createMenu('Run Manual Test As...')
    .addItem('Monday', 'testForMonday')
    .addItem('Tuesday', 'testForTuesday')
    .addItem('Wednesday', 'testForWednesday')
    .addItem('Thursday', 'testForThursday')
    .addItem('Friday', 'testForFriday');

  const staffSubMenu = ui.createMenu('Staff Directory')
    .addItem('Create Copies for Selected Staff', 'createCopiesForSelectedRows')
    .addItem('Debug Staff Directory Selection', 'debugStaffDirectorySelection');

  menu.addSeparator()
      .addSubMenu(testSubMenu)
      .addSeparator()
      .addSubMenu(staffSubMenu)
      .addToUi();
}

/**
 * Debug function to test Staff Directory selection and data reading.
 * This helps troubleshoot issues with the copy creation function.
 */
function debugStaffDirectorySelection() {
  const SPREADSHEET_ID = CONSTANTS.SPREADSHEET_ID;
  const STAFF_DIRECTORY_SHEET_NAME = CONSTANTS.STAFF_DIRECTORY_SHEET_NAME;
  const MASTER_PRESENTATION_ID = CONSTANTS.MASTER_PRESENTATION_ID;
  const COLUMNS = CONSTANTS.STAFF_DIRECTORY_COLUMNS;

  Logger.log('=== DEBUG: Starting Staff Directory Selection Debug ===');

  // Test spreadsheet and sheet access
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log(`✓ Successfully opened spreadsheet: ${spreadsheet.getName()}`);
    
    const staffSheet = spreadsheet.getSheetByName(STAFF_DIRECTORY_SHEET_NAME);
    if (!staffSheet) {
      const message = `✗ Staff Directory sheet '${STAFF_DIRECTORY_SHEET_NAME}' not found.`;
      Logger.log(message);
      SpreadsheetApp.getUi().alert('Debug Result', message, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    Logger.log(`✓ Successfully found Staff Directory sheet`);

    // Test master presentation access
    try {
      const masterPresentation = SlidesApp.openById(MASTER_PRESENTATION_ID);
      Logger.log(`✓ Successfully accessed master presentation: ${masterPresentation.getName()}`);
    } catch (e) {
      Logger.log(`✗ Cannot access master presentation: ${e.message}`);
    }

    // Test selection
    let selectedRange = null;
    const activeSheet = SpreadsheetApp.getActiveSheet();
    Logger.log(`Current active sheet: ${activeSheet.getName()}`);
    
    if (activeSheet && activeSheet.getName() === STAFF_DIRECTORY_SHEET_NAME) {
      selectedRange = activeSheet.getActiveRange();
      Logger.log(`✓ Got selection from active sheet: ${selectedRange.getA1Notation()}`);
    } else {
      Logger.log(`✗ Active sheet is not Staff Directory. Please switch to Staff Directory sheet.`);
      SpreadsheetApp.getUi().alert('Debug Result', 'Please make sure you are on the Staff Directory sheet and have selected some rows before running this debug function.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Test data reading
    let debugInfo = [`Active Sheet: ${activeSheet.getName()}`, `Selection: ${selectedRange.getA1Notation()}`, `Rows: ${selectedRange.getRow()} to ${selectedRange.getLastRow()}`, '', 'Row Data:'];
    
    for (let row = selectedRange.getRow(); row <= selectedRange.getLastRow(); row++) {
      if (row === 1) {
        debugInfo.push(`Row ${row}: [HEADER ROW - SKIPPED]`);
        continue;
      }
      
      const rowData = staffSheet.getRange(row, 1, 1, 5).getValues()[0];
      const firstName = rowData[COLUMNS.FIRST_NAME] ? String(rowData[COLUMNS.FIRST_NAME]).trim() : '';
      const lastName = rowData[COLUMNS.LAST_NAME] ? String(rowData[COLUMNS.LAST_NAME]).trim() : '';
      const email = rowData[COLUMNS.EMAIL] ? String(rowData[COLUMNS.EMAIL]).trim() : '';
      const existingUrl = rowData[COLUMNS.AGENDA_URL] ? String(rowData[COLUMNS.AGENDA_URL]).trim() : '';
      const existingId = rowData[COLUMNS.SLIDE_ID] ? String(rowData[COLUMNS.SLIDE_ID]).trim() : '';
      
      debugInfo.push(`Row ${row}: "${firstName}" | "${lastName}" | "${email}" | HasURL: ${!!existingUrl} | HasID: ${!!existingId}`);
      
      if (!firstName || !lastName || !email) {
        debugInfo.push(`  ⚠️  ISSUE: Missing required data (first name, last name, or email)`);
      } else {
        debugInfo.push(`  ✓ Valid data found`);
      }
    }

    const message = debugInfo.join('\n');
    Logger.log(message);
    SpreadsheetApp.getUi().alert('Debug Results', message, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (e) {
    const errorMessage = `Debug Error: ${e.message}`;
    Logger.log(errorMessage);
    SpreadsheetApp.getUi().alert('Debug Error', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Serves the HTML file for the web app interface.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Daily Agendas');
}

/**
 * Fetches the agenda data from the 'Current_Day_Agendas' sheet.
 * *** THIS FUNCTION HAS BEEN CORRECTED TO RETURN DATA IN THE EXPECTED FORMAT ***
 */
function getAgendaData() {
  try {
    // Validate CONSTANTS is available
    if (typeof CONSTANTS === 'undefined') {
      Logger.log('CRITICAL ERROR: CONSTANTS object is undefined. Constants.js may not be loaded.');
      return { error: 'Configuration error: CONSTANTS not loaded' };
    }

    const SPREADSHEET_ID = CONSTANTS.SPREADSHEET_ID;
    const DATA_SHEET_NAME = CONSTANTS.DATA_SHEET_NAME;

    // Validate required constants
    if (!SPREADSHEET_ID) {
      Logger.log('CRITICAL ERROR: SPREADSHEET_ID is not defined in CONSTANTS.');
      return { error: 'Configuration error: SPREADSHEET_ID not defined' };
    }

    if (!DATA_SHEET_NAME) {
      Logger.log('CRITICAL ERROR: DATA_SHEET_NAME is not defined in CONSTANTS.');
      return { error: 'Configuration error: DATA_SHEET_NAME not defined' };
    }

    Logger.log(`getAgendaData called - SPREADSHEET_ID: ${SPREADSHEET_ID}, DATA_SHEET_NAME: ${DATA_SHEET_NAME}`);

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

    const dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
    if (!dataSheet) {
      Logger.log(`ERROR: Data sheet '${DATA_SHEET_NAME}' not found.`);
      return { error: `Data sheet '${DATA_SHEET_NAME}' not found` };
    }

    const range = dataSheet.getDataRange();
    const values = range.getValues();
    const formulas = range.getFormulas();

    if (values.length <= 1) {
      Logger.log('No data rows found in Current_Day_Agendas sheet (only headers or empty).');
      return { payload: [] }; // Return the expected structure with an empty array
    }

    const headers = values[0];
    const data = [];

    Logger.log(`Processing ${values.length - 1} rows of data...`);

    for (let i = 1; i < values.length; i++) {
      const obj = {};
      const currentRowValues = values[i];
      const currentRowFormulas = formulas[i];

      headers.forEach((header, j) => {
        const cleanedHeader = header.replace(/[^a-zA-Z0-9]/g, '');
        if (currentRowFormulas[j]) {
          obj[cleanedHeader] = currentRowFormulas[j];
        } else {
          obj[cleanedHeader] = currentRowValues[j];
        }
      });
      data.push(obj);
    }

    Logger.log(`Successfully prepared ${data.length} agenda records.`);

    // Return the data wrapped in a 'payload' object, as the client expects.
    return { payload: data };

  } catch (e) {
    Logger.log(`ERROR in getAgendaData: ${e.message}`);
    Logger.log(`Stack trace: ${e.stack}`);
    // Return an error object that the client can understand.
    return { error: `Failed to fetch agenda data: ${e.message}` };
  }
}


/**
 * Placeholder function for creating a PDF from selected agenda data.
 */
function createPdf(selectedData) {
  Logger.log("createPdf function called on server, but PDF generation is handled client-side.");
  return "";
}
