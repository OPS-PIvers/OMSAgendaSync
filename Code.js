/**
 * @fileoverview Main script file for extracting agenda text from Google Slides.
 * This file contains the core logic for processing presentations and interacting with the Google Sheet.
 * It relies on a separate 'Constants.gs' file for configuration.
 */

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
 * Extracts the full text from a shape and the URL of the first hyperlink found.
 * If a link is found, it returns a Google Sheets HYPERLINK formula.
 * @param {GoogleAppsScript.Slides.TextRange} textRange The TextRange from a shape.
 * @returns {string} The plain text content, or a HYPERLINK formula string if a link is found.
 */
function extractTextAndFirstLink(textRange) {
  const fullText = textRange.asString().trim();
  if (fullText === '') {
    return 'N/A'; // Return a default value for empty text boxes
  }

  let firstLinkUrl = null;
  const runs = textRange.getRuns();
  for (const run of runs) {
    // CORRECTED AND VERIFIED METHOD: Get the TextStyle, then the Link object, then the URL.
    const link = run.getTextStyle().getLink();
    if (link) {
      firstLinkUrl = link.getUrl();
      // If a URL is found, we've got what we need and can exit the loop.
      if (firstLinkUrl) {
        break;
      }
    }
  }

  if (firstLinkUrl) {
    // Create a formula for Google Sheets. We must escape double quotes within the text.
    return `=HYPERLINK("${firstLinkUrl}", "${fullText.replace(/"/g, '""')}")`;
  } else {
    return fullText;
  }
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
    Logger.log(`Error: Configuration sheet '${CONFIG_SHEET_NAME}' not found.`);
    SpreadsheetApp.getUi().alert(`Error: Configuration sheet '${CONFIG_SHEET_NAME}' not found. Please ensure it exists.`);
    return;
  }

  let dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
  if (!dataSheet) {
    Logger.log(`Error: Data sheet '${DATA_SHEET_NAME}' not found. Please ensure it exists.`);
    SpreadsheetApp.getUi().alert(`Error: Data sheet '${DATA_SHEET_NAME}' not found. Please ensure it exists.`);
    return;
  }

  dataSheet.clearContents();
  dataSheet.appendRow([
    'Teacher Last Name', 'Class Name', 'Day of Week', 'Turn In', 'Activities',
    'Practice Work', 'Upcoming', 'Grade Level'
  ]);

  const today = new Date();
  const dayOfWeek = dayToTest || Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE');
  Logger.log(`Running extraction for: ${dayOfWeek}`);

  if (!BOX_COORDINATES.hasOwnProperty(dayOfWeek)) {
    const message = dayToTest ?
      `The provided test day '${dayToTest}' has no coordinates defined.` :
      `Today is ${dayOfWeek}. No agenda extraction scheduled for this day.`;
    Logger.log(message);
    SpreadsheetApp.getUi().alert(message);
    return;
  }

  const monday = getMondayOfCurrentWeek();
  const formattedMonday = Utilities.formatDate(monday, Session.getScriptTimeZone(), 'M/d/yyyy');
  const weekOfText = `WEEK OF ${formattedMonday}`.toUpperCase();
  Logger.log(`Searching for slides with the text: "${weekOfText}"`);

  const currentDayBoxes = BOX_COORDINATES[dayOfWeek];
  const upcomingBox = BOX_COORDINATES['Upcoming'];

  const configDataRange = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 4);
  const configValues = configDataRange.getValues();

  if (configValues.length === 0 || configValues[0].every(cell => !cell)) {
    Logger.log('No presentation IDs found in the configuration sheet.');
    SpreadsheetApp.getUi().alert('No presentation IDs found in the configuration sheet.');
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
      for (const slide of slides) {
        const shapes = slide.getShapes();
        for (const shape of shapes) {
          if (shape.getText().asString().toUpperCase().includes(weekOfText)) {
            agendaSlide = slide;
            break;
          }
        }
        if (agendaSlide) {
          break;
        }
      }

      if (!agendaSlide) {
        throw new Error(`Slide with text "${weekOfText}" not found.`);
      }

      const pageElements = agendaSlide.getPageElements();
      let topBoxText = 'N/A', midBoxText = 'N/A', botBoxText = 'N/A', upcomingText = 'N/A';
      const tolerance = CONSTANTS.TOLERANCE;

      const matchesBox = (shape, targetBox) => {
        return Math.abs(shape.getLeft() - targetBox.x) < tolerance &&
               Math.abs(shape.getTop() - targetBox.y) < tolerance &&
               Math.abs(shape.getWidth() - targetBox.width) < tolerance &&
               Math.abs(shape.getHeight() - targetBox.height) < tolerance;
      };

      pageElements.forEach(element => {
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          const shape = element.asShape();
          const textRange = shape.getText();
          if (textRange.isEmpty()) return;

          const cellValue = extractTextAndFirstLink(textRange);

          if (matchesBox(shape, currentDayBoxes.top)) topBoxText = cellValue;
          else if (matchesBox(shape, currentDayBoxes.middle)) midBoxText = cellValue;
          else if (matchesBox(shape, currentDayBoxes.bottom)) botBoxText = cellValue;
          else if (matchesBox(shape, upcomingBox)) upcomingText = cellValue;
        }
      });

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

  Logger.log('All text extraction for current day complete and data appended to Google Sheet.');
  SpreadsheetApp.getUi().alert(
    'Daily Agenda Extraction Complete!',
    'Data for ' + dayOfWeek + ' has been extracted and compiled into the "' + DATA_SHEET_NAME + '" tab.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// --- NEW TESTING FUNCTIONS ---
function testForMonday() { extractTextForCurrentDayAgenda('Monday'); }
function testForTuesday() { extractTextForCurrentDayAgenda('Tuesday'); }
function testForWednesday() { extractTextForCurrentDayAgenda('Wednesday'); }
function testForThursday() { extractTextForCurrentDayAgenda('Thursday'); }
function testForFriday() { extractTextForCurrentDayAgenda('Friday'); }
// --- END NEW TESTING FUNCTIONS ---


/**
 * Creates a custom menu in the Google Sheet UI for manual script execution.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Agenda Tools')
    .addItem('Run Daily Agenda Extraction Now', 'extractTextForCurrentDayAgenda');

  const testSubMenu = ui.createMenu('Run Manual Test As...')
    .addItem('Monday', 'testForMonday')
    .addItem('Tuesday', 'testForTuesday')
    .addItem('Wednesday', 'testForWednesday')
    .addItem('Thursday', 'testForThursday')
    .addItem('Friday', 'testForFriday');

  menu.addSeparator()
      .addSubMenu(testSubMenu)
      .addToUi();
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
  const SPREADSHEET_ID = CONSTANTS.SPREADSHEET_ID;
  const DATA_SHEET_NAME = CONSTANTS.DATA_SHEET_NAME;

  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
    if (!dataSheet) {
      throw new Error(`Data sheet '${DATA_SHEET_NAME}' not found.`);
    }

    const range = dataSheet.getDataRange();
    const values = range.getValues();
    const formulas = range.getFormulas();

    if (values.length <= 1) {
      Logger.log('No data found in Current_Day_Agendas sheet.');
      return { payload: [] }; // Return the expected structure with an empty array
    }

    const headers = values[0];
    const data = [];

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
    
    // Return the data wrapped in a 'payload' object, as the client expects.
    return { payload: data };

  } catch (e) {
    Logger.log(`Error in getAgendaData: ${e.message}`);
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
