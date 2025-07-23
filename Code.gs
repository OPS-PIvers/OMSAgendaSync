/**
 * Extracts text from specific text boxes on Google Slide presentations
 * for the current day of the week, as listed in a config sheet.
 * Appends the extracted data to a main data sheet.
 */
function extractTextForCurrentDayAgenda() {
  // IMPORTANT: Your actual Google Sheet ID provided by you
  const SPREADSHEET_ID = '1nlrti40eQpWJsmfbszM8ARaa3i7sd9Uu4-vFAPFWqZg'; 
  const CONFIG_SHEET_NAME = 'Presentation_IDs';
  const DATA_SHEET_NAME = 'Current_Day_Agendas';

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

  // Clear previous data in the data sheet for the current day's run
  dataSheet.clearContents(); 
  dataSheet.appendRow([
    'Teacher Last Name',
    'Class Name',
    'Day of Week',
    'Top Box Text (Turn in:)',
    'Middle Box Text (Daily Activities:)',
    'Bottom Box Text (Practice Work:)',
    'Upcoming Text',
    'Grade Level'
  ]); // Re-add header row

  // --- Define ALL Text Box Coordinates (in points) ---
  // These are derived from your screenshots (1 inch = 72 points)
  // Format: { x, y, width, height }
  const BOX_COORDINATES = {
    'Monday': {
      top: { x: 34.56, y: 102.24, width: 123.12, height: 23.04 },
      middle: { x: 33.12, y: 150.48, width: 123.12, height: 121.68 },
      bottom: { x: 32.4, y: 296.64, width: 123.12, height: 23.04 }
    },
    'Tuesday': {
      top: { x: 173.52, y: 102.24, width: 123.12, height: 23.04 },
      middle: { x: 174.24, y: 150.48, width: 123.12, height: 121.68 },
      bottom: { x: 174.24, y: 296.64, width: 123.12, height: 23.04 }
    },
    'Wednesday': {
      top: { x: 312.48, y: 102.24, width: 123.12, height: 23.04 },
      middle: { x: 311.76, y: 150.48, width: 123.12, height: 121.68 },
      bottom: { x: 174.24, y: 296.64, width: 123.12, height: 23.04 } // Note: Wednesday bottom X seems to match Tuesday's in your data. Double check if this is intended.
    },
    'Thursday': {
      top: { x: 450.72, y: 102.24, width: 123.12, height: 23.04 },
      middle: { x: 453.60, y: 150.48, width: 123.12, height: 121.68 },
      bottom: { x: 452.16, y: 296.64, width: 123.12, height: 23.04 }
    },
    'Friday': {
      top: { x: 592.56, y: 102.24, width: 123.12, height: 23.04 },
      middle: { x: 593.28, y: 150.48, width: 123.12, height: 121.68 },
      bottom: { x: 593.28, y: 296.64, width: 123.12, height: 23.04 }
    },
    'Upcoming': { x: 218.16, y: 329.04, width: 355.68, height: 23.04 } // This box is consistent across days
  };

  // Get the current day of the week (e.g., "Monday", "Tuesday")
  const today = new Date();
  const dayOfWeek = Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE'); // 'EEEE' gives full day name

  // Check if it's a weekday we're interested in and coordinates are defined
  if (!BOX_COORDINATES.hasOwnProperty(dayOfWeek)) {
    Logger.log(`Today is ${dayOfWeek}. No agenda extraction scheduled for this day.`);
    SpreadsheetApp.getUi().alert('No agenda extraction scheduled for today (' + dayOfWeek + ').');
    return;
  }

  const currentDayBoxes = BOX_COORDINATES[dayOfWeek];
  const upcomingBox = BOX_COORDINATES['Upcoming']; // Always use the same upcoming box

  // Read presentation IDs and metadata from the config sheet
  // Assumes data is in Column A (Slide ID), B (Teacher Last Name), C (Class Name), D (Grade Level)
  const configDataRange = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 4); // Start from row 2, 4 columns
  const configValues = configDataRange.getValues();

  if (configValues.length === 0 || configValues[0].every(cell => !cell)) {
    Logger.log('No presentation IDs found in the configuration sheet.');
    SpreadsheetApp.getUi().alert('No presentation IDs found in the configuration sheet. Please add them to the "' + CONFIG_SHEET_NAME + '" tab.');
    return;
  }

  Logger.log(`Found ${configValues.length} presentation entries to process.`);

  configValues.forEach(row => {
    const presentationId = String(row[0]).trim(); // Column A
    const teacherLastName = String(row[1]).trim(); // Column B
    const className = String(row[2]).trim();     // Column C
    const gradeLevel = String(row[3]).trim();    // Column D

    if (!presentationId) {
      Logger.log('Skipping empty presentation ID row.');
      return; // Skip if Slide ID is empty
    }

    try {
      const presentation = SlidesApp.openById(presentationId);
      const slides = presentation.getSlides();

      if (slides.length === 0) {
        Logger.log(`Presentation ${presentationId} has no slides. Skipping.`);
        dataSheet.appendRow([teacherLastName, className, dayOfWeek, 'N/A', 'N/A', 'N/A', 'N/A', gradeLevel, `Error: No slides in presentation`]);
        return;
      }

      const firstSlide = slides[0]; // Assuming agenda is always on the first slide
      const pageElements = firstSlide.getPageElements();

      let topBoxText = '';
      let midBoxText = '';
      let botBoxText = '';
      let upcomingText = '';

      // Set a small tolerance for position and size matching
      const tolerance = 5; // points

      // Helper function to check if a shape matches a target box within tolerance
      const matchesBox = (shape, targetBox) => {
        const shapeX = shape.getLeft();
        const shapeY = shape.getTop();
        const shapeW = shape.getWidth();
        const shapeH = shape.getHeight();

        return Math.abs(shapeX - targetBox.x) < tolerance &&
               Math.abs(shapeY - targetBox.y) < tolerance &&
               Math.abs(shapeW - targetBox.width) < tolerance &&
               Math.abs(shapeH - targetBox.height) < tolerance;
      };

      pageElements.forEach(element => {
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          const shape = element.asShape();
          const textContent = shape.getText().asString().trim();
          
          if (textContent === '') return; // Skip empty text boxes

          if (matchesBox(shape, currentDayBoxes.top)) {
            topBoxText = textContent;
          } else if (matchesBox(shape, currentDayBoxes.middle)) {
            midBoxText = textContent;
          } else if (matchesBox(shape, currentDayBoxes.bottom)) {
            botBoxText = textContent;
          } else if (matchesBox(shape, upcomingBox)) {
            upcomingText = textContent;
          }
        }
      });

      // Append the collected data to the 'Current_Day_Agendas' sheet
      dataSheet.appendRow([
        teacherLastName,
        className,
        dayOfWeek,
        topBoxText,
        midBoxText,
        botBoxText,
        upcomingText,
        gradeLevel
      ]);
      Logger.log(`Processed: ${teacherLastName} - ${className} for ${dayOfWeek}`);

    } catch (e) {
      Logger.log(`Error processing presentation ID ${presentationId} (${teacherLastName}, ${className}): ${e.message}`);
      dataSheet.appendRow([teacherLastName, className, dayOfWeek, 'ERROR', 'ERROR', 'ERROR', 'ERROR', gradeLevel, `Error: ${e.message}`]);
    }
  });

  Logger.log('All text extraction for current day complete and data appended to Google Sheet.');
  SpreadsheetApp.getUi().alert(
    'Daily Agenda Extraction Complete!',
    'Data for ' + dayOfWeek + ' has been extracted and compiled into the "' + DATA_SHEET_NAME + '" tab.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Helper function to add custom menu (for manual trigger)
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Agenda Tools')
      .addItem('Run Daily Agenda Extraction Now', 'extractTextForCurrentDayAgenda')
      .addToUi();
}

/**
 * Serves the HTML file for the web app.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Daily Agendas');
}

/**
 * Fetches the agenda data from the 'Current_Day_Agendas' sheet.
 * This function is called by the client-side JavaScript in the web app.
 * @returns {Array<Object>} An array of objects, where each object represents a row of agenda data.
 */
function getAgendaData() {
  // IMPORTANT: Your actual Google Sheet ID provided by you
  const SPREADSHEET_ID = '1nlrti40eQpWJsmfbszM8ARaa3i7sd9Uu4-vFAPFWqZg'; 
  const DATA_SHEET_NAME = 'Current_Day_Agendas';

  try {
    Logger.log(`Attempting to open spreadsheet with ID: ${SPREADSHEET_ID}`); 
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
    if (!dataSheet) {
      Logger.log(`Error in getAgendaData: Data sheet '${DATA_SHEET_NAME}' not found.`);
      return []; 
    }

    const range = dataSheet.getDataRange();
    const values = range.getValues();

    if (values.length <= 1) { // Only header row or empty
      Logger.log('No data found in Current_Day_Agendas sheet (only headers or empty).');
      return [];
    }

    const headers = values[0];
    const data = values.slice(1).map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        // Clean up header names for easier JavaScript access (e.g., remove spaces, special chars)
        const cleanedHeader = header.replace(/[^a-zA-Z0-9]/g, ''); // Remove non-alphanumeric
        obj[cleanedHeader] = row[i];
      });
      return obj;
    });
    return data;
  } catch (e) {
    Logger.log(`Error in getAgendaData: ${e.message}`);
    throw new Error(`Failed to fetch agenda data: ${e.message}`);
  }
}
