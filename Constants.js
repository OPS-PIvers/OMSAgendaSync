/**
 * @fileoverview This file contains constants used throughout the Google Apps Script project.
 * Centralizing constants here makes the code more maintainable and easier to configure.
 */

const CONSTANTS = {
  /**
   * The unique identifier of the Google Sheet that contains the configuration and data.
   * @type {string}
   */
  SPREADSHEET_ID: '1nlrti40eQpWJsmfbszM8ARaa3i7sd9Uu4-vFAPFWqZg',

  /**
   * The name of the sheet that contains the list of Google Slide presentation IDs,
   * teacher names, class names, and grade levels.
   * @type {string}
   */
  CONFIG_SHEET_NAME: 'Presentation_IDs',

  /**
   * The name of the sheet where the extracted agenda data will be stored.
   * @type {string}
   */
  DATA_SHEET_NAME: 'Current_Day_Agendas',

  /**
   * The prefix used for archive sheet names. Archive sheets are named with this prefix
   * followed by year and month (e.g., 'Archive_2024_01').
   * @type {string}
   */
  ARCHIVE_SHEET_PREFIX: 'Archive_',

  /**
   * A tolerance value (in points) for matching the position and size of shapes on the slides.
   * This helps account for minor variations in shape placement.
   * @type {number}
   */
  TOLERANCE: 5,

  /**
   * The name of the sheet that contains the staff directory with columns:
   * A: First Name, B: Last Name, C: Email Address, D: Agenda URL, E: Slide ID
   * @type {string}
   */
  STAFF_DIRECTORY_SHEET_NAME: 'Staff Directory',

  /**
   * The Google Slides presentation ID that will be copied for each teacher.
   * This should be set to the master template presentation ID.
   * @type {string}
   */
  MASTER_PRESENTATION_ID: '1QO9b7830WZmmWgPB5ZWm1-QkI5bvqqj6i_pUo85EFkw',

  /**
   * Column indices for the Staff Directory sheet (0-based indexing)
   * @type {Object}
   */
  STAFF_DIRECTORY_COLUMNS: {
    FIRST_NAME: 0,    // Column A
    LAST_NAME: 1,     // Column B
    EMAIL: 2,         // Column C
    AGENDA_URL: 3,    // Column D
    SLIDE_ID: 4       // Column E
  },

  /**
   * An object containing the precise coordinates and dimensions (x, y, width, height)
   * for the text boxes to be extracted from the Google Slides.
   * The coordinates are organized by the day of the week.
   * @type {Object.<string, Object>}
   */
  BOX_COORDINATES: {
    'Monday': {
      top: { x: 43.50, y: 124.70, width: 153.17, height: 38.69 },    // "Turn In"
      middle: { x: 43.50, y: 194.49, width: 153.17, height: 104.88 }, // "Activities"
      bottom: { x: 42.71, y: 329.03, width: 153.17, height: 51.02 }  // "Practice Work"
    },
    'Tuesday': {
      top: { x: 212.61, y: 124.70, width: 157.58, height: 38.69 },   // "Turn In"
      middle: { x: 212.61, y: 194.49, width: 157.58, height: 104.88 },// "Activities"
      bottom: { x: 211.82, y: 329.03, width: 157.58, height: 51.02 } // "Practice Work"
    },
    'Wednesday': {
      top: { x: 383.29, y: 124.70, width: 157.58, height: 38.69 },   // "Turn In"
      middle: { x: 383.29, y: 194.49, width: 157.58, height: 104.88 },// "Activities"
      bottom: { x: 382.50, y: 329.03, width: 157.58, height: 51.02 } // "Practice Work"
    },
    'Thursday': {
      top: { x: 553.98, y: 124.70, width: 157.58, height: 39.66 },   // "Turn In"
      middle: { x: 553.98, y: 194.49, width: 157.58, height: 104.88 },// "Activities"
      bottom: { x: 553.19, y: 329.03, width: 157.58, height: 51.02 } // "Practice Work"
    },
    'Friday': {
      top: { x: 727.50, y: 124.70, width: 161.06, height: 39.66 },   // "Turn In"
      middle: { x: 727.50, y: 194.49, width: 161.06, height: 104.88 },// "Activities"
      bottom: { x: 726.71, y: 329.03, width: 161.06, height: 51.02 } // "Practice Work"
    },
    'Upcoming': { x: 148.66, y: 392.40, width: 709.13, height: 31.23 }
  }
};
