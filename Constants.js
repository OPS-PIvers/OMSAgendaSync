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
   * A tolerance value (in points) for matching the position and size of shapes on the slides.
   * This helps account for minor variations in shape placement.
   * @type {number}
   */
  TOLERANCE: 5,

  /**
   * An object containing the precise coordinates and dimensions (x, y, width, height)
   * for the text boxes to be extracted from the Google Slides.
   * The coordinates are organized by the day of the week.
   * @type {Object.<string, Object>}
   */
  BOX_COORDINATES: {
    'Monday': {
      top: { x: 43.50, y: 129.64, width: 153.17, height: 33.87 },    // "Turn In"
      middle: { x: 43.50, y: 198.31, width: 153.17, height: 101.06 }, // "Activities"
      bottom: { x: 42.71, y: 334.90, width: 153.17, height: 45.14 }  // "Practice Work"
    },
    'Tuesday': {
      top: { x: 212.61, y: 129.64, width: 157.58, height: 33.87 },   // "Turn In"
      middle: { x: 212.61, y: 198.31, width: 157.58, height: 101.06 },// "Activities"
      bottom: { x: 211.82, y: 334.90, width: 157.58, height: 45.14 } // "Practice Work"
    },
    'Wednesday': {
      top: { x: 383.29, y: 129.64, width: 157.58, height: 33.87 },   // "Turn In"
      middle: { x: 383.29, y: 198.31, width: 157.58, height: 101.06 },// "Activities"
      bottom: { x: 382.50, y: 334.90, width: 157.58, height: 45.14 } // "Practice Work"
    },
    'Thursday': {
      top: { x: 553.98, y: 129.64, width: 157.58, height: 34.72 },   // "Turn In"
      middle: { x: 553.98, y: 198.31, width: 157.58, height: 101.06 },// "Activities"
      bottom: { x: 553.19, y: 334.90, width: 157.58, height: 45.14 } // "Practice Work"
    },
    'Friday': {
      top: { x: 727.50, y: 129.64, width: 161.06, height: 34.72 },   // "Turn In"
      middle: { x: 727.50, y: 198.31, width: 161.06, height: 101.06 },// "Activities"
      bottom: { x: 726.71, y: 334.90, width: 161.06, height: 45.14 } // "Practice Work"
    },
    'Upcoming': { x: 148.66, y: 392.40, width: 709.13, height: 31.23 }
  }
};
