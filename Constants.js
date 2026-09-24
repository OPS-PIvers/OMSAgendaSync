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
   * The name of the sheet listing agenda fields the last extraction run could not
   * find (no text box in that area of the slide, or no slide for this week).
   * Rewritten on every run, so it always reflects the current state.
   * @type {string}
   */
  ISSUES_SHEET_NAME: 'Extraction_Issues',

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
   * The slide areas each agenda field is read from, in points. A text box
   * belongs to the zone its centre point falls in, so teachers can move or
   * resize a box freely as long as its centre stays on the right card.
   * Every boundary sits in the gap between two cards of the template.
   * Ranges are [start, end): a centre exactly on a boundary goes to the later zone.
   * Anything above ROWS.top[0] (title, "WEEK OF", day headers) is ignored.
   * @type {Object}
   */
  ZONES: {
    // Day columns (x). The outer edges are open so nothing falls off the slide.
    COLUMNS: {
      'Monday': [-Infinity, 204.2],
      'Tuesday': [204.2, 376.3],
      'Wednesday': [376.3, 547.0],
      'Thursday': [547.0, 719.1],
      'Friday': [719.1, Infinity]
    },
    // Field rows (y) within each day column.
    ROWS: {
      top: [104, 179.4],     // "Turn In"
      middle: [179.4, 314.2], // "Activities"
      bottom: [314.2, 386.2]  // "Practice Work"
    },
    // The Upcoming strip spans the full slide width below the day cards.
    UPCOMING: [386.2, Infinity]
  }
};
