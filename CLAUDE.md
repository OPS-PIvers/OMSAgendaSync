# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

OMSAgendaSync is a Google Apps Script web application that extracts agenda text from Google Slides presentations and displays them in a web interface. The system automatically syncs teacher agendas from standardized Google Slides templates and presents them in a searchable, filterable web interface with PDF export capabilities.

## Architecture

### Core Components

- **Code.js**: Main Google Apps Script file containing all backend logic for:
  - Text extraction from Google Slides using precise coordinate matching
  - Data processing and storage in Google Sheets
  - Trigger management for automated extraction and archiving
  - Web app server-side functions

- **index.html**: Single-page web application frontend with:
  - Tailwind CSS styling and responsive design
  - Grade-level filtering and teacher search functionality
  - PDF export capabilities using jsPDF
  - Real-time data loading from Google Apps Script backend

- **Constants.js**: Configuration file containing:
  - Google Sheets and Slides IDs
  - Precise text box coordinates for each day of the week
  - Sheet names and column mappings
  - Staff directory structure

- **appsscript.json**: Google Apps Script project configuration with:
  - Sheets API v4 enabled
  - Web app deployment settings (anonymous access)
  - V8 runtime configuration

## Key Data Flow

1. **Slides → Script**: `extractTextForCurrentDayAgenda()` extracts text from specific coordinates in Google Slides presentations
2. **Script → Sheets**: Processed data is stored in "Current_Day_Agendas" sheet with hyperlink preservation
3. **Sheets → Web App**: Frontend loads data via `google.script.run` calls to backend functions
4. **Archive Process**: Daily trigger moves current data to archive sheets (e.g., "Archive_2024_09")

## Text Extraction System

The application uses precise coordinate matching to extract text from specific areas of Google Slides templates:
- Each day of the week has three text boxes: "Turn In", "Activities", "Practice Work"
- Coordinates are defined in `CONSTANTS.BOX_COORDINATES` with tolerance for minor variations
- The system preserves hyperlinks by converting them to Google Sheets HYPERLINK formulas

## Development Commands

### Deployment Process (from GEMINI.md)
After any code changes, follow this sequence:

```bash
# Stage and commit changes
git add .
git commit -m "FEAT: Your descriptive commit message here"

# Push to remote repository
git push

# Deploy to Google Apps Script
clasp push

# Update web app deployment (replace with actual version number)
clasp redeploy AKfycbwtKGbS9PtKwSVgHUsN03r451weFHmEkK2QrtsLx0_XwmDoiFWa53rwXcn3TqoFRSKDWg --versionNumber <LATEST_VERSION_NUMBER>
```

### Google Apps Script Commands

```bash
# Push code to GAS project
clasp push

# Pull latest code from GAS project  
clasp pull

# Check deployment status
clasp status

# List all deployments
clasp deployments

# View project info
clasp version
```

## Configuration

### Google Sheets Structure
- **Presentation_IDs**: Configuration sheet with teacher information and slide IDs
- **Current_Day_Agendas**: Live data storage for extracted agenda content  
- **Staff Directory**: Teacher contact information and presentation links
- **Archive_YYYY_MM**: Monthly archive sheets for historical data

### Key Constants to Update
- `SPREADSHEET_ID`: Main Google Sheet containing all configuration and data
- `MASTER_PRESENTATION_ID`: Template presentation copied for new teachers
- `BOX_COORDINATES`: Precise pixel coordinates for text extraction areas

## Automation

The system includes two automated triggers:
- **Hourly Extraction**: `runDailyExtractionTrigger()` extracts current day agenda data
- **Daily Archive**: `runDailyArchiveTrigger()` moves data to archive sheets at 11:30 PM

## Web App Features

- Grade-level filtering (6th, 7th, 8th grades)
- Teacher name search functionality
- PDF export with selective teacher inclusion
- Responsive design using Tailwind CSS
- Real-time data synchronization with Google Sheets

## Important Notes

- Clear browser cache or use incognito mode after redeployment to see changes
- The web app has anonymous access for school-wide availability
- All hyperlinks in slides are preserved as clickable HYPERLINK formulas in sheets
- Text extraction relies on precise coordinate matching - template modifications require coordinate updates