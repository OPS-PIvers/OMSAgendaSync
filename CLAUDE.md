# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

OMSAgendaSync is a Google Apps Script web application that extracts agenda text from Google Slides presentations and displays them in a web interface. The system automatically syncs teacher agendas from standardized Google Slides templates and presents them in a searchable, filterable web interface with PDF export capabilities.

## Architecture

### Core Components

- **Code.js**: Main Google Apps Script file containing all backend logic for:
  - Text extraction from Google Slides by matching text boxes to slide zones
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
  - Slide zones (day columns and field rows) used to locate agenda text boxes
  - Sheet names and column mappings
  - Staff directory structure

- **appsscript.json**: Google Apps Script project configuration with:
  - Sheets API v4 enabled
  - Web app deployment settings (anonymous access)
  - V8 runtime configuration

## Key Data Flow

1. **Slides → Script**: `extractTextForCurrentDayAgenda()` extracts text from the text boxes in each day's zones of the Google Slides presentations
2. **Script → Sheets**: Processed data is stored in "Current_Day_Agendas" sheet with hyperlink preservation
3. **Sheets → Web App**: Frontend loads data via `google.script.run` calls to backend functions
4. **Archive Process**: Daily trigger moves current data to archive sheets (e.g., "Archive_2024_09")

## Text Extraction System

The application finds agenda text by which zone of the slide each text box sits in:
- Each day of the week has three fields: "Turn In", "Activities", "Practice Work", plus a full-width "Upcoming" strip
- Zones are defined in `CONSTANTS.ZONES` as day columns (x) and field rows (y), in points, with every boundary in the gap between two template cards
- A text box belongs to the zone its **centre point** falls in, so teachers can move or resize boxes freely as long as the centre stays on the right card
- Several boxes in one zone are joined top to bottom; an empty box counts as found but blank
- The template's cards, labels and day headers are the slide background image, so only teacher text boxes (and the title) are shapes
- Fields with no text box in their zone, and decks with no slide for this week, are listed on the "Extraction_Issues" sheet, rewritten every run
- The system preserves hyperlinks by converting them to Google Sheets HYPERLINK formulas

## Development Commands

### Deployment Process

Deployment is automated. Pushing to `main` triggers
`.github/workflows/deploy.yml`, which runs `clasp push`, creates a script
version described by the head commit, points the live web app deployment at
that version, and fails the job if the deployment did not actually move.

```bash
git add .
git commit -m "FEAT: Your descriptive commit message here"
git push
```

The workflow only fires when `Code.js`, `Constants.js`, `index.html`,
`appsscript.json`, `.clasp.json`, `.claspignore`, or the workflow file itself
changes. To run it by hand:

```bash
gh workflow run "Deploy to Apps Script" --repo OPS-PIvers/OMSAgendaSync
```

It authenticates with the `CLASPRC_JSON` repository secret, which holds the
contents of a local `~/.clasprc.json`. If deploys start failing at the
"Authenticate clasp" step, that token was revoked or expired — run `clasp login`
locally, then re-upload it (PowerShell):

```bash
Get-Content -Raw "$env:USERPROFILE\.clasprc.json" | gh secret set CLASPRC_JSON --repo OPS-PIvers/OMSAgendaSync
```

### Manual Deployment

Only needed to deploy uncommitted work or to roll back to an older version. CI
will overwrite a manual deployment on the next push to `main`.

```bash
clasp push
clasp create-version "Description of the change"
clasp redeploy AKfycbwtKGbS9PtKwSVgHUsN03r451weFHmEkK2QrtsLx0_XwmDoiFWa53rwXcn3TqoFRSKDWg --versionNumber <VERSION>
```

### Google Apps Script Commands

```bash
# Push code to GAS project
clasp push

# Pull latest code from GAS project
clasp pull

# List the files that would be pushed
clasp status

# List deployments and the version each serves (--json for machine-readable output)
clasp deployments

# List script versions
clasp versions
```

## Configuration

### Google Sheets Structure
- **Presentation_IDs**: Configuration sheet with teacher information and slide IDs
- **Current_Day_Agendas**: Live data storage for extracted agenda content  
- **Staff Directory**: Teacher contact information and presentation links
- **Extraction_Issues**: Agenda fields the latest extraction run could not find
- **Archive_YYYY_MM**: Monthly archive sheets for historical data

### Key Constants to Update
- `SPREADSHEET_ID`: Main Google Sheet containing all configuration and data
- `MASTER_PRESENTATION_ID`: Template presentation copied for new teachers
- `ZONES`: Slide zones (in points) that each agenda field is read from

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
- Text extraction relies on the template's card layout - if the template's cards move, update `CONSTANTS.ZONES` to match
- Never write a full `script.google.com/macros/s/.../exec` URL inside an inline `<script>` block in `index.html`. HtmlService's sanitizer strips macro URLs from script content (though not from markup attributes), truncating the string literal and breaking the entire script so no event listeners bind. Read the URL off an element's `href` instead.