# Agenda Archive System Implementation Plan

## Overview
Currently, the system overwrites previous day's agenda data each time `extractTextForCurrentDayAgenda()` runs via `dataSheet.clearContents()` at Code.js:83. We need to implement an archive system that preserves historical data and provides calendar-based access through the web app.

## Implementation Steps

### 1. Backend Changes (Code.js)

**Modify Data Storage Approach:**
- Replace current overwrite behavior with append-to-archive approach
- Add date column to all archived records
- Create new archive sheet management functions
- Update `extractTextForCurrentDayAgenda()` to write to archive instead of clearing

**New Functions to Add:**
- `getOrCreateArchiveSheet()` - Creates monthly archive sheets (e.g., "Archive_2024_01")
- `archiveCurrentDayData()` - Moves current day data to appropriate archive sheet
- `getArchivedDataForDate(date)` - Retrieves agenda data for specific date
- `getAvailableArchiveDates()` - Returns list of dates with archived data

### 2. Archive Sheet Structure
- Monthly sheets named: "Archive_YYYY_MM" (e.g., "Archive_2024_01")
- Schema: Date | Teacher Last Name | Class Name | Day of Week | Turn In | Activities | Practice Work | Upcoming | Grade Level
- Keep current "Current_Day_Agendas" sheet for today's data display

### 3. Frontend Changes (index.html)

**Calendar UI Component:**
- Add calendar icon button to header
- Implement date picker modal using native HTML5 date input
- Style to match existing design system
- Position near current date display

**Archive Data Integration:**
- Add `getArchivedAgendaData(date)` Google Apps Script function call
- Modify existing `fetchData()` to accept optional date parameter
- Update UI state management to handle current vs. archived data viewing
- Add visual indicator when viewing archived data

**UI State Management:**
- Show "Current Day" vs "Archive: [Date]" in header
- Add "Back to Today" button when viewing archives
- Disable PDF/print functions for archived data or modify to show archive date

### 4. Constants Updates
- Add `ARCHIVE_SHEET_PREFIX: 'Archive_'` to CONSTANTS object
- No changes needed to existing coordinate system

### 5. Data Migration Strategy
- Current "Current_Day_Agendas" remains as-is for today's data
- Historical data (if any exists) can be manually moved to archive sheets
- New system starts fresh with next data collection run

## Files to Modify
- `Code.js` - Add archive functions, modify extraction logic
- `index.html` - Add calendar UI and archive data fetching
- `Constants.js` - Add archive sheet prefix constant

## Technical Considerations
- Archive sheets auto-created monthly to prevent single sheet size limits  
- Date-based partitioning enables efficient data retrieval
- Maintains backward compatibility with existing web app functionality
- Preserves all current features while adding archive capability