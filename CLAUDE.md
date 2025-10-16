# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a **Google Apps Script** project that implements an Inventory Management System for tracking physical items in storage boxes. The system integrates with Google Sheets (for data storage), Google Forms (for data entry), Google Docs (for report generation), and Gmail (for notifications).

## File Structure

- `p1.txt` - Main Apps Script code file (should be deployed as `Code.gs` in Google Apps Script)
- `p2.txt` - HTML sidebar interface (should be deployed as `Sidebar.html` in Google Apps Script)

**Important**: These `.txt` files are local copies. The actual deployment happens in the Google Apps Script editor at script.google.com.

## Development Workflow

### Deployment Process
1. Open the Google Sheets spreadsheet where this script is bound
2. Navigate to Extensions > Apps Script
3. Copy contents of `p1.txt` into `Code.gs`
4. Create a new HTML file named `Sidebar.html` and copy contents of `p2.txt`
5. Save and test using the spreadsheet

### Testing
- Test the `onOpen()` trigger by refreshing the spreadsheet
- Use "Run" button in Apps Script editor to test individual functions
- Check Execution log (View > Logs) for debugging output
- Test form integration by submitting test entries through the Google Form

### Common Modifications
When making changes, always test in the Apps Script editor before copying back to local files:
1. Edit code in Apps Script editor
2. Test thoroughly
3. Copy working code back to `p1.txt` or `p2.txt`

## Architecture

### Data Flow
1. **Input**: Users submit items via Google Form (referenced in `getFormUrl()`)
2. **Storage**: Form responses populate "Form Responses 1" sheet in Google Sheets
3. **Processing**: Apps Script reads sheet data, generates reports/emails
4. **Output**: Google Docs reports stored in "Inventory System/Box Reports" Drive folder

### Key Components

**Main Script (p1.txt / Code.gs)**
- `onOpen()` - Automatic menu creation when spreadsheet opens
- `generateSingleBoxReport(boxNumber)` - Creates Google Doc report for a specific box
- `generateAllBoxReports()` - Bulk report generation for all boxes
- `emailBoxSummary(boxNumber, email)` - Sends HTML email summary
- `getRecentImages(limit)` - Retrieves recent image uploads for sidebar display
- `createBoxDocument(boxNumber, items)` - Formats and creates Google Doc with statistics

**Sidebar Interface (p2.txt / Sidebar.html)**
- Control panel with gradient UI design
- Quick actions: Open form, refresh data
- Box-specific operations: Generate report, email summary
- Recent images thumbnail grid with lazy loading
- Real-time status notifications

### Data Schema
The script expects these columns in "Form Responses 1" sheet:
- Timestamp
- Item name
- Item description
- Item Category
- Quantity
- Room location
- Storage box Number
- Action Plan
- Image Upload

### Google Drive Structure
```
Inventory System/          (parent folder)
  └── Box Reports/         (generated reports stored here)
```

## Important Constraints

### Google Apps Script Limitations
- **Execution time limit**: 6 minutes for consumer accounts
- **Trigger quota**: Limited daily executions for time-driven triggers
- **UrlFetch quota**: Limited external requests per day
- **Email quota**: ~100 emails per day for consumer Gmail accounts

### Script Dependencies
- Requires spreadsheet to have "Form Responses 1" sheet
- Form URL hardcoded in `getFormUrl()` at line 44
- Uses `Session.getScriptTimeZone()` for date formatting
- Requires Drive folder creation permissions

## Form URL Configuration

The Google Form URL is hardcoded in `getFormUrl()` (p1.txt:44). To update:
```javascript
function getFormUrl() {
  return 'YOUR_FORM_URL_HERE';
}
```

## Common Issues

### "Form Responses 1" not found
- Ensure Google Form is linked to the spreadsheet
- Verify sheet name is exactly "Form Responses 1"

### Images not displaying in sidebar
- Image URLs must be publicly accessible
- Check Google Drive sharing settings on uploaded images
- Verify "Image Upload" column exists in responses

### Reports not generating
- Check execution logs for column index errors
- Verify all expected columns exist in sheet
- Ensure Drive folder permissions allow file creation

### Popup blockers
- Sidebar includes fallback link if `window.open()` is blocked
- Users may need to allow popups for the spreadsheet domain

## Code Patterns

### Error Handling
Functions use try-catch with console.error logging and user-facing error messages via UI alerts or toast notifications.

### Server-Client Communication
Sidebar uses `google.script.run` with success/failure handlers:
```javascript
google.script.run
  .withSuccessHandler(callback)
  .withFailureHandler(errorHandler)
  .serverFunction(params);
```

### Document Generation
Reports use DocumentApp API with styled paragraphs, headings, tables, and horizontal rules for visual separation.
