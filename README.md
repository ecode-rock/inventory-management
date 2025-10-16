# 📦 Inventory Management System

A Google Apps Script-based inventory management system for tracking physical items in storage boxes with QR code generation, automated reports, and email notifications.

## Features

- 📝 **Google Forms Integration** - Easy data entry via web form
- 📊 **Automated Report Generation** - Create detailed box reports with statistics
- 📧 **Email Summaries** - Send inventory summaries to any email address
- 🔗 **QR Code Generation** - Automatic QR codes for quick report access
- 📷 **Image Management** - Upload and view item images
- 🎨 **Modern UI** - Beautiful sidebar control panel with gradient design
- 📁 **Google Drive Integration** - Organized folder structure for reports and QR codes

## Project Structure

```
your-awesome-project/
├── Code.gs              # Main Apps Script code
├── Sidebar.html         # UI sidebar interface
├── appsscript.json      # Apps Script manifest
├── .claspignore         # Files to exclude from clasp push
├── .gitignore          # Files to exclude from Git
├── CLAUDE.md           # AI assistant guidance
└── README.md           # This file
```

## Setup Instructions

### Prerequisites

- Node.js (v14 or higher)
- Google account with access to Google Sheets
- clasp CLI tool (installed globally)

### Installation

1. **Clone this repository**
   ```bash
   git clone <your-repo-url>
   cd your-awesome-project
   ```

2. **Install clasp (if not already installed)**
   ```bash
   npm install -g @google/clasp
   ```

3. **Login to Google**
   ```bash
   clasp login
   ```
   This will open a browser window for authentication.

4. **Create a new Apps Script project OR link to existing**

   **For new project:**
   ```bash
   clasp create --title "Inventory Management System" --type sheets
   ```

   **For existing project:**
   - Get your Script ID from the Apps Script editor URL:
     `https://script.google.com/home/projects/[SCRIPT_ID]/edit`
   - Create `.clasp.json` file:
     ```json
     {
       "scriptId": "YOUR_SCRIPT_ID_HERE",
       "rootDir": "."
     }
     ```

5. **Push code to Apps Script**
   ```bash
   clasp push
   ```

6. **Open the project in browser**
   ```bash
   clasp open
   ```

## Development Workflow

### Daily Development

1. **Edit files locally** (Code.gs, Sidebar.html)
   ```bash
   code Code.gs
   ```

2. **Push changes to Apps Script**
   ```bash
   clasp push
   ```

3. **View execution logs**
   ```bash
   clasp logs
   ```

4. **Pull changes from Apps Script** (if edited in browser)
   ```bash
   clasp pull
   ```

### Git Workflow

1. **Stage changes**
   ```bash
   git add .
   ```

2. **Commit changes**
   ```bash
   git commit -m "Description of changes"
   ```

3. **Push to GitHub**
   ```bash
   git push origin main
   ```

## Configuration

### Update Form URL

Edit `Code.gs` line 44 to point to your Google Form:

```javascript
function getFormUrl() {
  return 'YOUR_GOOGLE_FORM_URL_HERE';
}
```

### Update Timezone

Edit `appsscript.json` to set your timezone:

```json
{
  "timeZone": "America/New_York"
}
```

Find your timezone: https://en.wikipedia.org/wiki/List_of_tz_database_time_zones

## Google Drive Folder Structure

The system automatically creates these folders in Google Drive:

```
Inventory System/
├── Box Reports/          # Generated box reports
└── qr code/             # QR code images
```

## Google Sheets Setup

1. Create a Google Sheet
2. Create a Google Form linked to the sheet
3. Form should collect these columns:
   - Timestamp
   - Item name
   - Item description
   - Item Category
   - Quantity
   - Room location
   - Storage box Number
   - Action Plan
   - Image Upload

4. Bind the Apps Script to the sheet:
   - Extensions → Apps Script
   - Use `clasp push` to deploy code

## Usage

### Via Spreadsheet Menu

1. Open your Google Sheet
2. Click **🗃️ Inventory Tools** in the menu bar
3. Select an option:
   - **📊 Open Control Panel** - Access sidebar UI
   - **📝 Open Inventory Form** - Fill out new item
   - **📦 Box Reports** - Generate reports
   - **🔧 Utilities** - Data processing tools

### Via Sidebar

1. Click **Open Control Panel** from menu
2. Use the sidebar to:
   - Generate single box reports
   - Email box summaries
   - Generate all box reports
   - View recent images

## Sharing the Form

Share your inventory form with others:

**Direct Link:**
Get the form URL from `getFormUrl()` function and share via:
- Email
- Text message
- QR code
- Embedded on website

**QR Code:**
Generate a QR code for the form URL and print it for easy scanning.

## Deployment

### Create a Versioned Deployment

```bash
clasp deploy -d "v1.0 - Initial Release"
```

### List Deployments

```bash
clasp deployments
```

### Update Deployment

```bash
clasp push
clasp deploy -d "v1.1 - Added QR codes" -i <deployment-id>
```

## Troubleshooting

### "Form Responses 1" sheet not found
- Ensure your Google Form is linked to the sheet
- Verify the sheet name is exactly "Form Responses 1"

### Images not displaying
- Check image URLs are publicly accessible
- Verify "Image Upload" column exists

### QR code generation fails
- Enable Google Chart API (it's free and public)
- Check internet connectivity
- Verify "qr code" folder exists in Drive

### clasp push fails
- Run `clasp login` to re-authenticate
- Check `.claspignore` isn't excluding necessary files
- Verify `.clasp.json` has correct script ID

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Version History

- **v1.0** (2025-10-16)
  - Initial release
  - Basic inventory tracking
  - Report generation
  - Email summaries
  - QR code generation
  - Image management

## License

This project is provided as-is for personal use.

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Review `CLAUDE.md` for development guidance
3. Check Google Apps Script documentation

## Acknowledgments

- Built with Google Apps Script
- QR codes generated using Google Chart API
- UI design inspired by Google Material Design
- Developed with assistance from Claude Code

---

**Generated with Claude Code** - AI-assisted development
