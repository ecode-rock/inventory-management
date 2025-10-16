/**
 * Inventory Management System - Main Script
 * Fixed version with working form opener and complete features
 */

// ========================================
// CONFIGURATION
// ========================================

/**
 * TEMPLATE DOCUMENT ID
 * To use a template:
 * 1. Create a Google Doc with your desired formatting
 * 2. Add placeholders: {{BOX_NUMBER}}, {{DATE}}, {{TOTAL_ITEMS}}, {{LOCATION}}, {{QR_CODE}}
 * 3. Copy the document ID from the URL (the part after /d/)
 * 4. Paste it below
 *
 * Leave empty ("") to create documents from scratch (current behavior)
 */
const TEMPLATE_DOC_ID = "";

// ========================================
// MAIN FUNCTIONS
// ========================================

// This function runs automatically when the spreadsheet is opened
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Create custom menu
  ui.createMenu('🗃️ Inventory Tools')
    .addItem('📊 Open Control Panel', 'showSidebar')
    .addSeparator()
    .addItem('📝 Open Inventory Form', 'openFormDirect')
    .addSeparator()
    .addSubMenu(ui.createMenu('📦 Box Reports')
      .addItem('Generate Single Box Report', 'promptForBoxReport')
      .addItem('Generate All Box Reports', 'generateAllBoxReports'))
    .addSubMenu(ui.createMenu('🎨 Templates')
      .addItem('Create Sample Template', 'createSampleTemplate')
      .addItem('About Templates', 'showTemplateHelp'))
    .addSubMenu(ui.createMenu('🔧 Utilities')
      .addItem('Process New Images', 'processImages')
      .addItem('Check Data Integrity', 'checkData'))
    .addSeparator()
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Inventory Tools menu has been added to the menu bar.',
    '✅ Ready',
    3
  );
}

// Show sidebar
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Inventory Control Panel')
    .setWidth(300);

  SpreadsheetApp.getUi().showSidebar(html);
}

// FIXED: Simple function that returns the form URL for the sidebar to use
function getFormUrl() {
  return 'https://docs.google.com/forms/d/e/1FAIpQLSfeIpxOdfX7t6H1ywhdJlWJITRqXR3q2XuEj1S2km4jfg1YwA/viewform?usp=header';
}

// Direct form opener for menu
function openFormDirect() {
  const formUrl = getFormUrl();
  const html = HtmlService.createHtmlOutput(
    '<script>window.open("' + formUrl + '", "_blank");google.script.host.close();</script>'
  ).setWidth(100).setHeight(50);

  SpreadsheetApp.getUi().showModalDialog(html, 'Opening form...');
}

// Prompt for box number and generate report (for menu)
function promptForBoxReport() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Generate Box Report', 'Enter box number:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const boxNumber = response.getResponseText().trim();
    if (boxNumber) {
      try {
        const result = generateSingleBoxReport(boxNumber);
        ui.alert('Success', 'Report created! Opening in new tab...', ui.ButtonSet.OK);

        // Open the created document
        const html = HtmlService.createHtmlOutput(
          '<script>window.open("' + result.documentUrl + '", "_blank");google.script.host.close();</script>'
        );
        ui.showModalDialog(html, 'Opening report...');
      } catch (error) {
        ui.alert('Error', error.toString(), ui.ButtonSet.OK);
      }
    }
  }
}

// Generate a report for a single box
function generateSingleBoxReport(boxNumber) {
  if (!boxNumber) {
    throw new Error('Box number is required');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Form Responses 1');

    if (!sheet) {
      throw new Error('Form Responses sheet not found');
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Find column indices
    const colIndices = {
      timestamp: headers.indexOf('Timestamp'),
      itemName: headers.indexOf('Item name'),
      description: headers.indexOf('Item description'),
      category: headers.indexOf('Item Category'),
      quantity: headers.indexOf('Quantity'),
      room: headers.indexOf('Room location'),
      boxNumber: headers.indexOf('Storage box Number'),
      actionPlan: headers.indexOf('Action Plan'),
      imageUrl: headers.indexOf('Image Upload')
    };

    // Filter data for this box
    const boxItems = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[colIndices.boxNumber] &&
          row[colIndices.boxNumber].toString().toUpperCase() === boxNumber.toUpperCase()) {
        boxItems.push({
          itemName: row[colIndices.itemName] || 'Unnamed Item',
          description: row[colIndices.description] || 'No description',
          category: row[colIndices.category] || 'Uncategorized',
          quantity: row[colIndices.quantity] || 1,
          room: row[colIndices.room] || 'Unknown',
          actionPlan: row[colIndices.actionPlan] || 'Keep',
          imageUrl: row[colIndices.imageUrl] || '',
          timestamp: row[colIndices.timestamp] || new Date()
        });
      }
    }

    if (boxItems.length === 0) {
      throw new Error('No items found for box ' + boxNumber);
    }

    const report = createBoxDocument(boxNumber, boxItems);

    return {
      success: true,
      message: 'Report created for box ' + boxNumber + ' with ' + boxItems.length + ' items',
      documentUrl: report.url,
      documentId: report.id
    };

  } catch (error) {
    console.error('Error generating report:', error);
    throw error;
  }
}

// Generate QR code for a URL and save to Drive
function generateQRCode(url, boxNumber) {
  try {
    // Get or create QR code folder
    const qrFolderId = getOrCreateFolder('qr code');

    // Use QuickChart.io API to generate QR code (replacement for deprecated Google Chart API)
    const qrSize = 300;
    const qrUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(url) + '&size=' + qrSize;

    // Fetch the QR code image
    const response = UrlFetchApp.fetch(qrUrl);
    const blob = response.getBlob();

    // Set proper content type for PNG image
    blob.setContentType('image/png');

    // Set filename
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
    const filename = 'QR_Box_' + boxNumber + '_' + dateStr + '.png';
    blob.setName(filename);

    // Save to Drive folder
    const qrFolder = DriveApp.getFolderById(qrFolderId);
    const file = qrFolder.createFile(blob);

    return {
      id: file.getId(),
      url: file.getUrl(),
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId(),
      qrApiUrl: qrUrl  // Return the original QuickChart URL for direct fetching
    };

  } catch (error) {
    console.error('Error generating QR code:', error);
    // Return null if QR code generation fails, don't break the whole report
    return null;
  }
}

// Create a formatted Google Doc for the box
function createBoxDocument(boxNumber, items) {
  try {
    const folderId = getOrCreateFolder('Box Reports');

    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const docName = 'Box ' + boxNumber + ' - Report (' + dateStr + ')';

    let doc, body;

    // Check if using template or creating from scratch
    if (TEMPLATE_DOC_ID && TEMPLATE_DOC_ID.trim() !== "") {
      // TEMPLATE MODE: Copy template and modify
      const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
      const newFile = templateFile.makeCopy(docName);
      doc = DocumentApp.openById(newFile.getId());
      body = doc.getBody();
    } else {
      // FROM SCRATCH MODE: Create new document
      doc = DocumentApp.create(docName);
      body = doc.getBody();
      body.clear();
    }

    // Generate QR code for this document
    const docUrl = doc.getUrl();
    const qrCode = generateQRCode(docUrl, boxNumber);

    if (TEMPLATE_DOC_ID && TEMPLATE_DOC_ID.trim() !== "") {
      // TEMPLATE MODE: Replace placeholders
      body.replaceText('{{BOX_NUMBER}}', boxNumber);
      body.replaceText('{{DATE}}', dateStr);
      body.replaceText('{{TOTAL_ITEMS}}', items.length.toString());
      body.replaceText('{{LOCATION}}', items[0].room);

      // Find and replace QR_CODE placeholder with actual QR code image
      if (qrCode) {
        const searchResult = body.findText('{{QR_CODE}}');
        if (searchResult) {
          const element = searchResult.getElement();
          const parent = element.getParent();

          try {
            // Insert QR code image after the placeholder text - fetch directly from QuickChart
            const qrImageBlob = UrlFetchApp.fetch(qrCode.qrApiUrl).getBlob();
            qrImageBlob.setContentType('image/png');
            const qrImage = parent.asParagraph().appendInlineImage(qrImageBlob);
            qrImage.setWidth(200);
            qrImage.setHeight(200);

            // Remove the placeholder text
            element.asText().setText('');
          } catch (error) {
            console.error('Error inserting QR code image:', error);
            body.replaceText('{{QR_CODE}}', 'QR Code Link: ' + qrCode.url);
          }
        }
      } else {
        body.replaceText('{{QR_CODE}}', '(QR code generation failed)');
      }

    } else {
      // FROM SCRATCH MODE: Build document structure
      // Add header
      const header = body.appendParagraph('📦 Storage Box ' + boxNumber + ' - Inventory Report');
      header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

      // Add metadata
      body.appendParagraph('Generated: ' + dateStr);
      body.appendParagraph('Total Items: ' + items.length);
      body.appendParagraph('Location: ' + items[0].room);

      // Add QR code
      if (qrCode) {
        body.appendParagraph('');
        const qrPara = body.appendParagraph('QR Code (scan to open this report):');
        qrPara.setBold(true);

        try {
          // Fetch and insert QR code image - fetch directly from QuickChart
          const qrImageBlob = UrlFetchApp.fetch(qrCode.qrApiUrl).getBlob();
          qrImageBlob.setContentType('image/png');
          const qrImage = body.appendImage(qrImageBlob);
          qrImage.setWidth(200);
          qrImage.setHeight(200);
        } catch (error) {
          console.error('Error inserting QR code image:', error);
          body.appendParagraph('QR Code Link: ' + qrCode.url);
        }
      }

      body.appendHorizontalRule();
    }

    // Create summary statistics
    const stats = getBoxStatistics(items);
    const statsSection = body.appendParagraph('📊 Summary Statistics');
    statsSection.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    statsSection.setBold(true);

    // Add statistics table
    const statsTable = [
      ['Category', 'Count'],
      ...Object.entries(stats.byCategory).map(([cat, count]) => [cat, count.toString()]),
      ['', ''],
      ['Action Plan', 'Count'],
      ...Object.entries(stats.byAction).map(([action, count]) => [action, count.toString()])
    ];

    const table = body.appendTable(statsTable);
    table.getRow(0).editAsText().setBold(true);
    const actionPlanRow = statsTable.findIndex(row => row[0] === 'Action Plan');
    if (actionPlanRow > 0) {
      table.getRow(actionPlanRow).editAsText().setBold(true);
    }

    body.appendHorizontalRule();

    // Add detailed item list
    const itemsSection = body.appendParagraph('📋 Detailed Item List');
    itemsSection.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    itemsSection.setBold(true);

    // Add each item
    items.forEach((item, index) => {
      const itemHeader = body.appendParagraph((index + 1) + '. ' + item.itemName);
      itemHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      itemHeader.setBold(true);

      body.appendListItem('Category: ' + item.category).setGlyphType(DocumentApp.GlyphType.BULLET);
      body.appendListItem('Quantity: ' + item.quantity).setGlyphType(DocumentApp.GlyphType.BULLET);
      body.appendListItem('Description: ' + item.description).setGlyphType(DocumentApp.GlyphType.BULLET);
      body.appendListItem('Action Plan: ' + item.actionPlan).setGlyphType(DocumentApp.GlyphType.BULLET);

      if (item.imageUrl && item.imageUrl.length > 0) {
        const imageText = body.appendListItem('Image: ').setGlyphType(DocumentApp.GlyphType.BULLET);
        imageText.appendText(item.imageUrl).setLinkUrl(item.imageUrl);
      }

      if (index < items.length - 1) {
        body.appendParagraph('');
      }
    });

    // Collect all image links
    const imageLinks = items
      .filter(item => item.imageUrl && item.imageUrl.length > 0)
      .map(item => ({
        itemName: item.itemName,
        imageUrl: item.imageUrl
      }));

    // Add image links section at the bottom if there are any images
    if (imageLinks.length > 0) {
      body.appendHorizontalRule();
      const imagesSection = body.appendParagraph('📷 Image Links');
      imagesSection.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      imagesSection.setBold(true);

      body.appendParagraph('Click on any link below to view the full-size image:');
      body.appendParagraph('');

      imageLinks.forEach((link, index) => {
        const linkPara = body.appendParagraph((index + 1) + '. ' + link.itemName + ': ');
        linkPara.appendText(link.imageUrl).setLinkUrl(link.imageUrl);
      });
    }

    // Add footer
    body.appendHorizontalRule();
    const footer = body.appendParagraph('--- End of Report ---');
    footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    footer.setItalic(true);
    footer.setForegroundColor('#666666');

    // Move to folder
    const file = DriveApp.getFileById(doc.getId());
    DriveApp.getFolderById(folderId).addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    return {
      id: doc.getId(),
      url: doc.getUrl(),
      qrCodeUrl: qrCode ? qrCode.url : null
    };

  } catch (error) {
    console.error('Error creating document:', error);
    throw error;
  }
}

// Calculate statistics for the box
function getBoxStatistics(items) {
  const stats = {
    byCategory: {},
    byAction: {},
    totalQuantity: 0
  };

  items.forEach(item => {
    const category = item.category || 'Uncategorized';
    stats.byCategory[category] = (stats.byCategory[category] || 0) + 1;

    const action = item.actionPlan || 'Keep';
    stats.byAction[action] = (stats.byAction[action] || 0) + 1;

    stats.totalQuantity += parseInt(item.quantity) || 1;
  });

  return stats;
}

// Get or create a folder in Drive
function getOrCreateFolder(folderName) {
  try {
    const parentFolders = DriveApp.getFoldersByName('Inventory System');
    let parentFolder;

    if (parentFolders.hasNext()) {
      parentFolder = parentFolders.next();
    } else {
      parentFolder = DriveApp.createFolder('Inventory System');
    }

    const subfolders = parentFolder.getFoldersByName(folderName);

    if (subfolders.hasNext()) {
      return subfolders.next().getId();
    } else {
      const newFolder = parentFolder.createFolder(folderName);
      return newFolder.getId();
    }

  } catch (error) {
    console.error('Error with folder:', error);
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      return folders.next().getId();
    } else {
      return DriveApp.createFolder(folderName).getId();
    }
  }
}

// Generate reports for all boxes
function generateAllBoxReports() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Form Responses 1');

    if (!sheet) {
      throw new Error('Form Responses sheet not found');
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const boxColIndex = headers.indexOf('Storage box Number');

    if (boxColIndex === -1) {
      throw new Error('Storage box Number column not found');
    }

    const uniqueBoxes = new Set();
    for (let i = 1; i < data.length; i++) {
      const boxNumber = data[i][boxColIndex];
      if (boxNumber && boxNumber.toString().trim() !== '') {
        uniqueBoxes.add(boxNumber.toString().trim().toUpperCase());
      }
    }

    if (uniqueBoxes.size === 0) {
      throw new Error('No box numbers found in the sheet');
    }

    const results = [];
    uniqueBoxes.forEach(boxNumber => {
      try {
        const result = generateSingleBoxReport(boxNumber);
        results.push({
          box: boxNumber,
          status: 'success',
          url: result.documentUrl
        });
      } catch (error) {
        results.push({
          box: boxNumber,
          status: 'error',
          error: error.toString()
        });
      }
    });

    const successCount = results.filter(r => r.status === 'success').length;
    const message = 'Generated ' + successCount + ' of ' + results.length + ' reports successfully';

    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Bulk Report Generation', 10);

    return results;

  } catch (error) {
    console.error('Error in bulk generation:', error);
    throw error;
  }
}

// NEW: Get recent images for thumbnails
function getRecentImages(limit) {
  limit = limit || 12;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Form Responses 1');

    if (!sheet) {
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const imageColIndex = headers.indexOf('Image Upload');
    const itemNameIndex = headers.indexOf('Item name');
    const boxNumIndex = headers.indexOf('Storage box Number');

    if (imageColIndex === -1) {
      return [];
    }

    const images = [];

    // Start from most recent (bottom of sheet) and work backwards
    for (let i = data.length - 1; i >= 1 && images.length < limit; i--) {
      const row = data[i];
      const imageUrl = row[imageColIndex];

      if (imageUrl && imageUrl.toString().trim() !== '') {
        images.push({
          url: imageUrl.toString(),
          itemName: row[itemNameIndex] || 'Unknown Item',
          boxNumber: row[boxNumIndex] || 'Unknown'
        });
      }
    }

    return images;

  } catch (error) {
    console.error('Error getting images:', error);
    return [];
  }
}

// NEW: Email box summary
function emailBoxSummary(boxNumber, recipientEmail) {
  if (!boxNumber) {
    throw new Error('Box number is required');
  }

  if (!recipientEmail || recipientEmail.trim() === '') {
    recipientEmail = Session.getActiveUser().getEmail();
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Form Responses 1');

    if (!sheet) {
      throw new Error('Form Responses sheet not found');
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const colIndices = {
      itemName: headers.indexOf('Item name'),
      description: headers.indexOf('Item description'),
      category: headers.indexOf('Item Category'),
      quantity: headers.indexOf('Quantity'),
      room: headers.indexOf('Room location'),
      boxNumber: headers.indexOf('Storage box Number'),
      actionPlan: headers.indexOf('Action Plan')
    };

    // Filter data for this box
    const boxItems = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[colIndices.boxNumber] &&
          row[colIndices.boxNumber].toString().toUpperCase() === boxNumber.toUpperCase()) {
        boxItems.push({
          itemName: row[colIndices.itemName] || 'Unnamed Item',
          description: row[colIndices.description] || 'No description',
          category: row[colIndices.category] || 'Uncategorized',
          quantity: row[colIndices.quantity] || 1,
          room: row[colIndices.room] || 'Unknown',
          actionPlan: row[colIndices.actionPlan] || 'Keep'
        });
      }
    }

    if (boxItems.length === 0) {
      throw new Error('No items found for box ' + boxNumber);
    }

    // Create email body
    let emailBody = '<html><body style="font-family: Arial, sans-serif;">';
    emailBody += '<h2 style="color: #667eea;">📦 Storage Box ' + boxNumber + ' - Inventory Summary</h2>';
    emailBody += '<p><strong>Total Items:</strong> ' + boxItems.length + '</p>';
    emailBody += '<p><strong>Location:</strong> ' + boxItems[0].room + '</p>';
    emailBody += '<hr>';

    emailBody += '<h3>Items in this box:</h3>';
    emailBody += '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%;">';
    emailBody += '<tr style="background-color: #f0f0f0;">';
    emailBody += '<th>Item</th><th>Category</th><th>Qty</th><th>Action Plan</th></tr>';

    boxItems.forEach(item => {
      emailBody += '<tr>';
      emailBody += '<td><strong>' + item.itemName + '</strong><br><small>' + item.description + '</small></td>';
      emailBody += '<td>' + item.category + '</td>';
      emailBody += '<td>' + item.quantity + '</td>';
      emailBody += '<td>' + item.actionPlan + '</td>';
      emailBody += '</tr>';
    });

    emailBody += '</table>';
    emailBody += '<hr>';
    emailBody += '<p style="color: #666; font-size: 12px;">Generated by Inventory Management System on ' + new Date().toLocaleDateString() + '</p>';
    emailBody += '</body></html>';

    // Send email
    MailApp.sendEmail({
      to: recipientEmail,
      subject: '📦 Box ' + boxNumber + ' Inventory Summary',
      htmlBody: emailBody
    });

    return {
      success: true,
      message: 'Email sent to ' + recipientEmail
    };

  } catch (error) {
    console.error('Error sending email:', error);
    throw error;
  }
}

/**
 * Creates a sample template document to help users get started
 * This creates a properly formatted template with all placeholders
 */
function createSampleTemplate() {
  try {
    const doc = DocumentApp.create('Box Report Template - SAMPLE');
    const body = doc.getBody();
    body.clear();

    // Add header with placeholder
    const header = body.appendParagraph('📦 Storage Box {{BOX_NUMBER}} - Inventory Report');
    header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    // Add metadata section
    body.appendParagraph('Generated: {{DATE}}');
    body.appendParagraph('Total Items: {{TOTAL_ITEMS}}');
    body.appendParagraph('Location: {{LOCATION}}');
    body.appendParagraph('');

    // QR Code section
    const qrPara = body.appendParagraph('QR Code (scan to open this report):');
    qrPara.setBold(true);
    body.appendParagraph('{{QR_CODE}}');
    body.appendParagraph('');

    body.appendHorizontalRule();

    // Instructions
    const instructions = body.appendParagraph('Template Instructions:');
    instructions.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    instructions.setForegroundColor('#FF0000');

    const instructionsList = body.appendParagraph('1. Customize this template with your formatting, colors, fonts, and logo');
    instructionsList.setForegroundColor('#FF0000');
    body.appendParagraph('2. Keep the placeholders: {{BOX_NUMBER}}, {{DATE}}, {{TOTAL_ITEMS}}, {{LOCATION}}, {{QR_CODE}}').setForegroundColor('#FF0000');
    body.appendParagraph('3. Copy this document\'s ID from the URL').setForegroundColor('#FF0000');
    body.appendParagraph('4. Paste the ID in Code.gs where TEMPLATE_DOC_ID = ""').setForegroundColor('#FF0000');
    body.appendParagraph('5. Delete these red instructions before using').setForegroundColor('#FF0000');

    body.appendHorizontalRule();

    // Note about dynamic content
    const note = body.appendParagraph('Note: Statistics tables and item lists will be added automatically below this line');
    note.setItalic(true);
    note.setForegroundColor('#666666');

    // Save to folder
    const folderId = getOrCreateFolder('Templates');
    const file = DriveApp.getFileById(doc.getId());
    DriveApp.getFolderById(folderId).addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Template Created!',
      'Sample template created successfully!\n\n' +
      'Document ID: ' + doc.getId() + '\n\n' +
      'Click OK to open the template and customize it.',
      ui.ButtonSet.OK_CANCEL
    );

    if (result == ui.Button.OK) {
      const html = HtmlService.createHtmlOutput(
        '<script>window.open("' + doc.getUrl() + '", "_blank");google.script.host.close();</script>'
      );
      ui.showModalDialog(html, 'Opening template...');
    }

    return {
      id: doc.getId(),
      url: doc.getUrl()
    };

  } catch (error) {
    console.error('Error creating template:', error);
    throw error;
  }
}

// Placeholder functions
function processImages() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Feature coming soon: Process images', 'Info', 3);
}

function checkData() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Feature coming soon: Check data integrity', 'Info', 3);
}

function showTemplateHelp() {
  const currentTemplateId = TEMPLATE_DOC_ID && TEMPLATE_DOC_ID.trim() !== "" ? TEMPLATE_DOC_ID : "Not configured";
  const mode = TEMPLATE_DOC_ID && TEMPLATE_DOC_ID.trim() !== "" ? "Template Mode" : "From Scratch Mode";

  const htmlContent = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>🎨 Document Templates</h2>
      <p><strong>Current Mode:</strong> ${mode}</p>
      <p><strong>Template ID:</strong> <span style="font-family: monospace; font-size: 11px;">${currentTemplateId}</span></p>
      <hr>
      <h3>How to Use Templates:</h3>
      <ol>
        <li>Click <strong>Templates > Create Sample Template</strong></li>
        <li>Customize the template with your branding</li>
        <li>Copy the document ID from the URL</li>
        <li>Open <strong>Extensions > Apps Script</strong></li>
        <li>Edit Code.gs and paste ID in TEMPLATE_DOC_ID</li>
        <li>Save and test!</li>
      </ol>
      <hr>
      <h3>Available Placeholders:</h3>
      <ul style="font-family: monospace; font-size: 12px;">
        <li>{{BOX_NUMBER}}</li>
        <li>{{DATE}}</li>
        <li>{{TOTAL_ITEMS}}</li>
        <li>{{LOCATION}}</li>
        <li>{{QR_CODE}}</li>
      </ul>
      <hr>
      <p style="color: #666; font-size: 11px;"><strong>Tip:</strong> Leave TEMPLATE_DOC_ID empty to generate documents from scratch (default behavior)</p>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Template Help');
}

function showAbout() {
  const htmlContent = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>📦 Inventory Management System</h2>
      <p><strong>Version:</strong> 1.0</p>
      <p><strong>Purpose:</strong> Organize and track physical items in storage boxes</p>
      <hr>
      <h3>Features:</h3>
      <ul>
        <li>Track items by box location</li>
        <li>Generate box content reports</li>
        <li>Email summaries</li>
        <li>View recent item images</li>
        <li>QR codes for quick report access</li>
        <li>Custom document templates</li>
      </ul>
      <hr>
      <p style="color: #666; font-size: 12px;">Built with Google Apps Script</p>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(400)
    .setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'About Inventory System');
}
