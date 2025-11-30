/**
 * Google Apps Script for Make.com Webhook Integration
 * 
 * This script triggers the Make.com webhook when a Google Form is submitted.
 * It sends form data to the webhook for instant processing.
 * 
 * Webhook URL: https://hook.eu2.make.com/aiflluclc1hiyci8slmm4rex79tfvn5u
 * 
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet
 * 2. Go to Extensions → Apps Script
 * 3. Paste this script
 * 4. Update the column mappings in the getFormData() function
 * 5. Install the trigger (see instructions below)
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

// Your Make.com webhook URL
const WEBHOOK_URL = 'https://hook.eu2.make.com/aiflluclc1hiyci8slmm4rex79tfvn5u';

// Sheet name (update if your sheet has a different name)
const SHEET_NAME = 'Form Responses 1'; // Change to your actual sheet name

// ============================================================================
// COLUMN MAPPING
// ============================================================================
// Based on your Google Sheet structure:
//
// IMPORTANT: Google Forms ALWAYS adds Timestamp as Column A (automatic)
// So your actual structure is:
//
// Column A: Timestamp (auto-added by Google Forms - not in your header row)
// Column B: Name
// Column C: Email
// Column D: Phone
// Column E: Language
// Column F: Status
// Column G: Date Added
// Column H: Email 1 Date
// Column I: Email 2 Date
// Column J: Email 3 Date
// Column K: Notes
// Column L: Response (attendance: "Yes, I'll attend" / "No, I cannot attend")
// Column M: Response Date
// Column N: Comments
//
// If your sheet does NOT use Google Forms auto-timestamp, adjust accordingly

const COLUMN_MAPPING = {
  TIMESTAMP: 1,      // Column A - Timestamp (auto-added by Google Forms)
  NAME: 2,           // Column B - Name
  EMAIL: 3,          // Column C - Email
  PHONE: 4,          // Column D - Phone
  LANGUAGE: 5,       // Column E - Language
  STATUS: 6,         // Column F - Status
  DATE_ADDED: 7,     // Column G - Date Added
  RESPONSE: 12,      // Column L - Response (attendance: "Yes, I'll attend" / "No, I cannot attend")
  RESPONSE_DATE: 13, // Column M - Response Date
  COMMENTS: 14       // Column N - Comments
  // Email dates (H, I, J) are managed by Make.com, not form submissions
};

// ============================================================================
// MAIN FUNCTION - TRIGGERED ON FORM SUBMIT
// ============================================================================

/**
 * This function is automatically triggered when a Google Form is submitted.
 * It extracts the form response data and sends it to the Make.com webhook.
 */
function onFormSubmit(e) {
  try {
    // Get the spreadsheet
    const sheet = e.source.getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log('Error: Sheet "' + SHEET_NAME + '" not found');
      return;
    }
    
    // Get the last row (the newly submitted form response)
    const lastRow = sheet.getLastRow();
    const rowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Extract form data based on column mapping
    const formData = getFormData(rowData);
    
    // Validate required fields
    if (!formData.email || formData.email.trim() === '') {
      Logger.log('Error: Email is required but not found');
      return;
    }
    
    // Send data to Make.com webhook
    const response = sendToWebhook(formData);
    
    // Log the result
    if (response.success) {
      Logger.log('Successfully sent data to webhook for: ' + formData.email);
    } else {
      Logger.log('Error sending to webhook: ' + response.error);
    }
    
  } catch (error) {
    Logger.log('Error in onFormSubmit: ' + error.toString());
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Extracts and structures form data from the row array
 * Update this function to match your form's column structure
 */
function getFormData(rowData) {
  // Extract data based on your sheet structure
  // Column indices are 0-based in the array (0 = Column A, 1 = Column B, etc.)
  
  // Get timestamp (use Date Added if Timestamp column doesn't exist)
  let timestamp = rowData[COLUMN_MAPPING.TIMESTAMP - 1];
  if (!timestamp && rowData[COLUMN_MAPPING.DATE_ADDED - 1]) {
    timestamp = rowData[COLUMN_MAPPING.DATE_ADDED - 1];
  }
  if (!timestamp) {
    timestamp = new Date();
  }
  // Convert to ISO string if it's a Date object
  if (timestamp instanceof Date) {
    timestamp = timestamp.toISOString();
  }
  
  const data = {
    timestamp: timestamp,
    name: rowData[COLUMN_MAPPING.NAME - 1] || '',
    email: rowData[COLUMN_MAPPING.EMAIL - 1] || '',
    phone: rowData[COLUMN_MAPPING.PHONE - 1] || '',
    language: rowData[COLUMN_MAPPING.LANGUAGE - 1] || 'en',
    status: rowData[COLUMN_MAPPING.STATUS - 1] || '',
    response: rowData[COLUMN_MAPPING.RESPONSE - 1] || '', // Attendance response
    responseDate: rowData[COLUMN_MAPPING.RESPONSE_DATE - 1] || '',
    comments: rowData[COLUMN_MAPPING.COMMENTS - 1] || ''
  };
  
  // Clean up email (trim whitespace, convert to lowercase)
  if (data.email) {
    data.email = data.email.toString().trim().toLowerCase();
  }
  
  // Clean up name (trim whitespace, capitalize)
  if (data.name) {
    data.name = data.name.toString().trim();
  }
  
  // Clean up phone (trim whitespace)
  if (data.phone) {
    data.phone = data.phone.toString().trim();
  }
  
  // Map response to attendance for backward compatibility
  // If you use "Response" column, it maps to "attendance" for Make.com
  data.attendance = data.response || '';
  
  return data;
}

/**
 * Sends data to the Make.com webhook
 */
function sendToWebhook(data) {
  try {
    const payload = {
      timestamp: data.timestamp,
      name: data.name,
      email: data.email,
      phone: data.phone || '',
      language: data.language || 'en',
      status: data.status || '',
      attendance: data.attendance || data.response || '', // Use response as attendance
      response: data.response || '',
      responseDate: data.responseDate || '',
      comments: data.comments || ''
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true // Don't throw errors, return them instead
    };
    
    const response = UrlFetchApp.fetch(WEBHOOK_URL, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode >= 200 && responseCode < 300) {
      return {
        success: true,
        statusCode: responseCode,
        response: responseText
      };
    } else {
      return {
        success: false,
        statusCode: responseCode,
        error: responseText
      };
    }
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Diagnostic function - Run this to see your sheet structure
 * This will help you identify which columns contain which data
 */
function diagnoseSheetStructure() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log('Error: Sheet "' + SHEET_NAME + '" not found');
      Logger.log('Available sheets: ' + SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName()).join(', '));
      return;
    }
    
    Logger.log('=== SHEET STRUCTURE DIAGNOSIS ===');
    Logger.log('Sheet Name: ' + SHEET_NAME);
    Logger.log('Total Columns: ' + sheet.getLastColumn());
    Logger.log('Total Rows: ' + sheet.getLastRow());
    Logger.log('');
    
    // Get header row
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('=== HEADER ROW (Row 1) ===');
    headerRow.forEach((header, index) => {
      const columnLetter = String.fromCharCode(65 + index); // Convert to A, B, C...
      Logger.log('Column ' + columnLetter + ' (' + (index + 1) + '): "' + (header || '(empty)') + '"');
    });
    Logger.log('');
    
    // Get first data row (if exists)
    if (sheet.getLastRow() > 1) {
      const firstDataRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      Logger.log('=== FIRST DATA ROW (Row 2) ===');
      firstDataRow.forEach((value, index) => {
        const columnLetter = String.fromCharCode(65 + index);
        Logger.log('Column ' + columnLetter + ' (' + (index + 1) + '): "' + (value || '(empty)') + '"');
      });
      Logger.log('');
    }
    
    // Get last data row
    if (sheet.getLastRow() > 1) {
      const lastRowNum = sheet.getLastRow();
      const lastDataRow = sheet.getRange(lastRowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
      Logger.log('=== LAST DATA ROW (Row ' + lastRowNum + ') ===');
      lastDataRow.forEach((value, index) => {
        const columnLetter = String.fromCharCode(65 + index);
        Logger.log('Column ' + columnLetter + ' (' + (index + 1) + '): "' + (value || '(empty)') + '"');
      });
      Logger.log('');
    }
    
    Logger.log('=== COLUMN MAPPING ANALYSIS ===');
    Logger.log('Based on current COLUMN_MAPPING settings:');
    Logger.log('Timestamp: Column ' + getColumnLetter(COLUMN_MAPPING.TIMESTAMP) + ' (' + COLUMN_MAPPING.TIMESTAMP + ')');
    Logger.log('Name: Column ' + getColumnLetter(COLUMN_MAPPING.NAME) + ' (' + COLUMN_MAPPING.NAME + ') = "' + (headerRow[COLUMN_MAPPING.NAME - 1] || '') + '"');
    Logger.log('Email: Column ' + getColumnLetter(COLUMN_MAPPING.EMAIL) + ' (' + COLUMN_MAPPING.EMAIL + ') = "' + (headerRow[COLUMN_MAPPING.EMAIL - 1] || '') + '"');
    Logger.log('Phone: Column ' + getColumnLetter(COLUMN_MAPPING.PHONE) + ' (' + COLUMN_MAPPING.PHONE + ') = "' + (headerRow[COLUMN_MAPPING.PHONE - 1] || '') + '"');
    Logger.log('Language: Column ' + getColumnLetter(COLUMN_MAPPING.LANGUAGE) + ' (' + COLUMN_MAPPING.LANGUAGE + ') = "' + (headerRow[COLUMN_MAPPING.LANGUAGE - 1] || '') + '"');
    Logger.log('Response: Column ' + getColumnLetter(COLUMN_MAPPING.RESPONSE) + ' (' + COLUMN_MAPPING.RESPONSE + ') = "' + (headerRow[COLUMN_MAPPING.RESPONSE - 1] || '') + '"');
    Logger.log('Comments: Column ' + getColumnLetter(COLUMN_MAPPING.COMMENTS) + ' (' + COLUMN_MAPPING.COMMENTS + ') = "' + (headerRow[COLUMN_MAPPING.COMMENTS - 1] || '') + '"');
    
    Logger.log('');
    Logger.log('=== RECOMMENDATIONS ===');
    if (!headerRow[COLUMN_MAPPING.NAME - 1] || !headerRow[COLUMN_MAPPING.NAME - 1].toString().toLowerCase().includes('name')) {
      Logger.log('⚠ WARNING: Column ' + COLUMN_MAPPING.NAME + ' does not appear to contain "Name"');
    }
    if (!headerRow[COLUMN_MAPPING.EMAIL - 1] || !headerRow[COLUMN_MAPPING.EMAIL - 1].toString().toLowerCase().includes('email')) {
      Logger.log('⚠ WARNING: Column ' + COLUMN_MAPPING.EMAIL + ' does not appear to contain "Email"');
    }
    
  } catch (error) {
    Logger.log('Error in diagnoseSheetStructure: ' + error.toString());
  }
}

/**
 * Helper function to convert column number to letter (1 = A, 2 = B, etc.)
 */
function getColumnLetter(columnNumber) {
  let letter = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}

/**
 * Manual test function - Use this to test the webhook without submitting a form
 * Run this function from the Apps Script editor to test
 */
function testWebhook() {
  // Create test data matching your sheet structure
  const testData = {
    timestamp: new Date().toISOString(),
    name: 'Test User',
    email: 'test@example.com',
    phone: '+968 1234 5678',
    language: 'en',
    status: '',
    attendance: 'Yes, I\'ll attend',
    response: 'Yes, I\'ll attend',
    responseDate: '',
    comments: 'This is a test submission'
  };
  
  Logger.log('Testing webhook with data: ' + JSON.stringify(testData));
  
  const response = sendToWebhook(testData);
  
  if (response.success) {
    Logger.log('Test successful! Status: ' + response.statusCode);
    Logger.log('Response: ' + response.response);
  } else {
    Logger.log('Test failed! Error: ' + response.error);
  }
}

// ============================================================================
// INSTALLATION INSTRUCTIONS
// ============================================================================

/**
 * TO INSTALL THE TRIGGER:
 * 
 * 1. Save this script in Apps Script editor
 * 2. Click on the clock icon (Triggers) in the left sidebar
 * 3. Click "+ Add Trigger" button
 * 4. Configure the trigger:
 *    - Choose which function to run: onFormSubmit
 *    - Choose which event source: From form
 *    - Select event type: On form submit
 *    - Failure notification settings: Choose your preference
 * 5. Click "Save"
 * 
 * TO UPDATE COLUMN MAPPING:
 * 
 * 1. Open your Google Sheet with form responses
 * 2. Check which columns contain which data:
 *    - Column A = Timestamp (usually automatic)
 *    - Column B = ? (check your form)
 *    - Column C = ? (check your form)
 *    - etc.
 * 3. Update the COLUMN_MAPPING object above with correct column numbers
 * 4. Update the getFormData() function if you have additional fields
 * 
 * TO TEST:
 * 
 * 1. Save the script
 * 2. Run the testWebhook() function from the editor
 * 3. Check the logs (View → Logs) to see the result
 * 4. Submit a test form response
 * 5. Check Make.com scenario execution
 * 
 * TROUBLESHOOTING:
 * 
 * - If webhook not triggering: Check trigger is installed correctly
 * - If wrong data: Verify column mappings match your form
 * - If errors: Check View → Logs for error messages
 * - If Make.com not receiving: Test webhook URL manually first
 */

