/**
 * Google Apps Script for Make.com Webhook Integration
 * 
 * This script triggers the Make.com webhook when a Google Form is submitted
 * or when a new row is added to the sheet.
 * 
 * WEBHOOK PURPOSE: Email 1 - Invitation workflow ONLY
 * This webhook is specifically for triggering the "Email 1 - Invitation" 
 * scenario in Make.com when a new lead is added.
 * 
 * Webhook URL: https://hook.eu2.make.com/m0i3qbv1dwbh41jd9g9vpek57dqmglnn
 * 
 * NOTE: This is NOT used for:
 * - Email 2 (Reminder) - uses scheduled trigger
 * - Email 3 (Final Call) - uses scheduled trigger
 * - Response Form (Accept/Reject) - uses separate webhook
 * 
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet
 * 2. Go to Extensions → Apps Script
 * 3. Delete any old code and paste this entire script
 * 4. Update SHEET_NAME if needed (currently set to 'leads')
 * 5. Save the script
 * 6. Install the trigger (see instructions at bottom)
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

// Your Make.com webhook URL
const WEBHOOK_URL = 'https://hook.eu2.make.com/m0i3qbv1dwbh41jd9g9vpek57dqmglnn';

// Sheet name (update if your sheet has a different name)
const SHEET_NAME = 'leads'; // Change this to match your sheet name exactly

// ============================================================================
// COLUMN MAPPING
// ============================================================================
// Based on your sheet structure:
// Column A: Name
// Column B: Email
// Column C: Phone
// Column D: Language
// Column E: Status
// Column F: Date Added
// Column G: Email 1 Date
// Column H: Email 2 Date
// Column I: Email 3 Date
// Column J: Notes
// Column K: Response
// Column L: Response Date
// Column M: Comments
//
// NOTE: If Google Forms automatically added a "Timestamp" column as Column A,
//       then add 1 to all column numbers below (Name becomes 2, Email becomes 3, etc.)

const COLUMN_MAPPING = {
  NAME: 1,           // Column A - Name
  EMAIL: 2,          // Column B - Email
  PHONE: 3,          // Column C - Phone
  LANGUAGE: 4,       // Column D - Language (should be "en" or "ar")
  STATUS: 5,         // Column E - Status
  DATE_ADDED: 6,     // Column F - Date Added
  EMAIL_1_DATE: 7,   // Column G - Email 1 Date
  EMAIL_2_DATE: 8,   // Column H - Email 2 Date
  EMAIL_3_DATE: 9,   // Column I - Email 3 Date
  NOTES: 10,         // Column J - Notes
  RESPONSE: 11,      // Column K - Response (attendance)
  RESPONSE_DATE: 12, // Column L - Response Date
  COMMENTS: 13       // Column M - Comments
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
    // Validate event object
    if (!e || !e.source) {
      Logger.log('✗ Error: This function must be triggered by a form submission.');
      Logger.log('To test manually, use testLastRow() or testWebhook() functions instead.');
      return;
    }
    
    // Get the sheet
    const sheet = e.source.getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log('✗ Error: Sheet "' + SHEET_NAME + '" not found');
      Logger.log('Available sheets: ' + e.source.getSheets().map(s => s.getName()).join(', '));
      return;
    }
    
    // Get the last row (newly submitted form response)
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      Logger.log('✗ Error: No data rows found (only header row exists)');
      return;
    }
    
    // Get the last row data
    const rowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Extract and structure the form data
    const formData = getFormData(rowData);
    
    // Validate that email exists (required field)
    if (!formData.email || formData.email.trim() === '') {
      Logger.log('✗ Error: Email is required but not found in row ' + lastRow);
      return;
    }
    
    // Send data to webhook
    Logger.log('→ Sending data to webhook for: ' + formData.email);
    const response = sendToWebhook(formData);
    
    if (response.success) {
      Logger.log('✓ Successfully sent data to webhook!');
      Logger.log('✓ Status code: ' + response.statusCode);
      Logger.log('✓ Check Make.com for the execution.');
    } else {
      Logger.log('✗ Error sending to webhook');
      Logger.log('✗ Status code: ' + response.statusCode);
      Logger.log('✗ Error: ' + (response.error || 'Unknown error'));
    }
    
  } catch (error) {
    Logger.log('✗ Error in onFormSubmit: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Extracts and structures form data from the row array
 * Row arrays in Google Sheets are 0-indexed, so subtract 1 from column mapping
 */
function getFormData(rowData) {
  // Get timestamp from Date Added column, or use current time
  let timestamp = rowData[COLUMN_MAPPING.DATE_ADDED - 1];
  
  // Convert to ISO string if it's a Date object
  if (timestamp instanceof Date) {
    timestamp = timestamp.toISOString();
  } else if (timestamp && typeof timestamp === 'string' && timestamp.trim() !== '') {
    // Keep as is if it's already a string
  } else {
    // Use current time if no date found
    timestamp = new Date().toISOString();
  }
  
  // Extract all fields from the row (convert to strings and trim)
  const data = {
    timestamp: timestamp,
    name: (rowData[COLUMN_MAPPING.NAME - 1] || '').toString().trim(),
    email: (rowData[COLUMN_MAPPING.EMAIL - 1] || '').toString().trim(),
    phone: (rowData[COLUMN_MAPPING.PHONE - 1] || '').toString().trim(),
    language: (rowData[COLUMN_MAPPING.LANGUAGE - 1] || 'en').toString().trim().toLowerCase(),
    status: (rowData[COLUMN_MAPPING.STATUS - 1] || '').toString().trim(),
    dateAdded: rowData[COLUMN_MAPPING.DATE_ADDED - 1] || '',
    response: (rowData[COLUMN_MAPPING.RESPONSE - 1] || '').toString().trim(),
    responseDate: (rowData[COLUMN_MAPPING.RESPONSE_DATE - 1] || '').toString().trim(),
    comments: (rowData[COLUMN_MAPPING.COMMENTS - 1] || '').toString().trim()
  };
  
  // Clean up email (convert to lowercase, trim whitespace)
  if (data.email) {
    data.email = data.email.toLowerCase().trim();
  }
  
  // Map response to attendance for backward compatibility with Make.com scenarios
  data.attendance = data.response || '';
  
  return data;
}

/**
 * Sends data to the Make.com webhook
 * Returns success/failure status
 */
function sendToWebhook(data) {
  try {
    // Structure the payload exactly as Make.com expects
    const payload = {
      timestamp: data.timestamp,
      name: data.name,
      email: data.email,
      phone: data.phone || '',
      language: data.language || 'en',
      status: data.status || '',
      attendance: data.attendance || data.response || '',
      response: data.response || '',
      responseDate: data.responseDate || '',
      comments: data.comments || ''
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true // Return errors instead of throwing
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

// ============================================================================
// TEST FUNCTIONS
// ============================================================================

/**
 * Diagnostic function - Shows your sheet structure
 * Run this to verify the column mapping is correct
 */
function diagnoseSheetStructure() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      Logger.log('✗ Error: Sheet "' + SHEET_NAME + '" not found');
      Logger.log('Available sheets: ' + spreadsheet.getSheets().map(s => s.getName()).join(', '));
      return;
    }
    
    Logger.log('=== SHEET STRUCTURE DIAGNOSIS ===');
    Logger.log('Sheet Name: ' + SHEET_NAME);
    Logger.log('Total Columns: ' + sheet.getLastColumn());
    Logger.log('Total Rows: ' + sheet.getLastRow());
    Logger.log('');
    
    // Get header row
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    Logger.log('=== HEADER ROW ===');
    headerRow.forEach((header, index) => {
      const columnNum = index + 1;
      let columnLetter = '';
      if (columnNum <= 26) {
        columnLetter = String.fromCharCode(64 + columnNum);
      } else {
        columnLetter = String.fromCharCode(64 + Math.floor((columnNum - 1) / 26)) + 
                       String.fromCharCode(64 + ((columnNum - 1) % 26) + 1);
      }
      Logger.log('Column ' + columnLetter + ' (' + columnNum + '): "' + (header || '(empty)') + '"');
    });
    
    Logger.log('');
    Logger.log('=== COLUMN MAPPING VERIFICATION ===');
    Logger.log('Name column ' + COLUMN_MAPPING.NAME + ' → "' + (headerRow[COLUMN_MAPPING.NAME - 1] || 'NOT FOUND') + '"');
    Logger.log('Email column ' + COLUMN_MAPPING.EMAIL + ' → "' + (headerRow[COLUMN_MAPPING.EMAIL - 1] || 'NOT FOUND') + '"');
    Logger.log('Language column ' + COLUMN_MAPPING.LANGUAGE + ' → "' + (headerRow[COLUMN_MAPPING.LANGUAGE - 1] || 'NOT FOUND') + '"');
    Logger.log('Response column ' + COLUMN_MAPPING.RESPONSE + ' → "' + (headerRow[COLUMN_MAPPING.RESPONSE - 1] || 'NOT FOUND') + '"');
    
  } catch (error) {
    Logger.log('✗ Error: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
  }
}

/**
 * Test function - Sends test data to webhook
 * Use this to test the webhook connection without submitting a form
 */
function testWebhook() {
  Logger.log('=== TESTING WEBHOOK CONNECTION ===');
  
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
    comments: 'This is a test submission from Apps Script'
  };
  
  Logger.log('Sending test data to webhook...');
  Logger.log('Test email: ' + testData.email);
  
  const response = sendToWebhook(testData);
  
  if (response.success) {
    Logger.log('✓ Test successful!');
    Logger.log('✓ Status code: ' + response.statusCode);
    Logger.log('✓ Check Make.com for the execution.');
  } else {
    Logger.log('✗ Test failed!');
    Logger.log('✗ Status code: ' + response.statusCode);
    Logger.log('✗ Error: ' + (response.error || 'Unknown error'));
  }
}

/**
 * Test function - Processes the last row from your sheet
 * Use this to test with actual data from your sheet
 */
function testLastRow() {
  try {
    Logger.log('=== TESTING WITH LAST ROW DATA ===');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      Logger.log('✗ Error: Sheet "' + SHEET_NAME + '" not found');
      Logger.log('Available sheets: ' + spreadsheet.getSheets().map(s => s.getName()).join(', '));
      return;
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      Logger.log('✗ Error: No data rows found. Sheet must have at least one row of data.');
      return;
    }
    
    Logger.log('Processing last row (' + lastRow + ') from sheet...');
    
    // Get the last row data
    const rowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Extract form data
    const formData = getFormData(rowData);
    
    Logger.log('');
    Logger.log('=== EXTRACTED DATA ===');
    Logger.log('Name: "' + formData.name + '"');
    Logger.log('Email: "' + formData.email + '"');
    Logger.log('Phone: "' + formData.phone + '"');
    Logger.log('Language: "' + formData.language + '"');
    Logger.log('Status: "' + formData.status + '"');
    Logger.log('Response: "' + formData.response + '"');
    Logger.log('Comments: "' + formData.comments + '"');
    Logger.log('');
    
    if (!formData.email || formData.email.trim() === '') {
      Logger.log('✗ Error: Email is required but not found in row ' + lastRow);
      return;
    }
    
    Logger.log('Sending data to webhook...');
    const response = sendToWebhook(formData);
    
    if (response.success) {
      Logger.log('✓ Successfully sent data to webhook!');
      Logger.log('✓ Status code: ' + response.statusCode);
      Logger.log('✓ Check Make.com for the execution.');
    } else {
      Logger.log('✗ Failed to send data to webhook!');
      Logger.log('✗ Status code: ' + response.statusCode);
      Logger.log('✗ Error: ' + (response.error || 'Unknown error'));
    }
    
  } catch (error) {
    Logger.log('✗ Error in testLastRow: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
  }
}

// ============================================================================
// SETUP INSTRUCTIONS
// ============================================================================
//
// STEP 1: VERIFY COLUMN MAPPING
// ------------------------------
// 1. Run diagnoseSheetStructure() function
// 2. Check the logs to verify column mapping matches your sheet
// 3. If columns don't match, update COLUMN_MAPPING above
//
// STEP 2: INSTALL THE TRIGGER
// ----------------------------
// 1. Save this script
// 2. Click the clock icon (⏰) in the Apps Script toolbar
// 3. Click "+ Add Trigger" button
// 4. Configure:
//    - Function: onFormSubmit
//    - Event source: From form
//    - Event type: On form submit
// 5. Click "Save"
// 6. Authorize the script if prompted
//
// STEP 3: TEST
// ------------
// 1. Run testWebhook() to test webhook connection
// 2. Run testLastRow() to test with actual sheet data
// 3. Submit a test form to verify full flow
//
// TROUBLESHOOTING:
// ----------------
// - If sheet not found: Check SHEET_NAME matches your sheet name exactly (case-sensitive)
// - If columns wrong: Run diagnoseSheetStructure() and adjust COLUMN_MAPPING
// - If webhook fails: Check WEBHOOK_URL is correct and Make.com scenario is active
// - Check execution logs: View → Execution log in Apps Script
//
// ============================================================================

