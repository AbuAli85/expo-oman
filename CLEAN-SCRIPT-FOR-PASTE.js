/**
 * Google Apps Script for Make.com Webhook Integration
 * 
 * This script triggers the Make.com webhook when a Google Form is submitted.
 * 
 * Webhook URL: https://hook.eu2.make.com/aiflluclc1hiyci8slmm4rex79tfvn5u
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

// Your Make.com webhook URL
const WEBHOOK_URL = 'https://hook.eu2.make.com/aiflluclc1hiyci8slmm4rex79tfvn5u';

// Sheet name (update if your sheet has a different name)
// Common names: 'Form Responses 1', 'Sheet1', 'named leads', etc.
const SHEET_NAME = 'leads'; // Change to your actual sheet name

// ============================================================================
// COLUMN MAPPING
// ============================================================================

const COLUMN_MAPPING = {
  NAME: 1,           // Column A - Name
  EMAIL: 2,          // Column B - Email
  PHONE: 3,          // Column C - Phone
  LANGUAGE: 4,       // Column D - Language
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

function onFormSubmit(e) {
  try {
    // Check if event object exists (only when triggered by form submission)
    if (!e || !e.source) {
      Logger.log('Error: This function must be triggered by a form submission.');
      Logger.log('To test manually, use testLastRow() or testWebhook() functions instead.');
      return;
    }
    
    const sheet = e.source.getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log('Error: Sheet "' + SHEET_NAME + '" not found');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    const rowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const formData = getFormData(rowData);
    
    if (!formData.email || formData.email.trim() === '') {
      Logger.log('Error: Email is required but not found');
      return;
    }
    
    const response = sendToWebhook(formData);
    
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

function getFormData(rowData) {
  // Get timestamp from Date Added, or use current time
  let timestamp = rowData[COLUMN_MAPPING.DATE_ADDED - 1];
  if (!timestamp) {
    timestamp = new Date();
  }
  if (timestamp instanceof Date) {
    timestamp = timestamp.toISOString();
  } else if (timestamp && typeof timestamp === 'string') {
    // Keep as is if it's already a string
  } else {
    timestamp = new Date().toISOString();
  }
  
  const data = {
    timestamp: timestamp,
    name: rowData[COLUMN_MAPPING.NAME - 1] || '',
    email: rowData[COLUMN_MAPPING.EMAIL - 1] || '',
    phone: rowData[COLUMN_MAPPING.PHONE - 1] || '',
    language: rowData[COLUMN_MAPPING.LANGUAGE - 1] || 'en',
    status: rowData[COLUMN_MAPPING.STATUS - 1] || '',
    dateAdded: rowData[COLUMN_MAPPING.DATE_ADDED - 1] || '',
    response: rowData[COLUMN_MAPPING.RESPONSE - 1] || '',
    responseDate: rowData[COLUMN_MAPPING.RESPONSE_DATE - 1] || '',
    comments: rowData[COLUMN_MAPPING.COMMENTS - 1] || ''
  };
  
  // Clean up email (trim whitespace, convert to lowercase)
  if (data.email) {
    data.email = data.email.toString().trim().toLowerCase();
  }
  
  // Clean up name (trim whitespace)
  if (data.name) {
    data.name = data.name.toString().trim();
  }
  
  // Clean up phone (trim whitespace)
  if (data.phone) {
    data.phone = data.phone.toString().trim();
  }
  
  // Map response to attendance for backward compatibility
  data.attendance = data.response || '';
  
  return data;
}

function sendToWebhook(data) {
  try {
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
      muteHttpExceptions: true
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
    
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('');
    Logger.log('=== HEADER ROW ===');
    headerRow.forEach((header, index) => {
      Logger.log('Column ' + String.fromCharCode(65 + index) + ' (' + (index + 1) + '): "' + (header || '(empty)') + '"');
    });
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Test function - Sends test data to webhook
 * Use this to test the webhook connection without submitting a form
 */
function testWebhook() {
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
  
  Logger.log('Testing webhook with sample data...');
  const response = sendToWebhook(testData);
  
  if (response.success) {
    Logger.log('✓ Test successful! Status: ' + response.statusCode);
    Logger.log('Check Make.com for the execution.');
  } else {
    Logger.log('✗ Test failed! Error: ' + response.error);
  }
}

/**
 * Test function - Processes the last row from your sheet
 * Use this to test with actual data from your sheet
 */
function testLastRow() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      Logger.log('Error: Sheet "' + SHEET_NAME + '" not found');
      Logger.log('Available sheets: ' + spreadsheet.getSheets().map(s => s.getName()).join(', '));
      return;
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      Logger.log('Error: No data rows found. Sheet must have at least one row of data.');
      return;
    }
    
    Logger.log('Processing last row (' + lastRow + ') from sheet...');
    
    const rowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const formData = getFormData(rowData);
    
    Logger.log('Data extracted:');
    Logger.log('  Name: ' + formData.name);
    Logger.log('  Email: ' + formData.email);
    Logger.log('  Language: ' + formData.language);
    
    if (!formData.email || formData.email.trim() === '') {
      Logger.log('Error: Email is required but not found in row ' + lastRow);
      return;
    }
    
    const response = sendToWebhook(formData);
    
    if (response.success) {
      Logger.log('✓ Successfully sent data to webhook!');
      Logger.log('✓ Status: ' + response.statusCode);
      Logger.log('Check Make.com for the execution.');
    } else {
      Logger.log('✗ Failed to send data to webhook!');
      Logger.log('✗ Error: ' + response.error);
    }
    
  } catch (error) {
    Logger.log('Error in testLastRow: ' + error.toString());
  }
}

