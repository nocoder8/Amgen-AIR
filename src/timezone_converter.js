/**
 * AIR - Simple Timezone Converter
 * Converts Schedule_start_time (UTC) to Local Time based on Location_country
 * Call convertAllTimezones() from your refresh script
 */

// Configuration: Update these based on your sheet structure
const CONFIG = {
  SCHEDULE_START_TIME_COL: 13, // Column M (Schedule_start_time)
  LOCATION_COUNTRY_COL: 21,    // Column U (Location_country)
  LOCAL_TIME_COL: 30,           // Column AD or wherever you want Local Time
  DATA_START_ROW: 2,           // Row where data starts (assuming row 1 is headers)
  SHEET_NAME: 'Enriched'       // Your sheet name
};

/**
 * Country to Timezone Offset Mapping (in hours from UTC)
 * Add more countries as needed
 */
const COUNTRY_TIMEZONE_MAP = {
  // India
  'India': 5.5,
  'IN': 5.5,
  
  // United States
  'United States': -5, // EST default
  'USA': -5,
  'US': -5,
  'America': -5,
  
  // United Kingdom
  'United Kingdom': 0,
  'UK': 0,
  'England': 0,
  
  // Canada
  'Canada': -5,
  'CA': -5,
  
  // Australia
  'Australia': 10,
  'AU': 10,
  
  // China
  'China': 8,
  'CN': 8,
  
  // Japan
  'Japan': 9,
  'JP': 9,
  
  // Germany
  'Germany': 1,
  'DE': 1,
  
  // France
  'France': 1,
  'FR': 1,
  
  // Brazil
  'Brazil': -3,
  'BR': -3,
  
  // Mexico
  'Mexico': -6,
  'MX': -6,
  
  // Singapore
  'Singapore': 8,
  'SG': 8,
  
  // Add more as needed...
};

/**
 * Get timezone offset in hours for a given country
 */
function getTimezoneOffset(country) {
  if (!country || typeof country !== 'string') {
    return 0;
  }
  
  const countryUpper = country.trim().toUpperCase();
  
  // Direct lookup
  if (COUNTRY_TIMEZONE_MAP[country]) {
    return COUNTRY_TIMEZONE_MAP[country];
  }
  
  // Case-insensitive lookup
  for (const [key, value] of Object.entries(COUNTRY_TIMEZONE_MAP)) {
    if (key.toUpperCase() === countryUpper) {
      return value;
    }
  }
  
  // Partial match
  for (const [key, value] of Object.entries(COUNTRY_TIMEZONE_MAP)) {
    if (countryUpper.includes(key.toUpperCase()) || key.toUpperCase().includes(countryUpper)) {
      return value;
    }
  }
  
  // Default to UTC (0) if country not found
  return 0;
}

/**
 * Convert UTC timestamp to local time based on country
 */
function convertToLocalTime(utcTimestamp, country) {
  if (!utcTimestamp) {
    return null;
  }
  
  // Convert to Date object if it's a number (Excel/Sheets serial number)
  let utcDate;
  if (typeof utcTimestamp === 'number') {
    utcDate = new Date((utcTimestamp - 25569) * 86400 * 1000);
  } else if (utcTimestamp instanceof Date) {
    utcDate = new Date(utcTimestamp);
  } else {
    return null;
  }
  
  if (isNaN(utcDate.getTime())) {
    return null;
  }
  
  const offsetHours = getTimezoneOffset(country);
  
  // Add offset hours
  const localDate = new Date(utcDate.getTime() + (offsetHours * 60 * 60 * 1000));
  
  return localDate;
}

/**
 * Main function - Call this from your refresh script
 * Converts all timezones silently in the background
 */
function convertAllTimezones() {
  Logger.log('=== convertAllTimezones START ===');
  const startTime = new Date();
  
  try {
    Logger.log('Step 1: Getting spreadsheet...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(`Spreadsheet: ${ss.getName()}`);
    
    Logger.log(`Step 2: Getting sheet "${CONFIG.SHEET_NAME}"...`);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      Logger.log(`ERROR: Sheet "${CONFIG.SHEET_NAME}" not found`);
      return;
    }
    Logger.log(`Sheet found: ${sheet.getName()}`);
    
    Logger.log('Step 3: Checking/creating header...');
    const headerRange = sheet.getRange(1, CONFIG.LOCAL_TIME_COL);
    const currentHeader = headerRange.getValue();
    if (!currentHeader || currentHeader.toString().trim() === '') {
      headerRange.setValue('Local Time');
      headerRange.setFontWeight('bold');
      Logger.log('Header "Local Time" created');
    } else {
      Logger.log(`Header already exists: "${currentHeader}"`);
    }
    
    Logger.log('Step 4: Finding last row with actual data...');
    
    // Find the actual last row with data (not just formulas)
    // Check Schedule_start_time column for actual data
    const maxRows = sheet.getLastRow();
    let lastDataRow = CONFIG.DATA_START_ROW - 1;
    
    // Strategy: Check from the beginning first (where data usually is)
    // Then check the end if needed
    const checkFromStart = Math.min(5000, maxRows - CONFIG.DATA_START_ROW + 1);
    
    Logger.log(`Checking first ${checkFromStart} rows for actual data...`);
    const startCheckData = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.SCHEDULE_START_TIME_COL, checkFromStart, 1).getValues();
    
    // Find last row with actual data from the beginning
    for (let i = startCheckData.length - 1; i >= 0; i--) {
      const value = startCheckData[i][0];
      if (value !== null && value !== '' && value !== undefined && value !== 0) {
        lastDataRow = CONFIG.DATA_START_ROW + i;
        Logger.log(`Found data in first ${checkFromStart} rows, last row: ${lastDataRow}`);
        break;
      }
    }
    
    // If no data found in first 5000 rows, check the last 2000 rows
    if (lastDataRow < CONFIG.DATA_START_ROW && maxRows > checkFromStart) {
      Logger.log(`No data in first ${checkFromStart} rows, checking last 2000 rows...`);
      const checkRange = Math.min(2000, maxRows - CONFIG.DATA_START_ROW + 1);
      const startCheckRow = Math.max(CONFIG.DATA_START_ROW, maxRows - checkRange + 1);
      
      const endCheckData = sheet.getRange(startCheckRow, CONFIG.SCHEDULE_START_TIME_COL, maxRows - startCheckRow + 1, 1).getValues();
      
      for (let i = endCheckData.length - 1; i >= 0; i--) {
        const value = endCheckData[i][0];
        if (value !== null && value !== '' && value !== undefined && value !== 0) {
          lastDataRow = startCheckRow + i;
          Logger.log(`Found data in last rows, last row: ${lastDataRow}`);
          break;
        }
      }
    }
    
    Logger.log(`Last row with actual data: ${lastDataRow} (out of ${maxRows} total rows)`);
    
    // Debug: Check a few sample rows to see what data looks like
    if (lastDataRow < CONFIG.DATA_START_ROW) {
      Logger.log('No data rows found. Checking sample rows for debugging...');
      for (let sampleRow = CONFIG.DATA_START_ROW; sampleRow <= Math.min(CONFIG.DATA_START_ROW + 5, maxRows); sampleRow++) {
        const sampleTime = sheet.getRange(sampleRow, CONFIG.SCHEDULE_START_TIME_COL).getValue();
        const sampleCountry = sheet.getRange(sampleRow, CONFIG.LOCATION_COUNTRY_COL).getValue();
        Logger.log(`Row ${sampleRow}: Schedule_time = "${sampleTime}" (type: ${typeof sampleTime}), Country = "${sampleCountry}"`);
      }
      Logger.log('No data rows to process');
      return;
    }
    
    const numRows = lastDataRow - CONFIG.DATA_START_ROW + 1;
    Logger.log(`Total rows to process: ${numRows}`);
    
    const chunkSize = 1000;
    const totalChunks = Math.ceil(numRows / chunkSize);
    Logger.log(`Processing in ${totalChunks} chunk(s) of ${chunkSize} rows each`);
    
    // Process in chunks for better performance and memory management
    for (let chunkNum = 0, startIdx = 0; startIdx < numRows; startIdx += chunkSize, chunkNum++) {
      Logger.log(`--- Processing chunk ${chunkNum + 1}/${totalChunks} ---`);
      const endIdx = Math.min(startIdx + chunkSize, numRows);
      const chunkRows = endIdx - startIdx;
      const startRow = CONFIG.DATA_START_ROW + startIdx;
      
      Logger.log(`Chunk ${chunkNum + 1}: Reading rows ${startRow} to ${startRow + chunkRows - 1}...`);
      
      // Read chunk data in one batch
      const scheduleTimes = sheet.getRange(startRow, CONFIG.SCHEDULE_START_TIME_COL, chunkRows, 1).getValues();
      const countries = sheet.getRange(startRow, CONFIG.LOCATION_COUNTRY_COL, chunkRows, 1).getValues();
      Logger.log(`Chunk ${chunkNum + 1}: Data read successfully`);
      
      Logger.log(`Chunk ${chunkNum + 1}: Processing ${chunkRows} rows...`);
      // Process chunk - only process rows with actual data
      const localTimeValues = [];
      let rowsWithData = 0;
      
      for (let i = 0; i < chunkRows; i++) {
        try {
          const scheduleTime = scheduleTimes[i][0];
          const country = countries[i][0];
          
          // Only process if Schedule_start_time has actual data (not empty, null, undefined, or 0)
          if (scheduleTime !== null && scheduleTime !== '' && scheduleTime !== undefined && scheduleTime !== 0) {
            const localTime = convertToLocalTime(scheduleTime, country);
            localTimeValues.push([localTime || '']);
            rowsWithData++;
          } else {
            // Skip empty rows - don't write anything (empty string)
            localTimeValues.push(['']);
          }
        } catch (error) {
          Logger.log(`Error processing row ${startRow + i}: ${error.toString()}`);
          localTimeValues.push(['']);
        }
      }
      Logger.log(`Chunk ${chunkNum + 1}: Processed ${rowsWithData} rows with data out of ${chunkRows} total`);
      
      Logger.log(`Chunk ${chunkNum + 1}: Writing results...`);
      // Write chunk results in one batch
      if (localTimeValues.length > 0) {
        sheet.getRange(startRow, CONFIG.LOCAL_TIME_COL, localTimeValues.length, 1)
          .setValues(localTimeValues)
          .setNumberFormat('yyyy-mm-dd hh:mm:ss');
        Logger.log(`Chunk ${chunkNum + 1}: Write complete`);
      }
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    Logger.log(`=== convertAllTimezones COMPLETE in ${duration.toFixed(2)} seconds ===`);
    
  } catch (error) {
    Logger.log(`ERROR in convertAllTimezones: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    // Fail silently to not break refresh script
  }
}

/**
 * onEdit trigger - Automatically converts timezone when you edit Schedule_start_time or Location_country
 * NO SETUP NEEDED - This runs automatically when you edit the sheet
 * Only watches columns M (Schedule_start_time) and U (Location_country)
 */
function onEdit(e) {
  try {
    Logger.log('=== onEdit triggered ===');
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    
    Logger.log(`Sheet: ${sheetName}`);
    
    // Only process if it's the target sheet
    if (sheetName !== CONFIG.SHEET_NAME) {
      Logger.log(`Skipping - not target sheet (expected: ${CONFIG.SHEET_NAME})`);
      return;
    }
    
    const editedRow = e.range.getRow();
    const editedCol = e.range.getColumn();
    
    Logger.log(`Edit detected: Row ${editedRow}, Column ${editedCol}`);
    
    // Only process if Schedule_start_time (M=13) or Location_country (U=21) was edited
    if (editedCol !== CONFIG.SCHEDULE_START_TIME_COL && 
        editedCol !== CONFIG.LOCATION_COUNTRY_COL) {
      Logger.log(`Skipping - column ${editedCol} not watched (watching ${CONFIG.SCHEDULE_START_TIME_COL} and ${CONFIG.LOCATION_COUNTRY_COL})`);
      return;
    }
    
    // Skip header row
    if (editedRow < CONFIG.DATA_START_ROW) {
      Logger.log('Skipping - header row');
      return;
    }
    
    Logger.log(`Processing row ${editedRow}...`);
    
    // Process just this row
    const scheduleTime = sheet.getRange(editedRow, CONFIG.SCHEDULE_START_TIME_COL).getValue();
    const country = sheet.getRange(editedRow, CONFIG.LOCATION_COUNTRY_COL).getValue();
    
    Logger.log(`Row ${editedRow}: Schedule_time = ${scheduleTime}, Country = ${country}`);
    
    if (scheduleTime) {
      const localTime = convertToLocalTime(scheduleTime, country);
      if (localTime) {
        sheet.getRange(editedRow, CONFIG.LOCAL_TIME_COL)
          .setValue(localTime)
          .setNumberFormat('yyyy-mm-dd hh:mm:ss');
        Logger.log(`Row ${editedRow}: Converted to ${localTime}`);
      } else {
        Logger.log(`Row ${editedRow}: Conversion returned null`);
        sheet.getRange(editedRow, CONFIG.LOCAL_TIME_COL).clearContent();
      }
    } else {
      Logger.log(`Row ${editedRow}: No schedule time, clearing local time`);
      sheet.getRange(editedRow, CONFIG.LOCAL_TIME_COL).clearContent();
    }
    
    Logger.log('=== onEdit complete ===');
    
  } catch (error) {
    Logger.log(`ERROR in onEdit: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    // Don't throw - we don't want to break the edit operation
  }
}

/**
 * Setup function to create a daily trigger at 9 AM
 * Run this ONCE to set up the automatic daily conversion
 */
function setupDailyTrigger() {
  try {
    Logger.log('=== Setting up daily trigger ===');
    
    // Delete existing triggers for this function to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'convertAllTimezones') {
        ScriptApp.deleteTrigger(triggers[i]);
        deletedCount++;
        Logger.log(`Deleted existing trigger: ${triggers[i].getUniqueId()}`);
      }
    }
    
    if (deletedCount > 0) {
      Logger.log(`Deleted ${deletedCount} existing trigger(s)`);
    }
    
    // Create a new trigger to run daily at 9 AM
    const trigger = ScriptApp.newTrigger('convertAllTimezones')
      .timeBased()
      .everyDays(1) // Run daily
      .atHour(9)    // 9 AM
      .create();
    
    Logger.log(`Daily trigger created successfully!`);
    Logger.log(`Trigger ID: ${trigger.getUniqueId()}`);
    Logger.log(`Will run: Daily at 9:00 AM`);
    Logger.log('=== Trigger setup complete ===');
    
    // Show confirmation (optional - can be removed if you don't want alerts)
    SpreadsheetApp.getUi().alert('Daily trigger created successfully!\n\nThe timezone converter will run automatically every day at 9:00 AM.');
    
  } catch (error) {
    Logger.log(`ERROR setting up trigger: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`Error setting up trigger: ${error.toString()}`);
    throw error;
  }
}

/**
 * Function to delete all existing triggers (useful for cleanup)
 */
function deleteAllTriggers() {
  try {
    Logger.log('=== Deleting all triggers ===');
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'convertAllTimezones') {
        ScriptApp.deleteTrigger(triggers[i]);
        deletedCount++;
        Logger.log(`Deleted trigger: ${triggers[i].getUniqueId()}`);
      }
    }
    
    Logger.log(`Deleted ${deletedCount} trigger(s)`);
    SpreadsheetApp.getUi().alert(`Deleted ${deletedCount} trigger(s)`);
    
  } catch (error) {
    Logger.log(`ERROR deleting triggers: ${error.toString()}`);
    throw error;
  }
}

/**
 * Function to list all existing triggers (for debugging)
 */
function listTriggers() {
  try {
    Logger.log('=== Listing all triggers ===');
    const triggers = ScriptApp.getProjectTriggers();
    
    if (triggers.length === 0) {
      Logger.log('No triggers found');
      return;
    }
    
    triggers.forEach(trigger => {
      Logger.log(`Trigger ID: ${trigger.getUniqueId()}`);
      Logger.log(`Handler: ${trigger.getHandlerFunction()}`);
      Logger.log(`Event Type: ${trigger.getEventType()}`);
      
      if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
        Logger.log(`Trigger Source: ${trigger.getTriggerSource()}`);
      }
    });
    
  } catch (error) {
    Logger.log(`ERROR listing triggers: ${error.toString()}`);
  }
}
