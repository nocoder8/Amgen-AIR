// AIR Volkscience - Exec Summary - Company-Level AI Interview Analytics Script v1.0 (Recruiter Breakdown)
// Company: Amgen
// To: Pavan Kumar
// When: Daily, 10 AM (Can be adjusted)
// This script analyzes data from the Enriched sheet to provide company-wide insights
// including a breakdown by recruiter.

// --- Configuration ---
const VS_EMAIL_RECIPIENT_RB = 'pkumar@eightfold.ai'; // <<< UPDATE EMAIL RECIPIENT
const VS_EMAIL_CC_RB = 'pkumar@eightfold.ai'; // Optional CC
// Assuming the Log Enhanced sheet is in a separate Spreadsheet
const VS_LOG_SHEET_SPREADSHEET_URL_RB = 'https://docs.google.com/spreadsheets/d/1VxSsCRHdxpfhVTmrLW_BJjqPD59rjEFnHiXxrkL79y4/edit?gid=0#gid=0'; // <<< VERIFY SPREADSHEET URL
const VS_LOG_SHEET_NAME_RB = 'Enriched'; // <<< VERIFY SHEET NAME
const VS_REPORT_TIME_RANGE_DAYS_RB = 9999999; // Set large number to effectively include all time
const VS_COMPANY_NAME_RB = "Amgen"; // Used in report titles etc.

// --- Configuration for Application Sheet (for Adoption Chart) ---
// NOTE: Application sheet not available for Amgen - set to null to skip
const APP_SHEET_SPREADSHEET_URL_RB = null; // <<< Application Sheet not available
const APP_SHEET_NAME_RB = null; // <<< Application Sheet not available
const APP_LAUNCH_DATE_RB = null; // <<< Application Sheet not available
const APP_MATCH_SCORE_THRESHOLD_RB = 4; // <<< Not used when APP_SHEET is null


// --- Main Functions ---

/**
 * Creates a trigger to run the recruiter breakdown report daily.
 */
function createRecruiterBreakdownTrigger() {
  // Delete existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'AIR_DailySummarytoAP') { // Updated Handler Name
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Create a new trigger to run daily at 10 AM
  ScriptApp.newTrigger('AIR_DailySummarytoAP') // Updated Handler Name
    .timeBased()
    .everyDays(1) // Run daily
    .atHour(10) // Keep 10 AM or adjust as needed
    .create();
  Logger.log(`Daily trigger created for AIR_DailySummarytoAP (at 10 AM)`);
  // SpreadsheetApp.getUi().alert(`Daily trigger created for ${VS_COMPANY_NAME_RB} AI Interview Recruiter Report (at 10 AM).`); // Removed: Cannot call getUi from trigger context
}

/**
 * Main function to generate and send the company-level AI interview report with recruiter breakdown.
 * Renamed from AIR_RecruiterBreakdown_Daily.
 */
function AIR_DailySummarytoAP() {
  try {
    Logger.log(`--- Starting ${VS_COMPANY_NAME_RB} AI Interview Daily Summary Report Generation ---`); // Updated log

    // 1. Get Log Sheet Data (Uses RB config)
    const logData = getLogSheetDataRB();
    if (!logData || !logData.rows || logData.rows.length === 0) {
      Logger.log('No data found in the log sheet or required columns missing. Skipping report generation.');
      // Optional: Send an email notification about missing data/columns
      // sendVsErrorNotificationRB("Report Skipped: No data or required columns found in Log_Enhanced sheet.");
      return;
    }
     Logger.log(`Successfully retrieved ${logData.rows.length} rows from log sheet.`);

    // 1b. Get Application Sheet Data (for Adoption Chart and AI Coverage)
    // NOTE: Application sheet not available for Amgen - skip this section
    let adoptionChartData = null;
    let hiringMetrics = null;
    let aiCoverageMetrics = null;
    let validationSheetUrl = null;
    let recruiterValidationSheets = null;
    
    if (APP_SHEET_SPREADSHEET_URL_RB && APP_SHEET_NAME_RB) {
        try {
            const appData = getApplicationDataForChartRB();
            if (appData && appData.rows) {
                Logger.log(`Successfully retrieved ${appData.rows.length} rows from application sheet.`);
                adoptionChartData = calculateAdoptionMetricsForChartRB(appData.rows, appData.colIndices);
                Logger.log(`Successfully calculated adoption chart metrics.`);
                
                // Calculate hiring metrics
                hiringMetrics = calculateHiringMetricsFromAppData(appData.rows, appData.colIndices);
                Logger.log(`Successfully calculated hiring metrics.`);
                
                // Calculate AI coverage metrics
                aiCoverageMetrics = calculateAICoverageMetricsRB(appData.rows, appData.colIndices);
                if (aiCoverageMetrics) {
                    Logger.log(`Successfully calculated AI coverage metrics. Total eligible: ${aiCoverageMetrics.totalEligible}, Total AI interviews: ${aiCoverageMetrics.totalAIInterviews}, Overall percentage: ${aiCoverageMetrics.overallPercentage}%`);
                } else {
                    Logger.log(`WARNING: AI coverage metrics calculation returned null. This could be due to missing required columns.`);
                }
                
                // Create validation sheet for candidate count comparison
                try {
                    validationSheetUrl = createCandidateCountValidationSheet(appData.rows, appData.colIndices);
                    Logger.log(`Successfully created validation sheet: ${validationSheetUrl}`);
                } catch (validationError) {
                    Logger.log(`Warning: Could not create validation sheet: ${validationError.toString()}`);
                }
                
                // Create detailed validation sheets for each recruiter
                try {
                    recruiterValidationSheets = createAllRecruiterValidationSheets(appData.rows, appData.colIndices);
                    if (recruiterValidationSheets) {
                        Logger.log(`Successfully created ${recruiterValidationSheets.successfulSheets} recruiter validation sheets`);
                    } else {
                        Logger.log(`Warning: Could not create recruiter validation sheets`);
                    }
                } catch (validationError) {
                    Logger.log(`Warning: Could not create recruiter validation sheets: ${validationError.toString()}`);
                }
                
                // Logger.log(`Adoption Chart Data: ${JSON.stringify(adoptionChartData, null, 2)}`);
            } else {
                Logger.log(`WARNING: No data retrieved from application sheet. Adoption chart, hiring metrics, and AI coverage will be skipped.`);
            }
        } catch (appError) {
            Logger.log(`ERROR retrieving or processing application data for adoption chart: ${appError.toString()}`);
            // Continue without adoption chart data
            // Optional: Send notification about this specific failure?
            sendVsErrorNotificationRB(`Error getting data for Adoption Chart from ${APP_SHEET_NAME_RB}`, appError.stack);
        }
    } else {
        Logger.log(`Application sheet not configured. Skipping adoption chart, hiring metrics, and AI coverage sections.`);
    }

    // 2. Filter Data by Time Range (using Interview_email_sent_at)
    const filteredData = filterDataByTimeRangeRB(logData.rows, logData.colIndices);
    if (filteredData.length === 0) {
        Logger.log(`No data found within the last ${VS_REPORT_TIME_RANGE_DAYS_RB} days. Skipping report.`);
        return;
    }
    Logger.log(`Filtered data to ${filteredData.length} rows based on the last ${VS_REPORT_TIME_RANGE_DAYS_RB} days.`);

    // 2a. Filter out excluded Feedback Template Names
    const feedbackTemplateIndex = logData.colIndices.hasOwnProperty('Feedback_template_name') ? logData.colIndices['Feedback_template_name'] : -1;
    const excludedTemplates = ["AI Coding Interview Metrics Feedback Form", "AI Functional Interview Feedback Form"];
    let templateFilteredData = filteredData;
    if (feedbackTemplateIndex !== -1) {
        const initialCount = templateFilteredData.length;
        templateFilteredData = templateFilteredData.filter(row => {
            if (row.length <= feedbackTemplateIndex) return true; // Keep rows without template column
            const templateName = row[feedbackTemplateIndex] ? String(row[feedbackTemplateIndex]).trim() : '';
            return !excludedTemplates.includes(templateName);
        });
        Logger.log(`Filtered out ${initialCount - templateFilteredData.length} rows with excluded Feedback_template_name. Count after template filter: ${templateFilteredData.length}`);
    } else {
        Logger.log("Feedback_template_name column not found. Skipping template filter.");
    }

    // 2b. Filter out specific Position Names
    const positionNameIndex = logData.colIndices.hasOwnProperty('Position_name') ? logData.colIndices['Position_name'] : -1;
    const positionToExclude = "AIR Testing";
    let finalFilteredData = templateFilteredData;
    if (positionNameIndex !== -1) {
        const initialCount = finalFilteredData.length;
        finalFilteredData = finalFilteredData.filter(row => {
            return !(row.length > positionNameIndex && row[positionNameIndex] === positionToExclude);
        });
        Logger.log(`Filtered out ${initialCount - finalFilteredData.length} rows with Position_name '${positionToExclude}'. Final count: ${finalFilteredData.length}`);
    } else {
        Logger.log("Skipping Position_name filter as column was not found.");
    }

    // Check if any data remains after all filters
    if (finalFilteredData.length === 0) {
         Logger.log(`No data remaining after position filtering. Skipping report.`);
         return;
    }

    // <<< Calculate Creator Last Sent Activity & Daily Trends >>>
    const creatorLastSentMap = new Map();
    const creatorDailyCounts = new Map(); // Map<CreatorId, Map<DateString, Count>>
    const creatorIdx_Log = logData.colIndices.hasOwnProperty('Creator_user_id') ? logData.colIndices['Creator_user_id'] : -1;
    const emailSentIdx_Log = logData.colIndices['Interview_email_sent_at']; // Already required

    if (creatorIdx_Log !== -1) {
        finalFilteredData.forEach(row => {
            if (row.length > Math.max(creatorIdx_Log, emailSentIdx_Log)) {
                const creatorId = row[creatorIdx_Log]?.trim();
                const rawSentDate = row[emailSentIdx_Log];
                if (creatorId && creatorId !== 'Unknown' && rawSentDate) {
                    const sentDate = vsParseDateSafeRB(rawSentDate);
                    if (sentDate) {
                        // Update Last Sent Date
                        if (!creatorLastSentMap.has(creatorId) || sentDate > creatorLastSentMap.get(creatorId)) {
                            creatorLastSentMap.set(creatorId, sentDate);
                        }

                        // Update Daily Count
                        const dateString = vsFormatDateRB(sentDate); // Use consistent format for map key
                        if (!creatorDailyCounts.has(creatorId)) {
                             creatorDailyCounts.set(creatorId, new Map());
                        }
                        const dailyMap = creatorDailyCounts.get(creatorId);
                        dailyMap.set(dateString, (dailyMap.get(dateString) || 0) + 1);
                    }
                }
            }
        });
        Logger.log(`Found last sent dates for ${creatorLastSentMap.size} creators and daily counts.`);
    } else {
        Logger.log(`Creator_user_id column not found in log sheet, cannot calculate last sent activity or trends.`);
    }

    const creatorActivityData = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Use start of today

    // Generate dates for the last 10 days (today back to 9 days ago)
    const trendDates = [];
    for (let i = 0; i <= 9; i++) {
        const date = new Date(today);
        date.setDate(today.getDate() - i);
        trendDates.push(date);
    }

    creatorLastSentMap.forEach((lastDate, creator) => {
        const timeDiff = today.getTime() - lastDate.getTime();
        const daysAgo = Math.floor(timeDiff / (1000 * 60 * 60 * 24)); // Calculate whole days

        // Build Daily Trend String
        const dailyMap = creatorDailyCounts.get(creator);
        const trendValues = trendDates.map(date => {
            const dayOfWeek = date.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
            if (dayOfWeek === 0) return 'Sun';
            if (dayOfWeek === 6) return 'Sat';
            const dateString = vsFormatDateRB(date);
            return dailyMap?.get(dateString) || 0;
        });
        const dailyTrend = trendValues.join(',');

        creatorActivityData.push({ creator: creator, daysAgo: daysAgo, dailyTrend: dailyTrend });
    });

    // Sort by days ago (most recent first), then alphabetically
    creatorActivityData.sort((a, b) => {
        if (a.daysAgo !== b.daysAgo) {
            return a.daysAgo - b.daysAgo;
        }
        return a.creator.localeCompare(b.creator);
    });
    // <<< End Creator Last Sent Activity Calculation >>>

    // <<< RESTORED: Deduplicate by Profile_id + Position_id, prioritizing by status rank >>>
    const profileIdIndex = logData.colIndices['Profile_id'];
    const positionIdIndex = logData.colIndices['Position_id'];
    const statusIndex = logData.colIndices['STATUS_COLUMN']; // Get the index determined earlier
    const groupedData = {}; // Key: "profileId_positionId", Value: { bestRank: rank, row: rowData }
    let skippedRowCount = 0;

    finalFilteredData.forEach(row => {
        // Ensure row has the necessary columns
        if (!row || row.length <= profileIdIndex || row.length <= positionIdIndex || row.length <= statusIndex) {
            skippedRowCount++;
            // Logger.log(`Skipping row during grouping due to missing ID or Status columns. Row: ${JSON.stringify(row)}`);
            return; // Skip this row
        }
        const profileId = row[profileIdIndex];
        const positionId = row[positionIdIndex];
        const status = row[statusIndex] ? String(row[statusIndex]).trim() : 'Unknown';

        if (!profileId || !positionId) { // Check for blank IDs
             skippedRowCount++;
            // Logger.log(`Skipping row during grouping due to blank Profile_id or Position_id. Row: ${JSON.stringify(row)}`);
            return; // Skip rows with blank IDs
        }

        const uniqueKey = `${profileId}_${positionId}`;
        const currentRank = vsGetStatusRankRB(status); // Use RB helper
        
        // Check if current row has SUBMITTED feedback
        const feedbackStatusIdx = logData.colIndices.hasOwnProperty('Feedback_status') ? logData.colIndices['Feedback_status'] : -1;
        const currentFeedbackStatus = (feedbackStatusIdx !== -1 && row[feedbackStatusIdx]) ? String(row[feedbackStatusIdx]).trim().toUpperCase() : '';
        const hasSubmittedFeedback = currentFeedbackStatus === 'SUBMITTED';

        if (!groupedData[uniqueKey]) {
            // First row for this key - store it
            groupedData[uniqueKey] = { bestRank: currentRank, row: row, hasSubmittedFeedback: hasSubmittedFeedback };
        } else {
            const existing = groupedData[uniqueKey];
            const existingFeedbackStatus = (feedbackStatusIdx !== -1 && existing.row[feedbackStatusIdx]) ? String(existing.row[feedbackStatusIdx]).trim().toUpperCase() : '';
            const existingHasSubmitted = existingFeedbackStatus === 'SUBMITTED';
            
            // Priority: Keep row with SUBMITTED feedback if either has it, otherwise keep better status
            if (hasSubmittedFeedback && !existingHasSubmitted) {
                // Current row has SUBMITTED, existing doesn't - prefer current
                groupedData[uniqueKey] = { bestRank: currentRank, row: row, hasSubmittedFeedback: true };
            } else if (!hasSubmittedFeedback && existingHasSubmitted) {
                // Existing has SUBMITTED, current doesn't - keep existing
                // Do nothing, keep existing row
            } else if (currentRank < existing.bestRank) {
                // Both have same feedback status, prefer better status rank
                groupedData[uniqueKey] = { bestRank: currentRank, row: row, hasSubmittedFeedback: hasSubmittedFeedback };
            }
            // Otherwise keep existing row
        }
    });

    if (skippedRowCount > 0) {
        Logger.log(`Skipped ${skippedRowCount} rows during deduplication due to missing IDs, status, or incomplete row data.`);
    }

    // Extract the best row for each unique key
    const deduplicatedData = Object.values(groupedData).map(entry => entry.row);

    Logger.log(`Deduplicated data based on Profile_id + Position_id (prioritizing status). Count changed from ${finalFilteredData.length} to ${deduplicatedData.length}.`);

    // Check if any data remains after deduplication
    if (deduplicatedData.length === 0) {
         Logger.log(`No data remaining after deduplication. Skipping report.`);
         return;
    }
    // <<< END RESTORED BLOCK >>>

    // 3. Calculate Metrics (Uses RB calculator)
    const metrics = calculateCompanyMetricsRB(deduplicatedData, logData.colIndices);
    Logger.log('Successfully calculated company metrics with recruiter breakdown.');
    
    // 3a. Count feedback from ALL rows (before deduplication) to capture all feedback submissions
    // This ensures we count all feedback even if there are duplicate Profile_id + Position_id combinations
    countFeedbackFromAllRows(finalFilteredData, logData.colIndices, metrics);
    Logger.log('Successfully counted feedback from all rows (including duplicates).');
    // Logger.log(`Calculated Metrics: ${JSON.stringify(metrics)}`); // Optional: Log detailed metrics

    // 4. Create HTML Report (Uses RB creator) - Pass adoption, activity data, and log creator index
    const htmlContent = createRecruiterBreakdownHtmlReport(metrics, adoptionChartData, creatorActivityData, creatorIdx_Log, hiringMetrics, validationSheetUrl, aiCoverageMetrics, recruiterValidationSheets);
    Logger.log('Successfully generated HTML report content.');

    // 5. Send Email (Uses RB functions/config)
    // Set static subject line for this specific report
    const reportTitle = `AI Recruiter Adoption: Daily Summary`; // <<< Renamed Subject
    sendVsEmailRB(VS_EMAIL_RECIPIENT_RB, VS_EMAIL_CC_RB, reportTitle, htmlContent);

    Logger.log(`--- AI Recruiter Adoption: Daily Summary generated and sent successfully! ---`); // Updated log message
    return `Report sent to ${VS_EMAIL_RECIPIENT_RB}`;

  } catch (error) {
    Logger.log(`Error in AIR_DailySummarytoAP: ${error.toString()} Stack: ${error.stack}`); // Updated log
    // Send error email (Uses RB notifier)
    sendVsErrorNotificationRB(`ERROR generating AI Recruiter Adoption: Daily Summary: ${error.toString()}`, error.stack);
    return `Error: ${error.toString()}`;
  }
}

// --- Data Retrieval and Processing Functions ---

/**
 * Counts feedback and pending from ALL rows (before deduplication) and adds to metrics.
 * This ensures we count all feedback submissions and pending requests even if there are duplicate Profile_id + Position_id combinations.
 * @param {Array<Array>} allRows All filtered rows before deduplication
 * @param {object} colIndices Column indices object
 * @param {object} metrics Metrics object to update
 */
function countFeedbackFromAllRows(allRows, colIndices, metrics) {
  const feedbackStatusIdx = colIndices.hasOwnProperty('Feedback_status') ? colIndices['Feedback_status'] : -1;
  const recruiterIdx = colIndices.hasOwnProperty('Recruiter_name') ? colIndices['Recruiter_name'] : -1;
  const creatorIdx = colIndices.hasOwnProperty('Creator_user_id') ? colIndices['Creator_user_id'] : -1;
  const jobFuncIdx = colIndices.hasOwnProperty('Job_function') ? colIndices['Job_function'] : -1;
  const countryIdx = colIndices.hasOwnProperty('Location_country') ? colIndices['Location_country'] : -1;
  
  if (feedbackStatusIdx === -1) {
    Logger.log('WARNING: Feedback_status column not found. Cannot count feedback/pending from all rows.');
    return;
  }
  
  // Reset feedback and pending counts (they were calculated from deduplicated data)
  metrics.totalFeedbackSubmitted = 0;
  Object.keys(metrics.byRecruiter).forEach(rec => {
    metrics.byRecruiter[rec].feedbackSubmitted = 0;
    metrics.byRecruiter[rec].pending = 0;
  });
  Object.keys(metrics.byCreator).forEach(crt => {
    metrics.byCreator[crt].feedbackSubmitted = 0;
    metrics.byCreator[crt].pending = 0;
  });
  Object.keys(metrics.byJobFunction).forEach(jf => {
    metrics.byJobFunction[jf].feedbackSubmitted = 0;
    metrics.byJobFunction[jf].pending = 0;
  });
  Object.keys(metrics.byCountry).forEach(ctry => {
    metrics.byCountry[ctry].feedbackSubmitted = 0;
    metrics.byCountry[ctry].pending = 0;
  });
  
  // Count feedback and pending from ALL rows
  allRows.forEach(row => {
    if (row.length > feedbackStatusIdx) {
      const feedbackStatusRaw = row[feedbackStatusIdx] ? String(row[feedbackStatusIdx]).trim() : '';
      const feedbackStatusNormalized = feedbackStatusRaw.toLowerCase().trim();
      const isFeedbackSubmitted = feedbackStatusNormalized === 'submitted';
      const isRequestedFeedback = feedbackStatusNormalized === 'requested';
      
      // Get breakdown values
      const recruiter = (recruiterIdx !== -1 && row.length > recruiterIdx && row[recruiterIdx]) ? String(row[recruiterIdx]).trim() : 'Unknown';
      const creator = (creatorIdx !== -1 && row.length > creatorIdx && row[creatorIdx]) ? String(row[creatorIdx]).trim() : 'Unknown';
      const jobFunc = (jobFuncIdx !== -1 && row.length > jobFuncIdx && row[jobFuncIdx]) ? String(row[jobFuncIdx]).trim() : 'Unknown';
      const country = (countryIdx !== -1 && row.length > countryIdx && row[countryIdx]) ? String(row[countryIdx]).trim() : 'Unknown';
      
      // Initialize if needed
      if (!metrics.byRecruiter[recruiter]) {
        metrics.byRecruiter[recruiter] = { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, recruiterSubmissionAwaited: 0, statusCounts: {} };
      }
      if (!metrics.byCreator[creator]) {
        metrics.byCreator[creator] = { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, recruiterSubmissionAwaited: 0, statusCounts: {} };
      }
      if (!metrics.byJobFunction[jobFunc]) {
        metrics.byJobFunction[jobFunc] = { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, recruiterSubmissionAwaited: 0, statusCounts: {} };
      }
      if (!metrics.byCountry[country]) {
        metrics.byCountry[country] = { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, statusCounts: {} };
      }
      
      // Count feedback submissions
      if (isFeedbackSubmitted) {
        metrics.totalFeedbackSubmitted++;
        metrics.byRecruiter[recruiter].feedbackSubmitted++;
        metrics.byCreator[creator].feedbackSubmitted++;
        metrics.byJobFunction[jobFunc].feedbackSubmitted++;
        metrics.byCountry[country].feedbackSubmitted++;
      }
      
      // Count pending (based on Feedback_status = REQUESTED)
      if (isRequestedFeedback) {
        metrics.byRecruiter[recruiter].pending++;
        metrics.byCreator[creator].pending++;
        metrics.byJobFunction[jobFunc].pending++;
        metrics.byCountry[country].pending++;
      }
    }
  });
  
  Logger.log(`Counted feedback from all rows: ${metrics.totalFeedbackSubmitted} total feedback submissions`);
}

/**
 * Reads and processes data from the Log_Enhanced sheet for Recruiter Breakdown.
 * @returns {object|null} Object { rows: Array<Array>, headers: Array<string>, colIndices: object } or null if error/no sheet/missing columns.
 */
function getLogSheetDataRB() {
  Logger.log(`Attempting to open log spreadsheet: ${VS_LOG_SHEET_SPREADSHEET_URL_RB}`);
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openByUrl(VS_LOG_SHEET_SPREADSHEET_URL_RB);
    Logger.log(`Opened log spreadsheet: ${spreadsheet.getName()}`);
  } catch (e) {
    Logger.log(`Error opening log spreadsheet by URL: ${e}`);
    throw new Error(`Could not open the specified Log Spreadsheet URL. Please verify the URL is correct and accessible: ${VS_LOG_SHEET_SPREADSHEET_URL_RB}`);
  }

  let sheet = spreadsheet.getSheetByName(VS_LOG_SHEET_NAME_RB);

  // Fallback sheet finding logic
  if (!sheet) {
    Logger.log(`Log sheet "${VS_LOG_SHEET_NAME_RB}" not found by name. Attempting to use sheet by gid or first sheet.`);
    const gidMatch = VS_LOG_SHEET_SPREADSHEET_URL_RB.match(/gid=(\d+)/);
    if (gidMatch && gidMatch[1]) {
      const gid = gidMatch[1];
      const sheets = spreadsheet.getSheets();
      sheet = sheets.find(s => s.getSheetId().toString() === gid);
      if (sheet) Logger.log(`Using log sheet by ID: "${sheet.getName()}"`);
    }
    if (!sheet) {
      sheet = spreadsheet.getSheets()[0];
      if (sheet) {
        Logger.log(`Warning: Using first available sheet in log spreadsheet: "${sheet.getName()}"`);
      } else {
        throw new Error(`Could not find any sheets in the log spreadsheet: ${VS_LOG_SHEET_SPREADSHEET_URL_RB}`);
      }
    }
  } else {
     Logger.log(`Using specified log sheet: "${sheet.getName()}"`);
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  if (data.length < 2) {
    Logger.log(`Not enough data in log sheet "${sheet.getName()}". Found ${data.length} rows. Expected headers + data.`);
    return null;
  }

  const headers = data[0].map(String);
  const rows = data.slice(1);

  // <<< DEBUGGING: Log the headers the script actually sees >>>
  Logger.log(`DEBUG: Headers found in sheet: ${JSON.stringify(headers)}`);
  // <<< END DEBUGGING >>>

  const requiredColumns = [
      'Interview_email_sent_at',
      'Profile_id',
      'Position_id',
      // Status column - prioritize Interview Status_Real
  ];
  const optionalColumns = [
      'Candidate_name',
      'Position_name',
      'Interview_status',
      'Interview Status_Real',
      'Schedule_start_time', 'Duration_minutes', 'Feedback_status', 'Feedback_json',
      'Match_stars', 'Location_country', 'Job_function', 'Position_id', 'Recruiter_name', // Ensure Recruiter_name is here
      'Creator_user_id', 'Reviewer_email', 'Hiring_manager_name',
      'Days_pending_invitation', 'Interview Status_Real',
      'Position_approved_date',
      'Feedback_template_name' // Added for filtering excluded templates
  ];

  const colIndices = {};
  const missingCols = [];

  // --- Find Status Column --- Enforce Interview Status_Real ---
  const statusColName = 'Interview_status_real'; // <<< Updated name
  const statusColIndex = headers.indexOf(statusColName);
  if (statusColIndex !== -1) {
      colIndices['STATUS_COLUMN'] = statusColIndex;
      Logger.log(`Using column "${statusColName}" (index ${statusColIndex}) for interview status analysis.`);
  } else {
      missingCols.push(statusColName);
  }
  // --- End Find Status Column ---

  requiredColumns.forEach(colName => {
    const index = headers.indexOf(colName);
    if (index === -1) {
      missingCols.push(colName);
    } else {
      colIndices[colName] = index;
    }
  });

  // Check for Recruiter_name specifically as it's needed for the breakdown
  if (headers.indexOf('Recruiter_name') === -1) {
      Logger.log(`WARNING: Optional column "Recruiter_name" not found. Recruiter breakdown will show 'Unknown'.`);
  }

  if (missingCols.length > 0) {
    Logger.log(`ERROR: Missing required column(s) in log sheet "${sheet.getName()}": ${missingCols.join(', ')}`);
    throw new Error(`Required column(s) not found in log sheet headers (Row 1): ${missingCols.join(', ')}`);
  }

  optionalColumns.forEach(colName => {
      const index = headers.indexOf(colName);
      if (index !== -1) {
          colIndices[colName] = index;
      } else if(colName !== 'Recruiter_name') { // Only log missing optional if not Recruiter_name (already warned)
          Logger.log(`Optional column "${colName}" not found.`);
      }
  });

  Logger.log(`Found required columns. Indices: ${JSON.stringify(colIndices)}`);
  return { rows, headers, colIndices };
}

/**
 * Reads and processes data from the Application Sheet (e.g., Active+Rejected) for the Adoption Chart.
 * @returns {object|null} Object { rows: Array<Array>, headers: Array<string>, colIndices: object } or null.
 */
function getApplicationDataForChartRB() {
  Logger.log(`--- Starting getApplicationDataForChartRB ---`);
  
  // Check if application sheet is configured
  if (!APP_SHEET_SPREADSHEET_URL_RB || !APP_SHEET_NAME_RB) {
    Logger.log(`Application sheet not configured. Returning null.`);
    return null;
  }
  
  Logger.log(`Attempting to open application spreadsheet: ${APP_SHEET_SPREADSHEET_URL_RB}`);
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openByUrl(APP_SHEET_SPREADSHEET_URL_RB);
    Logger.log(`Opened application spreadsheet: ${spreadsheet.getName()}`);
  } catch (e) {
    Logger.log(`Error opening application spreadsheet by URL: ${e}`);
    // Throw error as this data is essential for the requested chart
    throw new Error(`Could not open the Application Spreadsheet URL: ${APP_SHEET_SPREADSHEET_URL_RB}. Please verify the URL.`);
  }

  let sheet = spreadsheet.getSheetByName(APP_SHEET_NAME_RB);

  // Fallback sheet finding logic (similar to weekly report)
  if (!sheet) {
    Logger.log(`App sheet "${APP_SHEET_NAME_RB}" not found by name. Trying by GID or first sheet.`);
    const gidMatch = APP_SHEET_SPREADSHEET_URL_RB.match(/gid=(\d+)/);
    if (gidMatch && gidMatch[1]) {
      const gid = gidMatch[1];
      const sheets = spreadsheet.getSheets();
      sheet = sheets.find(s => s.getSheetId().toString() === gid);
      if (sheet) Logger.log(`Using app sheet by ID: "${sheet.getName()}"`);
    }
    if (!sheet) {
      sheet = spreadsheet.getSheets()[0];
      if (!sheet) {
        throw new Error(`No sheets found in application spreadsheet: ${APP_SHEET_SPREADSHEET_URL_RB}`);
      }
      Logger.log(`Warning: App sheet "${APP_SHEET_NAME_RB}" not found. Using first sheet: "${sheet.getName()}"`);
    }
  } else {
     Logger.log(`Using specified app sheet: "${sheet.getName()}"`);
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  // Expect headers in Row 2, data starts Row 3 (like AIR_Weekly_Recruiter_Report.js)
  if (data.length < 3) {
    Logger.log(`Not enough data in app sheet "${sheet.getName()}" (expected headers in row 2). Cannot generate adoption chart.`);
    return null; // Return null, main function should handle this
  }

  const headers = data[1].map(String); // Headers from Row 2
  const rows = data.slice(2); // Data from Row 3 onwards

  Logger.log(`DEBUG: App Sheet Headers found in row 2: ${JSON.stringify(headers)}`);

  // --- Find Match Stars Column (Copied from weekly report logic) ---
  let matchStarsColIndex = -1;
  const exactMatchCol = 'Match_stars';
  matchStarsColIndex = headers.indexOf(exactMatchCol);
  if (matchStarsColIndex === -1) {
    Logger.log(`"${exactMatchCol}" column not found directly. Searching for alternatives...`);
    const possibleMatchColumns = ['Match_score', 'Match score', 'Match Stars', 'MatchStars', 'Match_Stars', 'Stars', 'Score'];
    for (const columnName of possibleMatchColumns) {
      matchStarsColIndex = headers.indexOf(columnName);
      if (matchStarsColIndex !== -1) {
        Logger.log(`Found match stars column as "${columnName}" at index ${matchStarsColIndex}`);
        break;
      }
    }
    // Add more fuzzy matching if needed here, similar to weekly report
  }
  if (matchStarsColIndex === -1) {
     Logger.log("WARNING: Could not find any suitable column for Match Stars/Score in App sheet. Adoption chart filter (≥4 Match) cannot be applied accurately.");
     // Proceed without it, the calculation function will handle this
  }
  // --- End Find Match Stars Column ---

  // Define columns needed for the adoption calculation
  const requiredAppColumns = [
      'Profile_id', 'Name', 'Last_stage', 'Ai_interview', 'Recruiter name', 'Application_status', 'Position_status', 'Application_ts', 'Position_id', 'Title'
      // Add other columns if the weekly report's generateSegmentMetrics uses them
  ];

  const appColIndices = {};
  const missingAppCols = [];

  requiredAppColumns.forEach(colName => {
      const index = headers.indexOf(colName);
      if (index === -1) {
          missingAppCols.push(colName);
      } else {
          appColIndices[colName] = index;
      }
  });

  // Add match stars index if found
  if (matchStarsColIndex !== -1) {
      appColIndices['Match_stars'] = matchStarsColIndex; // Use a consistent key
  }

  // Add Position_approved_date index if found
  const positionApprovedDateIndex = headers.indexOf('Position approved date');
  if (positionApprovedDateIndex !== -1) {
      appColIndices['Position approved date'] = positionApprovedDateIndex;
      Logger.log(`Found Position approved date column at index ${positionApprovedDateIndex}`);
  } else {
      Logger.log(`Optional column "Position approved date" not found. Candidate count comparison will be unavailable.`);
      // Try to find similar column names
      const possibleNames = ['Position approved date', 'Position_approved_date', 'Position Approved Date', 'Approved_date', 'Approved Date'];
      for (const name of possibleNames) {
          const index = headers.indexOf(name);
          if (index !== -1) {
              Logger.log(`Found similar column "${name}" at index ${index}. Please update the script to use this column name.`);
              break;
          }
      }
      
      // Search for any column containing "position" and "approved" or "date"
      const positionRelatedColumns = headers.filter(header => 
          header.toLowerCase().includes('position') && 
          (header.toLowerCase().includes('approved') || header.toLowerCase().includes('date'))
      );
      if (positionRelatedColumns.length > 0) {
          Logger.log(`Found position-related columns: ${positionRelatedColumns.join(', ')}`);
      }
  }

  if (missingAppCols.length > 0) {
    Logger.log(`ERROR: Missing required column(s) in app sheet "${sheet.getName()}" for adoption chart: ${missingAppCols.join(', ')}`);
    throw new Error(`Required column(s) for adoption chart not found in app sheet headers (Row 2): ${missingAppCols.join(', ')}`);
  }

  Logger.log(`Found required columns for app data chart. Indices: ${JSON.stringify(appColIndices)}`);
  return { rows, headers, colIndices: appColIndices };
}

/**
 * Calculates adoption metrics based on application data, mirroring the weekly report logic.
 * Filters for post-launch, >=4 match score (if possible), and calculates adoption based on eligibility.
 * @param {Array<Array>} appRows Raw rows from the application sheet.
 * @param {object} appColIndices Column indices map for the application sheet.
 * @returns {object} An object containing recruiter adoption data { recruiter: string, totalCandidates: number, takenAI: number, adoptionRate: number }.
 */
function calculateAdoptionMetricsForChartRB(appRows, appColIndices) {
  Logger.log(`--- Starting calculateAdoptionMetricsForChartRB ---`);

  const matchStarsColIndex = appColIndices.hasOwnProperty('Match_stars') ? appColIndices['Match_stars'] : -1;
  const launchDate = APP_LAUNCH_DATE_RB; // Use configured launch date
  const scoreThreshold = APP_MATCH_SCORE_THRESHOLD_RB; // Use configured threshold
  const applyMatchFilter = matchStarsColIndex !== -1;

  // 1. Filter for Post-Launch Date
  let postLaunchCandidates = appRows.filter(row => {
    const rawDate = row.length > appColIndices['Application_ts'] ? row[appColIndices['Application_ts']] : null;
    if (rawDate === null || rawDate === undefined || rawDate === '') return false;
    const applicationDate = vsParseDateSafeRB(rawDate); // Use RB helper
    return applicationDate && applicationDate >= launchDate;
  });
  Logger.log(`Total post-launch candidates (valid date): ${postLaunchCandidates.length}`);

  // 2. Filter by Match Score (if possible)
  let filteredCandidates = postLaunchCandidates;
  if (applyMatchFilter) {
    Logger.log(`Filtering segment by Match Score >= ${scoreThreshold}. Initial count: ${postLaunchCandidates.length}. Score Column Index: ${matchStarsColIndex}`);
    filteredCandidates = postLaunchCandidates.filter(row => {
      if (row.length <= matchStarsColIndex) return false;
      const scoreValue = row[matchStarsColIndex];
      const matchScore = parseFloat(scoreValue);
      return !isNaN(matchScore) && matchScore >= scoreThreshold;
    });
    Logger.log(`After match score filter, count: ${filteredCandidates.length}`);
  } else {
    Logger.log(`Match score filter not applied (column index: ${matchStarsColIndex}). Using all ${postLaunchCandidates.length} post-launch candidates for adoption chart.`);
  }

  // 3. Calculate Eligibility and Adoption (based on weekly report logic)
  const recruiterMap = {};
  let totalEligibleForRate = 0;
  let totalTakenForRate = 0;

  filteredCandidates.forEach(row => {
    const aiInterview = row.length > appColIndices['Ai_interview'] ? row[appColIndices['Ai_interview']] : null;
    const appStatus = row.length > appColIndices['Application_status'] ? row[appColIndices['Application_status']]?.toLowerCase() : null;
    const recruiter = (row.length > appColIndices['Recruiter name'] && row[appColIndices['Recruiter name']]) ? row[appColIndices['Recruiter name']] : 'Unassigned';

    let isEligible = false;
    let tookAI = false;

    if (aiInterview === 'Y') {
      isEligible = true;
      tookAI = true;
    } else if (aiInterview === 'N' || aiInterview === null || aiInterview === undefined || aiInterview === '') {
      // Eligible if not 'Y' AND not 'Rejected'
      if (appStatus !== 'rejected') {
        isEligible = true;
        tookAI = false;
      }
    }

    if (isEligible) {
      totalEligibleForRate++;
      if (!recruiterMap[recruiter]) {
        recruiterMap[recruiter] = { totalEligible: 0, taken: 0 };
      }
      recruiterMap[recruiter].totalEligible++;

      if (tookAI) {
        totalTakenForRate++;
        recruiterMap[recruiter].taken++;
      }
    }
  });

  Logger.log(`Adoption Chart Metrics: Total eligible (post-launch, >=${scoreThreshold} match) = ${totalEligibleForRate}. Total taken AI = ${totalTakenForRate}.`);

  // 4. Format Recruiter Data
  const recruiterAdoptionData = Object.keys(recruiterMap).map(recruiter => {
    const data = recruiterMap[recruiter];
    const adoptionRate = data.totalEligible > 0 ? parseFloat(((data.taken / data.totalEligible) * 100).toFixed(1)) : 0;
    return {
      recruiter: recruiter,
      totalCandidates: data.totalEligible, // Eligible candidates for this recruiter
      takenAI: data.taken,
      adoptionRate: adoptionRate
    };
  }).sort((a, b) => a.recruiter.localeCompare(b.recruiter)); // Sort alphabetically

  // Return structure expected by the chart generation code
  return { recruiterAdoptionData, hasMatchStarsColumn: matchStarsColIndex !== -1, matchScoreThreshold: scoreThreshold };
}

/**
 * Filters the data based on a time range (e.g., last N days based on Interview_email_sent_at).
 * @param {Array<Array>} rows The data rows.
 * @param {object} colIndices Map of column names to indices.
 * @returns {Array<Array>} Filtered rows.
 */
function filterDataByTimeRangeRB(rows, colIndices) {
  if (!colIndices.hasOwnProperty('Interview_email_sent_at')) {
      Logger.log("WARNING: Cannot filter by time range - 'Interview_email_sent_at' column index not found.");
      return rows;
  }

  const sentAtIndex = colIndices['Interview_email_sent_at'];
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - VS_REPORT_TIME_RANGE_DAYS_RB); // Use RB config
  const cutoffTimestamp = cutoffDate.getTime();

  Logger.log(`Filtering data for interviews sent on or after ${cutoffDate.toLocaleDateString()}`);

  const filteredRows = rows.filter(row => {
    if (row.length <= sentAtIndex) return false;
    const rawDate = row[sentAtIndex];
    const sentDate = vsParseDateSafeRB(rawDate); // Use RB helper
    return sentDate && sentDate.getTime() >= cutoffTimestamp;
  });

  return filteredRows;
}


/**
 * Calculates company-level metrics including recruiter breakdown from the filtered data.
 * @param {Array<Array>} filteredRows The filtered data rows.
 * @param {object} colIndices Map of column names to indices.
 * @returns {object} An object containing calculated metrics.
 */
function calculateCompanyMetricsRB(filteredRows, colIndices) {
  const COMPLETION_RATE_MATURITY_DAYS_RB = 1; // Exclude invites sent < 1 day ago for KPI box calc

  // Calculate the exact timestamp 24 hours ago from now
  const now = new Date();
  const cutoffTimestampForCompletionRate = now.getTime() - (48 * 60 * 60 * 1000); // Use 48 hours
  const cutoffDateForLog = new Date(cutoffTimestampForCompletionRate);
  Logger.log(`Calculating KPI box completion rate only for invites sent before ${cutoffDateForLog.toLocaleString()} (48-hour cutoff)`);

  // September 10th, 2025 cutoff for Post-Sept 10th metrics
  const SEPT_10_2025 = new Date('2025-09-10');
  const SEPT_10_2025_TIMESTAMP = SEPT_10_2025.getTime();
  Logger.log(`Calculating Post-Sept 10th completion rate for invites sent after ${SEPT_10_2025.toLocaleDateString()}`);


  const metrics = {
    reportStartDate: (() => { const d = new Date(); d.setDate(d.getDate() - VS_REPORT_TIME_RANGE_DAYS_RB); return vsFormatDateRB(d); })(), // Use RB config/helpers
    reportEndDate: vsFormatDateRB(new Date()), // Use RB helper
    totalSent: filteredRows.length, // This remains the absolute total after filtering/deduplication
    totalScheduled: 0,
    totalCompleted: 0, // Absolute total completed
    totalFeedbackSubmitted: 0,
    sentToScheduledRate: 0,
    scheduledToCompletedRate: 0,
    completedToFeedbackRate: 0,
    sentToScheduledDaysSum: 0,
    sentToScheduledCount: 0,
    completedToFeedbackDaysSum: 0,
    completedToFeedbackCount: 0,
    matchStarsSum: 0,
    matchStarsCount: 0,
    completionRateByJobFunction: {}, // Kept for consistency, but maybe removed if not needed
    avgTimeToFeedbackByCountry: {}, // Kept for consistency
    interviewStatusDistribution: {},
    // Raw data storage for breakdowns
    byJobFunction: {},
    byCountry: {},
    byRecruiter: {}, // <<< ADDED Recruiter Breakdown
    byCreator: {}, // <<< ADDED Creator Breakdown
    // Timeseries data
    dailySentCounts: {},
    // --- Counters and Rates for KPI Box ---
    matureKpiTotalSent: 0,       // Denominator for adjusted KPI rate
    matureKpiTotalCompleted: 0,  // Numerator for adjusted KPI rate
    kpiCompletionRateAdjusted: 0,// The adjusted rate (%) for the KPI box
    completionRateOriginal: 0,   // The original rate (%) using all invites (for footnote)
    // --- Post-Sept 10th Metrics ---
    postSept10TotalSent: 0,
    postSept10TotalCompleted: 0,
    postSept10MatureTotalSent: 0,      // Excludes last 48 hours
    postSept10MatureTotalCompleted: 0, // Excludes last 48 hours
    postSept10CompletionRate: 0,       // Overall post-Sept 10th rate
    postSept10KpiCompletionRate: 0,    // Mature post-Sept 10th rate (for main box)
    // --- Post-Sept 10th Average Time Calculation ---
    postSept10SentToScheduledDaysSum: 0,
    postSept10SentToScheduledCount: 0,
    postSept10AvgTimeToScheduleDays: null
  };

  // --- Status Definitions (Consistent) ---
  const STATUSES_FOR_AVG_TIME_CALC = ['SCHEDULED', 'COMPLETED'];
  const COMPLETED_STATUSES = ['COMPLETED']; // <<< UPDATED: Strict definition for all metrics
  const PENDING_STATUSES = ['PENDING', 'INVITED', 'EMAIL SENT'];
  const FEEDBACK_SUBMITTED_STATUS = 'Submitted';
  const RECRUITER_SUBMISSION_AWAITED_FEEDBACK = 'AI_RECOMMENDED';

  // --- Column Indices (Check existence) ---
  const statusIdx = colIndices['STATUS_COLUMN'];
  const sentAtIdx = colIndices['Interview_email_sent_at'];
  const scheduledAtIdx = colIndices.hasOwnProperty('Schedule_start_time') ? colIndices['Schedule_start_time'] : -1;
  const candidateNameIdx = colIndices.hasOwnProperty('Candidate_name') ? colIndices['Candidate_name'] : -1;
  const feedbackStatusIdx = colIndices.hasOwnProperty('Feedback_status') ? colIndices['Feedback_status'] : -1;
  
  // Debug: Log Feedback_status column status and collect unique values
  const uniqueFeedbackStatuses = new Set();
  if (feedbackStatusIdx === -1) {
    Logger.log('WARNING: Feedback_status column not found in sheet');
  } else {
    Logger.log(`Feedback_status column found at index: ${feedbackStatusIdx}`);
  }
  const durationIdx = colIndices.hasOwnProperty('Duration_minutes') ? colIndices['Duration_minutes'] : -1;
  const matchStarsIdx = colIndices.hasOwnProperty('Match_stars') ? colIndices['Match_stars'] : -1;
  const jobFuncIdx = colIndices.hasOwnProperty('Job_function') ? colIndices['Job_function'] : -1;
  const countryIdx = colIndices.hasOwnProperty('Location_country') ? colIndices['Location_country'] : -1;
  const recruiterIdx = colIndices.hasOwnProperty('Recruiter_name') ? colIndices['Recruiter_name'] : -1; // <<< GET Recruiter Index
  const creatorIdx = colIndices.hasOwnProperty('Creator_user_id') ? colIndices['Creator_user_id'] : -1; // <<< GET Creator Index

  filteredRows.forEach(row => {
    // <<< MOVED: Define core values at the beginning of the loop >>>
    const statusRaw = row[statusIdx] ? String(row[statusIdx]).trim() : 'Unknown';
    const jobFunc = (jobFuncIdx !== -1 && row[jobFuncIdx]) ? String(row[jobFuncIdx]).trim() : 'Unknown';
    const country = (countryIdx !== -1 && row[countryIdx]) ? String(row[countryIdx]).trim() : 'Unknown';
    const recruiter = (recruiterIdx !== -1 && row[recruiterIdx]) ? String(row[recruiterIdx]).trim() : 'Unknown'; // <<< GET Recruiter Name
    const creator = (creatorIdx !== -1 && row[creatorIdx]) ? String(row[creatorIdx]).trim() : 'Unknown'; // <<< GET Creator ID
    const feedbackStatusRaw = (feedbackStatusIdx !== -1 && row[feedbackStatusIdx]) ? String(row[feedbackStatusIdx]).trim() : '';

    // --- Get Sent Date ---
    const sentDate = vsParseDateSafeRB(row[sentAtIdx]); // Use RB helper
    const isMatureForCompletionRate = sentDate && sentDate.getTime() < cutoffTimestampForCompletionRate; // Check if sent *before* the exact 24hr cutoff timestamp
    const isPostSept10 = sentDate && sentDate.getTime() >= SEPT_10_2025_TIMESTAMP; // Check if sent after Sept 10th, 2025

    // --- Increment Mature Sent Count for KPI ---
    if (isMatureForCompletionRate) {
        metrics.matureKpiTotalSent++;
    }

    // --- Increment Post-Sept 10th Sent Count ---
    if (isPostSept10) {
        metrics.postSept10TotalSent++;
        // Also track mature count for Post-Sept 10th (excludes last 48 hours)
        if (isMatureForCompletionRate) {
            metrics.postSept10MatureTotalSent++;
        }
    }

    // --- Daily Sent Counts ---
    if (sentDate) {
        const dateString = vsFormatDateRB(sentDate); // Use RB helper
        metrics.dailySentCounts[dateString] = (metrics.dailySentCounts[dateString] || 0) + 1;
    }

    // --- Initialize Breakdown Structures if they don't exist ---
    if (!metrics.byJobFunction[jobFunc]) {
        metrics.byJobFunction[jobFunc] = { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, recruiterSubmissionAwaited: 0, statusCounts: {} };
    }
    if (!metrics.byCountry[country]) {
        metrics.byCountry[country] = { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, statusCounts: {} };
    }
    if (!metrics.byRecruiter[recruiter]) { // <<< INITIALIZE Recruiter
        metrics.byRecruiter[recruiter] = { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, recruiterSubmissionAwaited: 0, statusCounts: {} };
    }
    if (!metrics.byCreator[creator]) { // <<< INITIALIZE Creator
        metrics.byCreator[creator] = { sent: 0, scheduled: 0, completed: 0, pending: 0, feedbackSubmitted: 0, recruiterSubmissionAwaited: 0, statusCounts: {} };
    }

    // --- Increment Base Counts (These always use the total number of records processed) ---
    metrics.byJobFunction[jobFunc].sent++;
    metrics.byCountry[country].sent++;
    metrics.byRecruiter[recruiter].sent++; // <<< INCREMENT Recruiter Sent
    metrics.byCreator[creator].sent++; // <<< INCREMENT Creator Sent
    metrics.interviewStatusDistribution[statusRaw] = (metrics.interviewStatusDistribution[statusRaw] || 0) + 1;
    metrics.byJobFunction[jobFunc].statusCounts[statusRaw] = (metrics.byJobFunction[jobFunc].statusCounts[statusRaw] || 0) + 1;
    metrics.byCountry[country].statusCounts[statusRaw] = (metrics.byCountry[country].statusCounts[statusRaw] || 0) + 1;
    metrics.byRecruiter[recruiter].statusCounts[statusRaw] = (metrics.byRecruiter[recruiter].statusCounts[statusRaw] || 0) + 1; // <<< INCREMENT Recruiter Status Count
    metrics.byCreator[creator].statusCounts[statusRaw] = (metrics.byCreator[creator].statusCounts[statusRaw] || 0) + 1; // <<< INCREMENT Creator Status Count

    // --- Calculate Avg Time Sent to Completion (Scheduled) ---
    // <<< UPDATED: Only calculate for strictly COMPLETED interviews >>>
    if (statusRaw === 'COMPLETED') {
        const candidateName = (candidateNameIdx !== -1 && row[candidateNameIdx]) ? row[candidateNameIdx] : 'Unknown Candidate';
        const scheduleDateForAvg = (scheduledAtIdx !== -1) ? vsParseDateSafeRB(row[scheduledAtIdx]) : null; // Use RB helper
        if (sentDate && scheduleDateForAvg) {
            const daysDiff = vsCalculateDaysDifferenceRB(sentDate, scheduleDateForAvg);
            if (daysDiff !== null) {
                // <<< UPDATED: Only include post-Sept 10th interviews for average time calculation >>>
                if (isPostSept10) {
                    metrics.postSept10SentToScheduledDaysSum += daysDiff;
                    metrics.postSept10SentToScheduledCount++;
                    // <<< ADDED: Detailed log for post-Sept 10th avg time calculation >>>
                    Logger.log(`PostSept10_AvgTimeCalc_Include: Candidate=[${candidateName}], Status=[${statusRaw}], Sent=[${sentDate.toISOString()}], Scheduled=[${scheduleDateForAvg.toISOString()}], DiffDays=[${daysDiff.toFixed(2)}]`);
                }
            }
        }
    }

    // --- Check if Scheduled (for breakdown counts) ---
    let isScheduledForCount = (statusRaw === 'SCHEDULED');
    if (isScheduledForCount) {
         metrics.totalScheduled++;
         metrics.byJobFunction[jobFunc].scheduled++;
         metrics.byCountry[country].scheduled++;
         metrics.byRecruiter[recruiter].scheduled++; // <<< INCREMENT Recruiter Scheduled
         metrics.byCreator[creator].scheduled++; // <<< INCREMENT Creator Scheduled
    }

    // --- Check if Pending (based on Feedback_status = REQUESTED) ---
    // Count pending based on Feedback_status = REQUESTED instead of interview status
    const feedbackStatusNormalized = feedbackStatusRaw.toLowerCase().trim();
    const isRequestedFeedback = feedbackStatusNormalized === 'requested';
    
    if (isRequestedFeedback) {
        metrics.byJobFunction[jobFunc].pending++;
        metrics.byCountry[country].pending++;
        metrics.byRecruiter[recruiter].pending++; // <<< INCREMENT Recruiter Pending
        metrics.byCreator[creator].pending++; // <<< INCREMENT Creator Pending
    }

    // --- Check if Completed ---
    let isCompleted = COMPLETED_STATUSES.includes(statusRaw);
    if (isCompleted) {
      metrics.totalCompleted++; // Increment original total completed
      metrics.byJobFunction[jobFunc].completed++; // Increment original breakdown completed
      metrics.byCountry[country].completed++;     // Increment original breakdown completed
      metrics.byRecruiter[recruiter].completed++; // Increment original breakdown completed
      metrics.byCreator[creator].completed++; // Increment original breakdown completed

      // Increment Mature Completed Count for KPI (ONLY if mature)
      if (isMatureForCompletionRate) {
          metrics.matureKpiTotalCompleted++;
      }

      // Increment Post-Sept 10th Completed Count
      if (isPostSept10) {
          metrics.postSept10TotalCompleted++;
          // Also track mature completed count for Post-Sept 10th (excludes last 48 hours)
          if (isMatureForCompletionRate) {
              metrics.postSept10MatureTotalCompleted++;
          }
      }

      // --- Calculate Match Stars ---
       if (matchStarsIdx !== -1 && row[matchStarsIdx] !== null && row[matchStarsIdx] !== '') {
           const stars = parseFloat(row[matchStarsIdx]);
           if (!isNaN(stars) && stars >= 0) {
               metrics.matchStarsSum += stars;
               metrics.matchStarsCount++;
           }
       }

       // --- Check for Feedback Submitted (case-insensitive match) ---
       // Match "SUBMITTED" (or any case variation like "Submitted", "submitted")
       const feedbackStatusNormalized = feedbackStatusRaw.toLowerCase().trim();
       const isFeedbackSubmitted = feedbackStatusNormalized === 'submitted';
       
       if (feedbackStatusIdx !== -1 && isFeedbackSubmitted) {
         metrics.totalFeedbackSubmitted++;
         metrics.byJobFunction[jobFunc].feedbackSubmitted++;
         metrics.byCountry[country].feedbackSubmitted++;
         metrics.byRecruiter[recruiter].feedbackSubmitted++; // <<< INCREMENT Recruiter Feedback Submitted
         metrics.byCreator[creator].feedbackSubmitted++; // <<< INCREMENT Creator Feedback Submitted
       }
       
       // Collect unique feedback status values for debugging
       if (feedbackStatusIdx !== -1 && feedbackStatusRaw) {
         uniqueFeedbackStatuses.add(feedbackStatusRaw);
       }

       // --- Check for Recruiter Submission Awaited (AI_RECOMMENDED in Feedback_status)
       if (feedbackStatusIdx !== -1 && feedbackStatusRaw === RECRUITER_SUBMISSION_AWAITED_FEEDBACK) {
           metrics.byJobFunction[jobFunc].recruiterSubmissionAwaited++;
           // Note: No country-specific count for this yet
           metrics.byRecruiter[recruiter].recruiterSubmissionAwaited++; // <<< INCREMENT Recruiter Submission Awaited
           metrics.byCreator[creator].recruiterSubmissionAwaited++; // <<< INCREMENT Creator Submission Awaited
       }
    }
  });

  // Debug: Log all unique feedback status values found
  if (feedbackStatusIdx !== -1) {
    Logger.log(`Found ${uniqueFeedbackStatuses.size} unique Feedback_status values: ${Array.from(uniqueFeedbackStatuses).join(', ')}`);
    Logger.log(`Total feedback submitted count: ${metrics.totalFeedbackSubmitted}`);
  }

  // --- Calculate Final Rates and Averages ---

  // Calculate ORIGINAL Completion Rate (for footnote)
  if (metrics.totalSent > 0) {
      metrics.completionRateOriginal = parseFloat(((metrics.totalCompleted / metrics.totalSent) * 100).toFixed(1));
  }

  // Calculate ADJUSTED Completion Rate (for KPI Box)
  if (metrics.matureKpiTotalSent > 0) {
      metrics.kpiCompletionRateAdjusted = parseFloat(((metrics.matureKpiTotalCompleted / metrics.matureKpiTotalSent) * 100).toFixed(1));
  }

  // Calculate Post-Sept 10th Completion Rates
  if (metrics.postSept10TotalSent > 0) {
      metrics.postSept10CompletionRate = parseFloat(((metrics.postSept10TotalCompleted / metrics.postSept10TotalSent) * 100).toFixed(1));
  }
  if (metrics.postSept10MatureTotalSent > 0) {
      metrics.postSept10KpiCompletionRate = parseFloat(((metrics.postSept10MatureTotalCompleted / metrics.postSept10MatureTotalSent) * 100).toFixed(1));
  }

  // Calculate other original rates
  if (metrics.totalSent > 0) {
      metrics.sentToScheduledRate = parseFloat(((metrics.totalScheduled / metrics.totalSent) * 100).toFixed(1));
      // Update status distribution calculation (uses totalSent)
      const statusCountsTemp = { ...metrics.interviewStatusDistribution };
      metrics.interviewStatusDistribution = {};
      for (const status in statusCountsTemp) {
          const count = statusCountsTemp[status];
          metrics.interviewStatusDistribution[status] = {
              count: count,
              percentage: parseFloat(((count / metrics.totalSent) * 100).toFixed(1))
          };
      }
  }
  if (metrics.totalScheduled > 0) {
      metrics.scheduledToCompletedRate = parseFloat(((metrics.totalCompleted / metrics.totalScheduled) * 100).toFixed(1));
  }
   if (metrics.totalCompleted > 0) {
      metrics.completedToFeedbackRate = parseFloat(((metrics.totalFeedbackSubmitted / metrics.totalCompleted) * 100).toFixed(1));
      if(metrics.matchStarsCount > 0) {
          metrics.avgMatchStars = parseFloat((metrics.matchStarsSum / metrics.matchStarsCount).toFixed(1));
      } else {
          metrics.avgMatchStars = null; // Ensure null if no stars
      }
   } else {
      metrics.avgMatchStars = null; // Ensure null if no completions
   }
   if (metrics.sentToScheduledCount > 0) {
       metrics.avgTimeToScheduleDays = parseFloat((metrics.sentToScheduledDaysSum / metrics.sentToScheduledCount).toFixed(1));
   } else {
       metrics.avgTimeToScheduleDays = null;
   }
   
   // Calculate Post-Sept 10th Average Time to Schedule
   if (metrics.postSept10SentToScheduledCount > 0) {
       metrics.postSept10AvgTimeToScheduleDays = parseFloat((metrics.postSept10SentToScheduledDaysSum / metrics.postSept10SentToScheduledCount).toFixed(1));
   } else {
       metrics.postSept10AvgTimeToScheduleDays = null;
   }
    if (metrics.completedToFeedbackCount > 0) {
        metrics.avgCompletedToFeedbackDays = parseFloat((metrics.completedToFeedbackDaysSum / metrics.completedToFeedbackCount).toFixed(1));
    } else {
         metrics.avgCompletedToFeedbackDays = null; // Example, if calculation added later
    }

    // <<< ADDED: Summary log before calculating average time >>>
    Logger.log(`AvgTimeCalc_Summary: Total Days Sum = ${metrics.sentToScheduledDaysSum.toFixed(2)}, Count = ${metrics.sentToScheduledCount}`);
    Logger.log(`PostSept10_AvgTimeCalc_Summary: Total Days Sum = ${metrics.postSept10SentToScheduledDaysSum.toFixed(2)}, Count = ${metrics.postSept10SentToScheduledCount}`);
    // <<< END ADDED >>>

    // <<< DEBUG LOGGING for KPI Rate >>>
    // <<< DEBUG LOGGING for KPI Rate >>>
    Logger.log(`KPI Rate Calculation: Mature Sent (Denominator) = ${metrics.matureKpiTotalSent}`);
    Logger.log(`KPI Rate Calculation: Mature Completed [Strict 'COMPLETED'] (Numerator) = ${metrics.matureKpiTotalCompleted}`);
    if (metrics.matureKpiTotalSent > 0) {
        Logger.log(`KPI Rate Calculation: Adjusted Rate = (${metrics.matureKpiTotalCompleted} / ${metrics.matureKpiTotalSent}) * 100 = ${metrics.kpiCompletionRateAdjusted}%`);
    } else {
        Logger.log(`KPI Rate Calculation: Adjusted Rate = N/A (Mature Sent is 0)`);
    }
    // <<< END DEBUG LOGGING >>>

  // --- Calculate Breakdown Metrics (Using ORIGINAL 'sent' and 'completed' counts for percentages) ---
  // Job Functions
  for (const func in metrics.byJobFunction) {
    const data = metrics.byJobFunction[func];
    data.scheduledRate = data.sent > 0 ? parseFloat(((data.scheduled / data.sent) * 100).toFixed(1)) : 0;
    data.completedNumber = data.completed; // Use original completed count
    data.completedPercentOfSent = data.sent > 0 ? parseFloat(((data.completed / data.sent) * 100).toFixed(1)) : 0; // Use original counts for %
    data.pendingNumber = data.pending;
    data.pendingPercentOfSent = data.sent > 0 ? parseFloat(((data.pending / data.sent) * 100).toFixed(1)) : 0;
    data.feedbackRate = data.completed > 0 ? parseFloat(((data.feedbackSubmitted / data.completed) * 100).toFixed(1)) : 0;
  }

  // Countries
  for (const ctry in metrics.byCountry) {
    const data = metrics.byCountry[ctry];
    data.completedNumber = data.completed; // Use original completed count
    data.completedPercentOfSent = data.sent > 0 ? parseFloat(((data.completed / data.sent) * 100).toFixed(1)) : 0; // Use original counts for %
    data.pendingNumber = data.pending;
    data.pendingPercentOfSent = data.sent > 0 ? parseFloat(((data.pending / data.sent) * 100).toFixed(1)) : 0;
    // Add other country-specific metrics here if needed
  }

  // Recruiters <<< CALCULATE Recruiter Breakdown Metrics (Using ORIGINAL counts)
  for (const rec in metrics.byRecruiter) {
    const data = metrics.byRecruiter[rec];
    // data.scheduledRate = data.sent > 0 ? parseFloat(((data.scheduled / data.sent) * 100).toFixed(1)) : 0; // Optional
    data.completedNumber = data.completed; // Use original completed count
    data.completedPercentOfSent = data.sent > 0 ? parseFloat(((data.completed / data.sent) * 100).toFixed(1)) : 0; // Use original counts for %
    data.pendingNumber = data.pending;
    data.pendingPercentOfSent = data.sent > 0 ? parseFloat(((data.pending / data.sent) * 100).toFixed(1)) : 0;
    data.feedbackRate = data.completed > 0 ? parseFloat(((data.feedbackSubmitted / data.completed) * 100).toFixed(1)) : 0; // Optional
  }

  // Creators <<< CALCULATE Creator Breakdown Metrics (Using ORIGINAL counts)
  for (const crt in metrics.byCreator) {
    const data = metrics.byCreator[crt];
    data.completedNumber = data.completed; // Use original completed count
    data.completedPercentOfSent = data.sent > 0 ? parseFloat(((data.completed / data.sent) * 100).toFixed(1)) : 0; // Use original counts for %
    data.pendingNumber = data.pending;
    data.pendingPercentOfSent = data.sent > 0 ? parseFloat(((data.pending / data.sent) * 100).toFixed(1)) : 0;
    data.feedbackRate = data.completed > 0 ? parseFloat(((data.feedbackSubmitted / data.completed) * 100).toFixed(1)) : 0; // Optional
  }


  Logger.log(`Metrics calculation complete (Recruiter). Total Sent: ${metrics.totalSent}, Completed: ${metrics.totalCompleted}`);
  Logger.log(`KPI Completion Rate (Adjusted): ${metrics.kpiCompletionRateAdjusted}%, Original Rate: ${metrics.completionRateOriginal}%`); // Log both rates
  Logger.log(`Post-Sept 10th Metrics - Sent: ${metrics.postSept10TotalSent}, Completed: ${metrics.postSept10TotalCompleted}, Mature Rate: ${metrics.postSept10KpiCompletionRate}%, Overall Rate: ${metrics.postSept10CompletionRate}%`);
  metrics.colIndices = colIndices;
  return metrics;
}

// --- Reporting Functions ---

/**
 * Generates the HTML for the table rows of the Recruiter Breakdown section.
 * Sorts recruiters by 'Sent' count descending.
 * @param {object} recruiterData The metrics.byRecruiter object.
 * @returns {string} HTML string for the table body rows.
 */
function generateRecruiterTableRowsHtml(recruiterData) {
    if (!recruiterData || Object.keys(recruiterData).length === 0) {
        return '<tr><td colspan="7" style="text-align:center; padding: 10px; border: 1px solid #e0e0e0; font-size: 12px;">No recruiter data found or Recruiter_name column missing.</td></tr>';
    }

    // Sort recruiters by Sent descending, keeping Unknown last
    const sortedRecruiters = Object.entries(recruiterData)
        .sort(([recA, dataA], [recB, dataB]) => {
            if (recA === 'Unknown') return 1;
            if (recB === 'Unknown') return -1;
            return dataB.sent - dataA.sent;
        });

    // Generate table rows HTML
    return sortedRecruiters
        .map(([rec, data], index) => {
            const bgColor = index % 2 === 0 ? '#fafafa' : '#ffffff';
            return `
                <tr style="background-color: ${bgColor};">
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; font-size: 12px; vertical-align: middle; font-weight: bold; width: 180px;">${rec}</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.sent}</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.completedNumber} (<span style="color: #0056b3;">${data.completedPercentOfSent}%</span>)</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.scheduled}</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.pendingNumber} (<span style="color: #0056b3;">${data.pendingPercentOfSent}%</span>)</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">${data.feedbackSubmitted}</td>
                    <td style="border: 1px solid #e0e0e0; padding: 6px 10px; text-align: center; font-size: 12px; vertical-align: middle;">
                      ${data.recruiterSubmissionAwaited > 0 ? 
                          `<span style="color: red; font-weight: bold;">${data.recruiterSubmissionAwaited}</span>` : 
                          data.recruiterSubmissionAwaited
                      }
                    </td>
                </tr>
            `;
        }).join('');
}

/**
 * Creates the HTML email report including recruiter breakdown.
 * Uses inline styles and table layouts for better email client compatibility.
 * @param {object} metrics The calculated metrics object.
 * @param {object} adoptionChartData The calculated adoption chart data object.
 * @param {Array<object>} creatorActivityData Array of {creator: string, daysAgo: number, dailyTrend: string}.
 * @param {number} creatorIdx_Log The index of the Creator_user_id column from the log sheet (-1 if not found).
 * @returns {string} The HTML content for the email body.
 */
function createRecruiterBreakdownHtmlReport(metrics, adoptionChartData, creatorActivityData, creatorIdx_Log, hiringMetrics, validationSheetUrl, aiCoverageMetrics, recruiterValidationSheets) {
  Logger.log(`DEBUG: AI Coverage Metrics in HTML report: ${aiCoverageMetrics ? 'Present' : 'Null/Undefined'}`);
  if (aiCoverageMetrics) {
    Logger.log(`DEBUG: AI Coverage Metrics details - Total eligible: ${aiCoverageMetrics.totalEligible}, Total AI interviews: ${aiCoverageMetrics.totalAIInterviews}, Overall percentage: ${aiCoverageMetrics.overallPercentage}%`);
  }
  
  // AI Insights feature removed
  // Helper to generate timeseries table (limited to last 7 days)
  const generateTimeseriesTable = (dailyCounts) => {
      const sortedDates = Object.keys(dailyCounts).sort((a, b) => {
          try {
              // Parsing DD-MMM-YY format
              const dateA = new Date(a.replace(/(\\d{2})-(\\w{3})-(\\d{2})/, '$2 $1, 20$3'));
              const dateB = new Date(b.replace(/(\\d{2})-(\\w{3})-(\\d{2})/, '$2 $1, 20$3'));
              return dateB - dateA; // Descending
          } catch (e) { return b.localeCompare(a); }
      });

      const sevenDaysAgo = new Date();
      sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
      sevenDaysAgo.setHours(0, 0, 0, 0);

      const filteredDates = sortedDates.filter(dateStr => {
          try {
              const date = new Date(dateStr.replace(/(\\d{2})-(\\w{3})-(\\d{2})/, '$2 $1, 20$3'));
              return date >= sevenDaysAgo;
          } catch (e) { return false; }
      });

      if (filteredDates.length === 0) return '<p style="font-size: 11px; color: #999; margin-top: 12px; text-align: center;">No invitations in last 7 days.</p>';

      let tableHtml = '<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;"><thead><tr><th style="padding: 8px; text-align: left; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">Date</th><th style="padding: 8px; text-align: right; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">Count</th></tr></thead><tbody>';
      filteredDates.forEach((date, index) => {
          tableHtml += `<tr><td style="padding: 10px 8px; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${date}</td><td style="padding: 10px 8px; text-align: right; font-size: 12px; color: #667eea; font-weight: 600; border-bottom: 1px solid #f5f5f5;">${dailyCounts[date]}</td></tr>`;
      });
      tableHtml += '</tbody></table>';
      return tableHtml;
  };

  const recruiterIdx = metrics.colIndices && metrics.colIndices.hasOwnProperty('Recruiter_name') ? metrics.colIndices['Recruiter_name'] : -1;
  const hasAdoptionData = adoptionChartData && adoptionChartData.recruiterAdoptionData && adoptionChartData.recruiterAdoptionData.length > 0;
  const hasMatchStarsColumnForAdoption = adoptionChartData && adoptionChartData.hasMatchStarsColumn;
  const adoptionScoreThreshold = adoptionChartData ? (adoptionChartData.matchScoreThreshold || APP_MATCH_SCORE_THRESHOLD_RB) : APP_MATCH_SCORE_THRESHOLD_RB;


  let html = `<!DOCTYPE html>
<html>
<head>
  <title>${VS_COMPANY_NAME_RB} AI Interview Daily Summary</title>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif; line-height: 1.5; color: #1a1a1a; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px; margin: 0; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%;">
  <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 900px;">
    <tr>
      <td align="center">
        <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #ffffff; border-radius: 16px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); margin: 0 auto; overflow: hidden;">
          <!-- Modern Header with Gradient -->
          <tr>
            <td style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 32px 24px; text-align: center;">
              <h1 style="color: #ffffff; font-size: 28px; font-weight: 700; margin: 0; letter-spacing: -0.5px; text-shadow: 0 2px 4px rgba(0,0,0,0.1);">AI Interview Daily Summary</h1>
              <p style="color: rgba(255,255,255,0.9); font-size: 14px; margin: 8px 0 0 0; font-weight: 400;">${VS_COMPANY_NAME_RB} • ${new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })}</p>
            </td>
          </tr>

          <!-- Modern KPI Cards -->
          <tr>
            <td style="padding: 24px;">
              <table border="0" cellpadding="0" cellspacing="12" width="100%" style="border-collapse: separate;">
                <tr>
                  <td width="25%" style="vertical-align: top; padding: 0;">
                    <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; padding: 20px; box-shadow: 0 4px 12px rgba(102, 126, 234, 0.3); text-align: center;">
                      <div style="font-size: 11px; font-weight: 600; color: rgba(255,255,255,0.9); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;">Invitations Sent</div>
                      <div style="font-size: 36px; font-weight: 700; color: #ffffff; line-height: 1;">${metrics.totalSent}</div>
                    </div>
                  </td>
                  <td width="25%" style="vertical-align: top; padding: 0;">
                    <div style="background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); border-radius: 12px; padding: 20px; box-shadow: 0 4px 12px rgba(245, 87, 108, 0.3); text-align: center;">
                      <div style="font-size: 11px; font-weight: 600; color: rgba(255,255,255,0.9); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;">Completion Rate</div>
                      <div style="font-size: 36px; font-weight: 700; color: #ffffff; line-height: 1;">${metrics.postSept10KpiCompletionRate}<span style="font-size: 20px;">%</span></div>
                    </div>
                    <div style="text-align: center; font-size: 9px; color: #999; margin-top: 6px;">Excl. <48hrs</div>
                  </td>
                  <td width="25%" style="vertical-align: top; padding: 0;">
                    <div style="background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); border-radius: 12px; padding: 20px; box-shadow: 0 4px 12px rgba(79, 172, 254, 0.3); text-align: center;">
                      <div style="font-size: 11px; font-weight: 600; color: rgba(255,255,255,0.9); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;">Avg Time</div>
                      <div style="font-size: 36px; font-weight: 700; color: #ffffff; line-height: 1;">${metrics.postSept10AvgTimeToScheduleDays !== null ? metrics.postSept10AvgTimeToScheduleDays : 'N/A'}<span style="font-size: 18px;">d</span></div>
                    </div>
                  </td>
                  <td width="25%" style="vertical-align: top; padding: 0;">
                    <div style="background: linear-gradient(135deg, #fa709a 0%, #fee140 100%); border-radius: 12px; padding: 20px; box-shadow: 0 4px 12px rgba(250, 112, 154, 0.3); text-align: center;">
                      <div style="font-size: 11px; font-weight: 600; color: rgba(255,255,255,0.9); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;">Match Stars</div>
                      <div style="font-size: 36px; font-weight: 700; color: #ffffff; line-height: 1;">${metrics.avgMatchStars !== null ? metrics.avgMatchStars : 'N/A'}</div>
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>



          <!-- Side-by-side Sections - Table Layout -->
          <tr>
            <td style="padding: 0 24px 24px;">
              <table border="0" cellpadding="0" cellspacing="12" width="100%" style="border-collapse: separate;">
                <tr>
                  <td width="50%" style="vertical-align: top; padding: 0;">
                    <div style="background: #ffffff; padding: 16px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
                      <div style="font-weight: 700; font-size: 14px; color: #1a1a1a; margin-bottom: 12px; letter-spacing: -0.2px;">Completion Status</div>
                      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
                         <thead><tr><th style="padding: 8px; text-align: left; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">Status</th><th style="padding: 8px; text-align: right; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">Count</th><th style="padding: 8px; text-align: right; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">%</th></tr></thead>
                         <tbody>
                         ${Object.entries(metrics.interviewStatusDistribution)
                                     .sort(([, dataA], [, dataB]) => dataB.count - dataA.count)
                                     .map(([status, data], index) => `<tr>
                                         <td style="padding: 10px 8px; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${status}</td>
                                         <td style="padding: 10px 8px; text-align: right; font-size: 12px; color: #1a1a1a; font-weight: 600; border-bottom: 1px solid #f5f5f5;">${data.count}</td>
                                         <td style="padding: 10px 8px; text-align: right; font-size: 12px; color: #667eea; font-weight: 600; border-bottom: 1px solid #f5f5f5;">${data.percentage}%</td>
                                     </tr>`).join('')}
                         </tbody>
                     </table>
                    </div>
                  </td>
                  <td width="50%" style="vertical-align: top; padding: 0;">
                     <div style="background: #ffffff; padding: 16px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
                       <div style="font-weight: 700; font-size: 14px; color: #1a1a1a; margin-bottom: 12px; letter-spacing: -0.2px;">Daily Invitations (7d)</div>
                       ${generateTimeseriesTable(metrics.dailySentCounts)}
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- Breakdown by Recruiter (Table) -->
          <tr>
            <td style="padding-top: 10px; padding-bottom: 10px;">
              <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-bottom: 16px;">
                 <div style="font-weight: 700; font-size: 15px; color: #1a1a1a; margin-bottom: 16px; letter-spacing: -0.3px;">Breakdown by Recruiter of the Position</div>
                 <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
             <thead>
                <tr>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">Recruiter</th>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">INVITATIONS<br>SENT</th>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">COMPLETED<br>INTERVIEWS</th>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">UPCOMING<br>SCHEDULED INTERVIEW</th>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">PENDING<br>CANDIDATE INTERVIEW</th>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">FEEDBACK SUBMITTED<br>BY CREATOR</th>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">CREATOR FEEDBACK<br>REVIEW PENDING</th>
                 </tr>
             </thead>
             <tbody>
                        ${(() => { // Start IIFE to contain logic
                            // Sort recruiters by Sent descending, keeping Unknown last
                            const sortedRecruiters = Object.entries(metrics.byRecruiter)
                                .sort(([recA, dataA], [recB, dataB]) => {
                          if (recA === 'Unknown') return 1;
                          if (recB === 'Unknown') return -1;
                                    return dataB.sent - dataA.sent;
                                });

                            // Generate table rows
                            return sortedRecruiters
                             .map(([rec, data], index) => {
                                return `
                                  <tr>
                                     <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 600; border-bottom: 1px solid #f5f5f5;">${rec}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 600; border-bottom: 1px solid #f5f5f5;">${data.sent}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.completedNumber} <span style="color: #667eea; font-size: 11px;">${data.completedPercentOfSent}%</span></td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.scheduled}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.pendingNumber} <span style="color: #667eea; font-size: 11px;">${data.pendingPercentOfSent}%</span></td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.feedbackSubmitted}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; border-bottom: 1px solid #f5f5f5;">
                                        ${data.recruiterSubmissionAwaited > 0 ? 
                                            `<span style="color: #f5576c; font-weight: 700;">${data.recruiterSubmissionAwaited}</span>` : 
                                            `<span style="color: #999;">${data.recruiterSubmissionAwaited}</span>`
                                        }
                                      </td>
                                  </tr>
                                `;
                            }).join('');
                        })()}
                          ${Object.keys(metrics.byRecruiter).length === 0 ? '<tr><td colspan="7" style="text-align:center; padding: 10px; border: 1px solid #e0e0e0; font-size: 12px;">No recruiter data found or Recruiter_name column missing.</td></tr>' : ''}
                      </tbody>
                  </table>
                 ${recruiterIdx === -1 ? '<p style="font-size: 0.85em; color: #757575; margin-top: 15px;">Recruiter breakdown is based on the "Recruiter_name" column, which was not found in the sheet.</p>' : ''}
              </div>
            </td>
          </tr>

          <!-- Breakdown by Creator (Table) -->
          <tr>
            <td style="padding: 0 24px 24px;">
              <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
                 <div style="font-weight: 700; font-size: 15px; color: #1a1a1a; margin-bottom: 16px; letter-spacing: -0.3px;">Breakdown by Creator</div>
                 <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
             <thead>
                <tr>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">Creator</th>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">INVITATIONS<br>SENT</th>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">COMPLETED<br>INTERVIEWS</th>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">UPCOMING<br>SCHEDULED INTERVIEW</th>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">PENDING<br>CANDIDATE INTERVIEW</th>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">FEEDBACK SUBMITTED<br>BY CREATOR</th>
                           <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">CREATOR FEEDBACK<br>REVIEW PENDING</th>
                 </tr>
             </thead>
             <tbody>
                        ${(() => { // Start IIFE to contain logic
                            // Sort creators by Sent descending, keeping Unknown last
                            const sortedCreators = Object.entries(metrics.byCreator)
                                .sort(([crtA, dataA], [crtB, dataB]) => {
                          if (crtA === 'Unknown') return 1;
                          if (crtB === 'Unknown') return -1;
                                    return dataB.sent - dataA.sent;
                                });

                            // Generate table rows
                            return sortedCreators
                             .map(([crt, data], index) => {
                                return `
                                  <tr>
                                     <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 600; border-bottom: 1px solid #f5f5f5;">${crt}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 600; border-bottom: 1px solid #f5f5f5;">${data.sent}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.completedNumber} <span style="color: #667eea; font-size: 11px;">${data.completedPercentOfSent}%</span></td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.scheduled}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.pendingNumber} <span style="color: #667eea; font-size: 11px;">${data.pendingPercentOfSent}%</span></td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.feedbackSubmitted}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; border-bottom: 1px solid #f5f5f5;">
                                        ${data.recruiterSubmissionAwaited > 0 ? 
                                            `<span style="color: #f5576c; font-weight: 700;">${data.recruiterSubmissionAwaited}</span>` : 
                                            `<span style="color: #999;">${data.recruiterSubmissionAwaited}</span>`
                                        }
                                      </td>
                                  </tr>
                                `;
                            }).join('');
                        })()}
                          ${Object.keys(metrics.byCreator).length === 0 ? '<tr><td colspan="7" style="text-align:center; padding: 10px; border: 1px solid #e0e0e0; font-size: 12px;">No creator data found or Creator_user_id column missing.</td></tr>' : ''}
                      </tbody>
                  </table>
                 ${(metrics.colIndices && metrics.colIndices.hasOwnProperty('Creator_user_id') ? metrics.colIndices['Creator_user_id'] : -1) === -1 ? '<p style="font-size: 0.85em; color: #757575; margin-top: 15px;">Creator breakdown is based on the "Creator_user_id" column, which was not found in the sheet.</p>' : ''}
              </div>
            </td>
          </tr>

          <!-- AI Interview Coverage Bar Chart -->
          ${aiCoverageMetrics ? `
          <tr>
            <td style="padding: 0 24px 24px;">
              ${generateAICoverageBarChartHtml(aiCoverageMetrics)}
            </td>
          </tr>
          ` : `
          <tr>
            <td style="padding: 0 24px 24px;">
              <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
                <div style="font-weight: 700; font-size: 15px; color: #1a1a1a; margin-bottom: 16px; letter-spacing: -0.3px;">AI Interview Coverage by Recruiter</div>
                <div style="text-align: center; padding: 60px 20px; color: #999; font-size: 14px; font-weight: 500;">Coming soon</div>
              </div>
            </td>
          </tr>
          `}

          <!-- Detailed Validation Sheets -->
          ${recruiterValidationSheets ? `
          <tr>
            <td style="padding: 0 24px 24px;">
              ${generateValidationSheetsHtml(recruiterValidationSheets)}
            </td>
          </tr>
          ` : `
          <tr>
            <td style="padding: 0 24px 24px;">
              <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
                <div style="font-weight: 700; font-size: 15px; color: #1a1a1a; margin-bottom: 16px; letter-spacing: -0.3px;">Detailed Validation Sheets</div>
                <div style="text-align: center; padding: 60px 20px; color: #999; font-size: 14px; font-weight: 500;">Coming soon</div>
              </div>
            </td>
          </tr>
          `}

          <!-- Creator Last Invite Activity -->
          <tr>
            <td style="padding: 0 24px 24px;">
              <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
                <div style="font-weight: 700; font-size: 15px; color: #1a1a1a; margin-bottom: 16px; letter-spacing: -0.3px;">Creator Last Invite Activity</div>
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
                  <thead>
                    <tr>
                      <th style="padding: 10px 8px; text-align: left; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">Creator</th>
                      <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">Last Sent</th>
                      <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">Trend (10d)</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${creatorActivityData && creatorActivityData.length > 0 ?
                        creatorActivityData.map((activity, index) => {
                            let daysAgoText = '';
                            if (activity.daysAgo === -1) {
                                daysAgoText = 'Today';
                            } else if (activity.daysAgo === 0) {
                                daysAgoText = 'Yesterday';
                            } else if (activity.daysAgo >= 1) {
                                const actualDays = activity.daysAgo + 1;
                                daysAgoText = `${actualDays}d ago`;
                            } else {
                                daysAgoText = 'Unknown';
                            }

                              return `
                              <tr>
                                <td style="padding: 12px 8px; font-size: 12px; color: #1a1a1a; font-weight: 600; border-bottom: 1px solid #f5f5f5;">${activity.creator}</td>
                                <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${daysAgoText}</td>
                                <td style="padding: 12px 8px; text-align: center; font-size: 11px; color: #667eea; font-weight: 500; font-family: 'SF Mono', Monaco, 'Cascadia Code', 'Roboto Mono', Consolas, 'Courier New', monospace; border-bottom: 1px solid #f5f5f5;">${activity.dailyTrend || 'N/A'}</td>
                              </tr>`;
                        }).join('')
                        :
                        '<tr><td colspan="3" style="text-align:center; padding: 20px; color: #999; font-size: 12px;">No creator activity data found.</td></tr>'
                    }
             </tbody>
         </table>
                 ${creatorIdx_Log === -1 ? '<p style="font-size: 11px; color: #999; margin-top: 12px;">Creator_user_id column not found.</p>' : ''}
     </div>
            </td>
          </tr>



     <!-- Breakdown by Job Function -->
          <tr>
             <td style="padding: 0 24px 24px;">
               <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
                  <div style="font-weight: 700; font-size: 15px; color: #1a1a1a; margin-bottom: 16px; letter-spacing: -0.3px;">Breakdown by Job Function</div>
                  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
             <thead>
                <tr>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">Job Function</th>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">INVITATIONS<br>SENT</th>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">COMPLETED<br>INTERVIEWS</th>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">UPCOMING<br>SCHEDULED INTERVIEW</th>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">PENDING<br>CANDIDATE INTERVIEW</th>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">FEEDBACK SUBMITTED<br>BY CREATOR</th>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">CREATOR FEEDBACK<br>REVIEW PENDING</th>
                 </tr>
             </thead>
             <tbody>
                 ${Object.entries(metrics.byJobFunction)
                     .sort(([funcA], [funcB]) => funcA.localeCompare(funcB))
                              .map(([func, data], index) => `
                                  <tr>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 600; border-bottom: 1px solid #f5f5f5;">${func}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 600; border-bottom: 1px solid #f5f5f5;">${data.sent}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.completedNumber} <span style="color: #667eea; font-size: 11px;">${data.completedPercentOfSent}%</span></td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.scheduled}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.pendingNumber} <span style="color: #667eea; font-size: 11px;">${data.pendingPercentOfSent}%</span></td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.feedbackSubmitted}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; border-bottom: 1px solid #f5f5f5;">
                                        ${data.recruiterSubmissionAwaited > 0 ? 
                                            `<span style="color: #f5576c; font-weight: 700;">${data.recruiterSubmissionAwaited}</span>` : 
                                            `<span style="color: #999;">${data.recruiterSubmissionAwaited}</span>`
                                        }
                                      </td>
                         </tr>
                     `).join('')}
             </tbody>
         </table>
     </div>
             </td>
          </tr>

     <!-- Breakdown by Location Country -->
           <tr>
             <td style="padding: 0 24px 24px;">
               <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
                  <div style="font-weight: 700; font-size: 15px; color: #1a1a1a; margin-bottom: 16px; letter-spacing: -0.3px;">Breakdown by Location Country</div>
                   <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
             <thead>
                <tr>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">Country</th>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">Sent</th>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">COMPLETED<br>INTERVIEWS</th>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">UPCOMING<br>SCHEDULED INTERVIEW</th>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">PENDING<br>CANDIDATE INTERVIEW</th>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">FEEDBACK SUBMITTED<br>BY CREATOR</th>
                            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0; line-height: 1.3;">CREATOR FEEDBACK<br>REVIEW PENDING</th>
                 </tr>
             </thead>
             <tbody>
                 ${Object.entries(metrics.byCountry)
                     .sort(([ctryA], [ctryB]) => ctryA.localeCompare(ctryB))
                              .map(([ctry, data], index) => `
                                  <tr>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 600; border-bottom: 1px solid #f5f5f5;">${ctry}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 600; border-bottom: 1px solid #f5f5f5;">${data.sent}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.completedNumber} <span style="color: #667eea; font-size: 11px;">${data.completedPercentOfSent}%</span></td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.scheduled}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.pendingNumber} <span style="color: #667eea; font-size: 11px;">${data.pendingPercentOfSent}%</span></td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.feedbackSubmitted}</td>
                                      <td style="padding: 12px 8px; text-align: center; font-size: 12px; border-bottom: 1px solid #f5f5f5;">
                                        ${data.recruiterSubmissionAwaited > 0 ? 
                                            `<span style="color: #f5576c; font-weight: 700;">${data.recruiterSubmissionAwaited}</span>` : 
                                            `<span style="color: #999;">${data.recruiterSubmissionAwaited}</span>`
                                        }
                                      </td>
                         </tr>
                     `).join('')}
             </tbody>
         </table>
     </div>
             </td>
          </tr>

          <!-- Modern Footer -->
          <tr>
            <td style="padding: 24px; background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); border-top: 1px solid #e0e0e0;">
              <div style="text-align: center; font-size: 10px; color: #666; line-height: 1.6;">
                <div style="margin-bottom: 4px;">*Avg Time uses Schedule Start Date as completion proxy (post-Sept 10, 2025)</div>
                <div style="margin-bottom: 4px;">**Completion Rate excludes invites sent < 48 hours. Table % includes all invites.</div>
                <div style="margin-bottom: 4px; font-weight: 600;">Overall completion rate: ${metrics.completionRateOriginal}%</div>
                <div style="margin-top: 8px; padding-top: 8px; border-top: 1px solid rgba(0,0,0,0.1); color: #999;">
                  Generated ${new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric', hour: '2-digit', minute: '2-digit' })} • ${Session.getScriptTimeZone()}
                </div>
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;

  return html;
}


/**
 * Sends an email with the recruiter breakdown report.
 * @param {string} recipient The primary email recipient.
 * @param {string} ccRecipient The CC email recipient (can be empty).
 * @param {string} subject The email subject.
 * @param {string} htmlBody The HTML content of the email.
 */
function sendVsEmailRB(recipient, ccRecipient, subject, htmlBody) {
  if (!recipient) {
     Logger.log("ERROR: Email recipient (RB) is empty. Cannot send email.");
     return;
  }
   if (!subject) {
     Logger.log("WARNING: Email subject (RB) is empty. Using default subject.");
     subject = `${VS_COMPANY_NAME_RB} AI Interview Recruiter Report`;
  }
   if (!htmlBody) {
     Logger.log("ERROR: Email body (RB) is empty. Cannot send email.");
     return;
  }

  const options = {
     to: recipient,
     subject: subject,
     htmlBody: htmlBody
  };

  if (ccRecipient && ccRecipient.trim() !== '' && ccRecipient.trim().toLowerCase() !== recipient.trim().toLowerCase()) {
    options.cc = ccRecipient;
    Logger.log(`Sending recruiter report email to ${recipient}, CC ${ccRecipient}`);
  } else {
     Logger.log(`Sending recruiter report email to ${recipient} (No CC or CC is same as recipient)`);
  }

  try {
      MailApp.sendEmail(options);
      Logger.log("Recruiter report email sent successfully.");
  } catch (e) {
     Logger.log(`ERROR sending recruiter report email: ${e.toString()}`);
     sendVsErrorNotificationRB(`CRITICAL: Failed to send recruiter report email to ${recipient}`, `Error: ${e.toString()}`); // Use RB notifier
  }
}

/**
 * Sends an error notification email for the Recruiter Breakdown script.
 * @param {string} errorMessage The main error message.
 * @param {string} [stackTrace=''] Optional stack trace.
 */
function sendVsErrorNotificationRB(errorMessage, stackTrace = '') {
   const recipient = VS_EMAIL_RECIPIENT_RB; // Use RB config
   if (!recipient) {
       Logger.log("CRITICAL ERROR: Cannot send error notification (RB) because VS_EMAIL_RECIPIENT_RB is not set.");
       return;
   }
   try {
       const subject = `ERROR: ${VS_COMPANY_NAME_RB} AI Recruiter Report Failed - ${new Date().toLocaleString()}`;
       let body = `Error generating/sending ${VS_COMPANY_NAME_RB} AI Interview Recruiter Report:

${errorMessage}

`;
       if (stackTrace) {
           body += `Stack Trace:
${stackTrace}

`;
       }
       body += `Log Sheet URL: ${VS_LOG_SHEET_SPREADSHEET_URL_RB}`; // Use RB config
       MailApp.sendEmail(recipient, subject, body);
       Logger.log(`Error notification email (RB) sent to ${recipient}.`);
    } catch (emailError) {
       Logger.log(`CRITICAL: Failed to send error notification email (RB) to ${recipient}: ${emailError}`);
    }
}


// --- Utility / Setup Functions ---

/**
 * Creates menu items for the Recruiter Breakdown report.
 */
function setupRecruiterBreakdownMenu() {
  try {
    SpreadsheetApp.getUi()
      .createMenu(`${VS_COMPANY_NAME_RB} AI Daily Summary`) // Menu Name Updated
      .addItem('Generate & Send Summary Now', 'AIR_DailySummarytoAP') // Updated Item Text & Function Name
      .addItem('Schedule Daily Summary (10 AM)', 'createRecruiterBreakdownTrigger') // Updated Item Text
      .addToUi();
  } catch (e) {
    Logger.log("Error creating Daily Summary menu (might happen if not opened from a Sheet): " + e); // Updated log
  }
}

// --- Helper Functions (Renamed with RB suffix for clarity, logic may be identical) ---
/**
 * Parses date strings safely.
 * @param {any} dateInput Input value.
 * @returns {Date|null} Parsed Date object or null.
 */
function vsParseDateSafeRB(dateInput) {
    if (dateInput === null || dateInput === undefined || dateInput === '') return null;
    if (typeof dateInput === 'number' && dateInput > 10000) {
       try {
           const jsTimestamp = (dateInput - 25569) * 86400 * 1000;
           const date = new Date(jsTimestamp);
            return !isNaN(date.getTime()) ? date : null;
       } catch (e) { /* Ignore */ }
    }
    const date = new Date(dateInput);
    return !isNaN(date.getTime()) ? date : null;
}

/**
 * Calculates time difference in days.
 * @param {Date|null} date1 Earlier date.
 * @param {Date|null} date2 Later date.
 * @returns {number|null} Difference in days or null.
 */
function vsCalculateDaysDifferenceRB(date1, date2) {
    if (!date1 || !date2) return null;
    const diffTime = date2.getTime() - date1.getTime();
    if (diffTime < 0) return null;
    return diffTime / (1000 * 60 * 60 * 24);
}

/**
 * Formats a Date object into DD-MMM-YY.
 * @param {Date|null} dateObject Date to format.
 * @returns {string} Formatted date string or 'N/A'.
 */
function vsFormatDateRB(dateObject) {
    if (!dateObject || !(dateObject instanceof Date) || isNaN(dateObject.getTime())) return 'N/A';
    const day = String(dateObject.getDate()).padStart(2, '0');
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const month = monthNames[dateObject.getMonth()];
    const year = String(dateObject.getFullYear()).slice(-2);
    return `${day}-${month}-${year}`;
}

/**
 * Assigns a numerical rank to interview statuses for prioritization.
 * @param {string} status Raw interview status.
 * @returns {number} Rank.
 */
function vsGetStatusRankRB(status) {
    const COMPLETED_STATUSES_RAW = ['COMPLETED', 'Feedback Provided', 'Pending Feedback', 'No Show'];
    const SCHEDULED_STATUS_RAW = 'SCHEDULED';
    const PENDING_STATUSES_RAW = ['PENDING', 'INVITED', 'EMAIL SENT'];

    if (COMPLETED_STATUSES_RAW.includes(status)) return 1;
    if (status === SCHEDULED_STATUS_RAW) return 2;
    if (PENDING_STATUSES_RAW.includes(status)) return 3;
    return 99;
}

//====================================================================================================
// --- Insight Formatting Helper (DISABLED - AI Insights Feature Removed) ---
//====================================================================================================
/**
 * Formats AI-generated insights text for HTML display by bolding known entities and percentages.
 * @param {string} insightsText The raw text insights from the LLM.
 * @param {object} metrics The main metrics object to extract known entity names.
 * @return {string} HTML formatted insights string.
 * @deprecated AI Insights feature has been removed from the email report
 */
function formatInsightsForHtml_DISABLED(insightsText, metrics) {
  if (!insightsText || typeof insightsText !== 'string') {
    return 'No insights available or an error occurred.';
  }

  let formattedText = insightsText;

  // 1. Collect known entities (recruiters, job functions, countries)
  let knownEntities = [];
  if (metrics) {
    if (metrics.byRecruiter) {
      knownEntities.push(...Object.keys(metrics.byRecruiter).filter(name => name && name.trim() !== '' && name !== 'Unknown' && name !== 'Unassigned'));
    }
    if (metrics.byJobFunction) {
      knownEntities.push(...Object.keys(metrics.byJobFunction).filter(name => name && name.trim() !== '' && name !== 'Unknown' && name !== 'Unassigned'));
    }
    if (metrics.byCountry) {
      knownEntities.push(...Object.keys(metrics.byCountry).filter(name => name && name.trim() !== '' && name !== 'Unknown' && name !== 'Unassigned'));
    }
    // Add any other specific entities you want to bold if they are reliably named.
    // Example: If you have a list of project names or specific terms stored elsewhere in metrics.
  }
  
  // Remove duplicates and filter out very short strings (e.g., single characters unless specifically desired)
  knownEntities = [...new Set(knownEntities)].filter(e => e.length > 2); // Min length 3 for an entity name
  // Sort by length descending. Crucial for correct replacement of substrings.
  knownEntities.sort((a, b) => b.length - a.length);

  // 2. Bold known entities (done first)
  knownEntities.forEach(entity => {
    // Escape special regex characters in the entity name itself
    const escapedEntity = entity.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    try {
      // Case-insensitive ('i') and global ('g') match, with word boundaries (\b)
      const regex = new RegExp(`\\b(${escapedEntity})\\b`, 'gi'); 
      formattedText = formattedText.replace(regex, '<strong>$1</strong>');
    } catch (e) {
      Logger.log(`Error creating or using regex for entity "${entity}": ${e.toString()}`);
      // Continue without bolding this specific entity if regex fails
    }
  });

  // 3. Bold percentages (e.g., 50.7%, 79.4%, 20%) - done after entities
  // This regex looks for digits, optionally with a decimal, followed by a % sign,
  // ensuring it's a whole word/number using word boundaries.
  formattedText = formattedText.replace(/(\b\d+(\.\d+)?%\b)/g, '<strong>$1</strong>');

  // 4. Replace newlines with <br> for HTML display (done last)
  formattedText = formattedText.replace(/\n/g, '<br>');

  return formattedText;
}

//====================================================================================================
// --- Gemini API Integration for Insights (DISABLED - AI Insights Feature Removed) ---
//====================================================================================================

/**
 * Fetches insights from the Google Gemini API based on summarized report data.
 * @deprecated AI Insights feature has been removed from the email report
 * @param {object} metrics The calculated company metrics.
 * @param {object} adoptionChartData The adoption chart data.
 * @param {Array<object>} recruiterActivityData Recruiter activity data.
 * @return {string} Textual insights from the LLM, or an error/status message.
 */
function fetchInsightsFromGeminiAPI_DISABLED(metrics, adoptionChartData, recruiterActivityData) {
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  if (!GEMINI_API_KEY) {
    Logger.log("ERROR: GEMINI_API_KEY not found in Script Properties. AI Insights will not be generated.");
    return "AI Insights could not be generated: API Key not configured in Script Properties.";
  }
  // Using gemini-1.5-flash-latest as an example for a fast and capable model.
  const GEMINI_API_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${GEMINI_API_KEY}`;

  try {
    // 1. Summarize your data
    let dataSummary = `AI Interview Report Summary (Data for ${VS_COMPANY_NAME_RB}):\n\n`; // VS_COMPANY_NAME_RB should be globally accessible
    dataSummary += `Overall Performance (${metrics.reportStartDate} to ${metrics.reportEndDate}):\n`;
    dataSummary += `- Total AI Invitations Sent: ${metrics.totalSent}\n`;
    dataSummary += `- Overall Completion Rate (KPI Adjusted, for invites older than 48 hours): ${metrics.kpiCompletionRateAdjusted}%\n`;
    dataSummary += `- Overall Completion Rate (All Time): ${metrics.completionRateOriginal}%\n`;
    dataSummary += `- Average Time Sent to Completion (Post-Sept 10th, Proxy: Schedule Start): ${metrics.postSept10AvgTimeToScheduleDays !== null ? metrics.postSept10AvgTimeToScheduleDays + ' days' : 'N/A'}\n`;
    dataSummary += `- Average Match Stars (for Completed Interviews): ${metrics.avgMatchStars !== null ? metrics.avgMatchStars : 'N/A'}\n`;

    // Revised Recruiter Performance Summary to consider all recruiters for insights
    if (metrics.byRecruiter && Object.keys(metrics.byRecruiter).length > 0) {
      const recruiterStats = Object.entries(metrics.byRecruiter)
        .filter(([name]) => name !== 'Unknown') // Exclude 'Unknown' for specific high/low stats
        .map(([name, data]) => ({
          name: name,
          sent: parseInt(data.sent) || 0,
          completed: parseInt(data.completedNumber) || 0,
          completionRate: parseFloat(data.completedPercentOfSent) || 0,
          pending: parseInt(data.pendingNumber) || 0
        }));

      if (recruiterStats.length > 0) {
        dataSummary += `\nRecruiter Performance Analysis (Based on ${recruiterStats.length} recruiters with known names):
`;

        // Sort by completion rate for high/low
        const sortedByCompletionRate = [...recruiterStats].sort((a, b) => b.completionRate - a.completionRate);
        dataSummary += `- Highest Completion Rate: ${sortedByCompletionRate[0].name} (${sortedByCompletionRate[0].completionRate}% from ${sortedByCompletionRate[0].sent} sent invites).
`;

        // For lowest, find someone with a minimum number of sent invites to make it meaningful
        const minSentForLowConsideration = 5; // Adjustable threshold
        const eligibleForLowest = sortedByCompletionRate.filter(r => r.sent >= minSentForLowConsideration);
        if (eligibleForLowest.length > 0) {
          const lowestPerformer = eligibleForLowest[eligibleForLowest.length - 1];
          if (lowestPerformer.name !== sortedByCompletionRate[0].name) { // Avoid repeating if only one eligible
            dataSummary += `- Lowest Completion Rate (among those with >=${minSentForLowConsideration} invites): ${lowestPerformer.name} (${lowestPerformer.completionRate}% from ${lowestPerformer.sent} sent invites).
`;
          }
        }

        // Sort by sent for most active
        const sortedBySent = [...recruiterStats].sort((a, b) => b.sent - a.sent);
        dataSummary += `- Most Invites Sent: ${sortedBySent[0].name} (${sortedBySent[0].sent} invites, ${sortedBySent[0].completionRate}% completion rate).
`;

        // Calculate overall average completion rate for this group
        let totalSentByKnownRecruiters = 0;
        let totalCompletedByKnownRecruiters = 0;
        recruiterStats.forEach(r => {
          totalSentByKnownRecruiters += r.sent;
          totalCompletedByKnownRecruiters += r.completed;
        });
        const avgCompletionRateKnown = totalSentByKnownRecruiters > 0 ? parseFloat(((totalCompletedByKnownRecruiters / totalSentByKnownRecruiters) * 100).toFixed(1)) : 0;
        dataSummary += `- Average Completion Rate (for these ${recruiterStats.length} recruiters): ${avgCompletionRateKnown}%\n`;
      }

      if (metrics.byRecruiter['Unknown'] && metrics.byRecruiter['Unknown'].sent > 0) {
        dataSummary += `- Invites Sent by 'Unknown' Recruiters: ${metrics.byRecruiter['Unknown'].sent} (completion rate: ${metrics.byRecruiter['Unknown'].completedPercentOfSent}%).\n`;
      }
    }

    // Enhanced Recruiter Last AI Invite Activity Summary
    if (recruiterActivityData && recruiterActivityData.length > 0) {
        dataSummary += `\nRecruiter Last AI Invite Activity:\n`;
        const recentActivityDisplayCount = 3; // How many most recent to display

        recruiterActivityData.slice(0, recentActivityDisplayCount).forEach(activity => {
            let daysAgoText = "";
            if (activity.daysAgo === 0) daysAgoText = "Today";
            else if (activity.daysAgo === 1) daysAgoText = "Yesterday";
            else daysAgoText = `${activity.daysAgo} days ago`;
            dataSummary += `- Recently Active: ${activity.recruiter} (Last AI invite: ${daysAgoText}, 10-day trend: ${activity.dailyTrend}).\n`;
        });

        const inactivityThresholdDays = 7; // Recruiters with no AI invites for this many days or more
        // Note: activity.daysAgo = 0 is today, 1 is yesterday. So >= 7 means 7 full days have passed (i.e., last activity was 7+ days ago)
        const lessActiveRecruiters = recruiterActivityData.filter(activity => activity.daysAgo >= inactivityThresholdDays);

        if (lessActiveRecruiters.length > 0) {
            dataSummary += `\n- Recruiters with Notably Low Recent AI Invite Activity (Last AI invite >= ${inactivityThresholdDays} days ago):\n`;
            lessActiveRecruiters.forEach(activity => {
                 let daysAgoText = `${activity.daysAgo} days ago`; // Simplified for this section
                 dataSummary += `  - ${activity.recruiter} (Last AI invite: ${daysAgoText}).\n`;
            });
        }
    }



    if (recruiterActivityData && recruiterActivityData.length > 0) {
        dataSummary += `\nRecruiter Last Invite Activity (Most Recent First):\n`;
        recruiterActivityData.slice(0,3).forEach(activity => {
            let daysAgoText = '';
            if (activity.daysAgo === -1) daysAgoText = 'Today';
            else if (activity.daysAgo === 0) daysAgoText = 'Yesterday';
            else if (activity.daysAgo >= 1) daysAgoText = `${activity.daysAgo + 1} calendar days ago`;
            else daysAgoText = 'Unknown';
            dataSummary += `- ${activity.recruiter}: Last invite sent ${daysAgoText}. Trend (last 10 workdays, excl. weekends): ${activity.dailyTrend}\n`;
        });
    }

    // Revised Performance by Job Function Summary
    if (metrics.byJobFunction && Object.keys(metrics.byJobFunction).length > 0) {
        const jobFunctionStats = Object.entries(metrics.byJobFunction)
            .map(([funcName, data]) => ({
                name: funcName,
                sent: parseInt(data.sent) || 0,
                completed: parseInt(data.completedNumber) || 0,
                completionRate: parseFloat(data.completedPercentOfSent) || 0,
                // recruiterSubmissionAwaited: parseInt(data.recruiterSubmissionAwaited) || 0 // If needed later
            }));

        if (jobFunctionStats.length > 0) {
            dataSummary += `\nPerformance by Job Function (Analysis based on ${jobFunctionStats.length} functions):
`;
            const minSentForConsideration = 5; // Job functions with at least this many sent invites considered for low/high
            const relevantJobFunctions = jobFunctionStats.filter(jf => jf.sent >= minSentForConsideration);

            if (relevantJobFunctions.length > 0) {
                // Sort by completion rate
                const sortedByCompletion = [...relevantJobFunctions].sort((a, b) => b.completionRate - a.completionRate);

                dataSummary += `- Highest Completion Rate (among functions with >=${minSentForConsideration} invites): ${sortedByCompletion[0].name} (${sortedByCompletion[0].completionRate}% from ${sortedByCompletion[0].sent} sent).
`;

                if (sortedByCompletion.length > 1) {
                    const lowestPerformer = sortedByCompletion[sortedByCompletion.length - 1];
                     if (lowestPerformer.name !== sortedByCompletion[0].name) { // Avoid repeating if only one eligible
                        dataSummary += `- Lowest Completion Rate (among functions with >=${minSentForConsideration} invites): ${lowestPerformer.name} (${lowestPerformer.completionRate}% from ${lowestPerformer.sent} sent).
`;
                    }
                }
                 // Add one or two other notable ones if they exist and are different, e.g., highest volume
                const sortedBySent = [...jobFunctionStats].sort((a,b) => b.sent - a.sent);
                if (sortedBySent.length > 0 && sortedBySent[0].name !== sortedByCompletion[0].name && (sortedByCompletion.length <=1 || sortedBySent[0].name !== sortedByCompletion[sortedByCompletion.length-1].name)) {
                     dataSummary += `- Highest Volume of Invites: ${sortedBySent[0].name} (${sortedBySent[0].sent} sent, ${sortedBySent[0].completionRate}% completion rate).\n`;
                }

                // Specifically add Engineering stats if it exists and meets threshold, or just add it if it exists
                const engineeringStats = jobFunctionStats.find(jf => jf.name.toLowerCase() === 'engineering');
                if (engineeringStats) {
                    dataSummary += `- Specific Stats for Engineering: ${engineeringStats.sent} invites sent, ${engineeringStats.completionRate}% completion rate.\n`;
                }

            } else {
                dataSummary += `- Insufficient data for detailed job function comparison (few functions with >=${minSentForConsideration} invites sent).
`;
            }
        }
    }

    // Performance by Country Summary
    if (metrics.byCountry && Object.keys(metrics.byCountry).length > 0) {
        const countryStats = Object.entries(metrics.byCountry)
            .map(([countryName, data]) => ({
                name: countryName,
                sent: parseInt(data.sent) || 0,
                completed: parseInt(data.completedNumber) || 0,
                completionRate: parseFloat(data.completedPercentOfSent) || 0,
            }));

        if (countryStats.length > 0) {
            dataSummary += `\nPerformance by Country (Analysis based on ${countryStats.length} countries):
`;
            const minSentForCountryConsideration = 5; // Countries with at least this many sent invites
            const relevantCountries = countryStats.filter(c => c.sent >= minSentForCountryConsideration);

            if (relevantCountries.length > 0) {
                const sortedByCompletion = [...relevantCountries].sort((a, b) => b.completionRate - a.completionRate);
                dataSummary += `- Country with Highest Completion Rate (>=${minSentForCountryConsideration} invites): ${sortedByCompletion[0].name} (${sortedByCompletion[0].completionRate}% from ${sortedByCompletion[0].sent} sent).
`;
                if (sortedByCompletion.length > 1) {
                    const lowestPerformer = sortedByCompletion[sortedByCompletion.length - 1];
                    if (lowestPerformer.name !== sortedByCompletion[0].name) {
                        dataSummary += `- Country with Lowest Completion Rate (>=${minSentForCountryConsideration} invites): ${lowestPerformer.name} (${lowestPerformer.completionRate}% from ${lowestPerformer.sent} sent).
`;
                    }
                }
                const sortedBySent = [...countryStats].sort((a,b) => b.sent - a.sent);
                if (sortedBySent.length > 0 && sortedBySent[0].name !== sortedByCompletion[0].name && (sortedByCompletion.length <= 1 || sortedBySent[0].name !== sortedByCompletion[sortedByCompletion.length-1].name)) {
                    dataSummary += `- Country with Highest Volume of Invites: ${sortedBySent[0].name} (${sortedBySent[0].sent} sent, ${sortedBySent[0].completionRate}% completion rate).
`;
                }
            } else {
                dataSummary += `- Insufficient data for detailed country comparison (few countries with >=${minSentForCountryConsideration} invites sent).
`;
            }
        }
    }
    
    dataSummary += "\nConsiderations: 'Completion Rate (KPI Adjusted)' excludes invites sent in the last 48 hours. 'Avg Time Sent to Completion' uses schedule start time as a proxy for completion. Adoption metrics are based on a specific cohort post-launch with a match score filter.\n";

    // 2. Construct the Prompt
    const promptText = `You are an expert data analyst for an HR department. Your task is to provide insightful commentary on a report summarizing AI interview adoption and recruiter activity for ${VS_COMPANY_NAME_RB}.
Based *only* on the following data summary, please generate 3 to 5 bullet-point key observations.
For each observation, identify the primary subject (e.g., overall performance, a specific job function, recruiter activity, country performance, AI adoption).
When discussing a specific category from the summary (like Recruiter Performance, Job Functions, Countries, or Adoption):
- Attempt to highlight the most significant variations by contrasting the best and worst performers (e.g., highest vs. lowest completion rates, most vs. least active) using the specific names and figures provided in the summary for that category.
- Call out specific names (recruiters, job functions, countries) when discussing these variations if the data supports it.
Focus on actionable insights, positive trends, areas that might need attention, or interesting patterns. 
Do not refer to data not present in the summary. Be concise, ensuring each bullet point is a distinct observation.
If data for the "Engineering" job function is present in the summary, please include one bullet point observation specifically about Engineering performance.

Data Summary:
---
${dataSummary}
---

Key Observations (3-5 bullet points):
`;

    const payload = {
      contents: [{ role: "user", parts: [{text: promptText}] }],
      generationConfig: {
        "temperature": 0.6,
        "maxOutputTokens": 500,
        "topP": 0.9,
        "topK": 40
      },
       safetySettings: [
        { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_MEDIUM_AND_ABOVE" },
        { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_MEDIUM_AND_ABOVE" },
        { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_MEDIUM_AND_ABOVE" },
        { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_MEDIUM_AND_ABOVE" }
      ]
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    Logger.log(`Sending data to Gemini API. Endpoint: ${GEMINI_API_ENDPOINT}. Summary length: ${dataSummary.length} chars.`);
    const response = UrlFetchApp.fetch(GEMINI_API_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      if (jsonResponse.candidates && jsonResponse.candidates.length > 0 &&
          jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts &&
          jsonResponse.candidates[0].content.parts.length > 0 && jsonResponse.candidates[0].content.parts[0].text) {
        
        let insights = jsonResponse.candidates[0].content.parts[0].text;
        Logger.log("Successfully received insights from Gemini API.");
        insights = insights.trim();
        // Attempt to format into bullet points if the model didn't already.
        if (!insights.startsWith("* ") && !insights.startsWith("- ") && !insights.startsWith("• ")) {
            insights = insights.split('\n').map(line => line.trim()).filter(line => line.length > 0).map(line => `• ${line}`).join('\n');
        }
        return insights;
      } else if (jsonResponse.candidates && jsonResponse.candidates[0].finishReason) {
         Logger.log(`Gemini API call finished with reason: ${jsonResponse.candidates[0].finishReason}.`);
         let detail = jsonResponse.candidates[0].finishReason;
         if(jsonResponse.candidates[0].safetyRatings) {
           detail += ` Safety Ratings: ${JSON.stringify(jsonResponse.candidates[0].safetyRatings)}`;
         }
         return `AI Insights generation issue: Model finished with reason '${detail}'. Content might be blocked or prompt needs adjustment.`;
      } else {
        Logger.log(`Gemini API response does not contain expected text data. Response: ${responseBody}`);
        return "AI Insights generation failed: Unexpected API response structure. Check logs.";
      }
    } else {
      Logger.log(`Error calling Gemini API: ${responseCode} - ${responseBody}`);
      return `Could not generate AI insights: API Error ${responseCode}. Details: ${responseBody.substring(0, 500)}. Check logs for full error.`;
    }
  } catch (error) {
    Logger.log(`Critical Error in fetchInsightsFromGeminiAPI: ${error.toString()} \nStack: ${error.stack}`);
    return `Could not generate AI insights due to an internal script error: ${error.message}. Check logs.`;
  }
}

/**
 * Sends an error notification email for the Recruiter Breakdown script.
 * @param {string} errorMessage The main error message.
 * @param {string} [stackTrace=''] Optional stack trace.
 */
function sendVsErrorNotificationRB(errorMessage, stackTrace = '') {
  const recipient = VS_EMAIL_RECIPIENT_RB; // Use RB config
  if (!recipient) {
      Logger.log("CRITICAL ERROR: Cannot send error notification (RB) because VS_EMAIL_RECIPIENT_RB is not set.");
      return;
  }
  try {
      const subject = `ERROR: ${VS_COMPANY_NAME_RB} AI Recruiter Report Failed - ${new Date().toLocaleString()}`;
      let body = `Error generating/sending ${VS_COMPANY_NAME_RB} AI Interview Recruiter Report:\n\n${errorMessage}\n\n`;
      if (stackTrace) {
          body += `Stack Trace:\n${stackTrace}\n\n`;
      }
      body += `Log Sheet URL: ${VS_LOG_SHEET_SPREADSHEET_URL_RB}`; // Use RB config
      MailApp.sendEmail(recipient, subject, body);
      Logger.log(`Error notification email (RB) sent to ${recipient}.`);
   } catch (emailError) {
      Logger.log(`CRITICAL: Failed to send error notification email (RB) to ${recipient}: ${emailError}`);
   }
}

/**
 * Calculates hiring metrics from application data for candidates who took AI interviews.
 * @param {Array<Array>} appRows Raw rows from the application sheet.
 * @param {object} appColIndices Column indices map for the application sheet.
 * @returns {object} An object containing hiring metrics.
 */
function calculateHiringMetricsFromAppData(appRows, appColIndices) {
  Logger.log(`--- Starting calculateHiringMetricsFromAppData ---`);

  const hiringStages = [
    'Offer Approvals', 'Offer Extended', 'Offer Declined', 'Pending Start', 'Hired'
  ];

  // ---- TOP LEVEL METRICS CALCULATION ----
  const topLevelHiringCandidates = appRows.filter(row => {
    const lastStage = row[appColIndices['Last_stage']];
    return lastStage && hiringStages.includes(lastStage);
  });
  Logger.log(`Found ${topLevelHiringCandidates.length} total candidates who reached hiring stages`);

  const recruiterNameIndex = appColIndices.hasOwnProperty('Recruiter name') ? appColIndices['Recruiter name'] : -1;
  const filteredHiringCandidates = topLevelHiringCandidates.filter(row => {
    if (recruiterNameIndex === -1) return true;
    const recruiterName = row[recruiterNameIndex] || '';
    return !recruiterName.toLowerCase().includes('samrudh') && !recruiterName.toLowerCase().includes('simran');
  });
  Logger.log(`After excluding Samrudh/Simran recruiters: ${filteredHiringCandidates.length} candidates at offer stage`);

  const aiCandidates = filteredHiringCandidates.filter(row => (row[appColIndices['Ai_interview']] || '') === 'Y');
  const nonAiCandidates = filteredHiringCandidates.filter(row => (row[appColIndices['Ai_interview']] || '') !== 'Y');
  Logger.log(`AI candidates at offer stage: ${aiCandidates.length}, Non-AI candidates: ${nonAiCandidates.length}`);

  const positionIdIndex = appColIndices.hasOwnProperty('Position_id') ? appColIndices['Position_id'] : -1;
  const uniquePositions = new Set(aiCandidates.map(row => row[positionIdIndex]).filter(id => id));

  const matchStarsIndex = appColIndices.hasOwnProperty('Match_stars') ? appColIndices['Match_stars'] : -1;
  let aiAvgMatchScore = null, nonAiAvgMatchScore = null;
  if (matchStarsIndex !== -1) {
    const getAvgScore = (candidates) => {
      const scores = candidates.map(row => parseFloat(row[matchStarsIndex])).filter(score => !isNaN(score) && score >= 0);
      return scores.length > 0 ? parseFloat((scores.reduce((sum, score) => sum + score, 0) / scores.length).toFixed(1)) : null;
    };
    aiAvgMatchScore = getAvgScore(aiCandidates);
    nonAiAvgMatchScore = getAvgScore(nonAiCandidates);
  }

  // ---- NUANCED CALCULATION for "Average Candidates Needed" ----
  let aiAvgCandidatesNeeded = null, nonAiAvgCandidatesNeeded = null;
  try {
    const candidatesForRatio = appRows.filter(row => {
      if (recruiterNameIndex === -1) return true;
      const recruiterName = row[recruiterNameIndex] || '';
      return !recruiterName.toLowerCase().includes('samrudh') && !recruiterName.toLowerCase().includes('simran');
    });

    const positionStats = {};
    const progressedStages = [
      'Hiring Manager Screen', 'Assessment', 'Onsite Interview', 'Final Interview', 'Candidate Withdrew', 'Candidate Hold',
      'Offer Approvals', 'Offer Extended', 'Offer Declined', 'Pending Start', 'Hired'
    ];
    
    candidatesForRatio.forEach(row => {
      const posId = row[positionIdIndex];
      if (!posId) return;
      if (!positionStats[posId]) {
        positionStats[posId] = { ai_progressed: 0, ai_offered: 0, non_ai_progressed: 0, non_ai_offered: 0, had_ai_interview: false };
      }
      const stats = positionStats[posId];
      const lastStage = row[appColIndices['Last_stage']] || '';
      const aiInterview = row[appColIndices['Ai_interview']] || '';
      
      const hasProgressed = progressedStages.includes(lastStage);
      
      if (aiInterview === 'Y') {
        stats.had_ai_interview = true;
        if (hasProgressed) {
          stats.ai_progressed++;
          if (hiringStages.includes(lastStage)) stats.ai_offered++;
        }
      } else {
        if (hasProgressed) {
          stats.non_ai_progressed++;
          if (hiringStages.includes(lastStage)) stats.non_ai_offered++;
        }
      }
    });

    const aiPositionRatios = [], nonAiPositionRatios = [];
    for (const posId in positionStats) {
      const stats = positionStats[posId];
      if (stats.had_ai_interview) {
        // For AI positions, only require that an offer was made.
        if (stats.ai_offered > 0) {
          aiPositionRatios.push(stats.ai_progressed / stats.ai_offered);
        }
      } else {
        // For Non-AI positions, keep the significance threshold.
        if (stats.non_ai_progressed >= 3 && stats.non_ai_offered > 0) {
          nonAiPositionRatios.push(stats.non_ai_progressed / stats.non_ai_offered);
        }
      }
    }

    if (aiPositionRatios.length > 0) aiAvgCandidatesNeeded = parseFloat((aiPositionRatios.reduce((a, b) => a + b, 0) / aiPositionRatios.length).toFixed(1));
    if (nonAiPositionRatios.length > 0) nonAiAvgCandidatesNeeded = parseFloat((nonAiPositionRatios.reduce((a, b) => a + b, 0) / nonAiPositionRatios.length).toFixed(1));
    Logger.log(`Calculated avg candidates needed. AI Ratios Count: ${aiPositionRatios.length}, Non-AI Ratios Count: ${nonAiPositionRatios.length}`);
  } catch (e) {
    Logger.log(`ERROR during nuanced average candidate calculation: ${e}`);
  }
  
  // ---- Final Metrics Object ----
  const positionApprovedIndex = appColIndices.hasOwnProperty('Position approved date') ? appColIndices['Position approved date'] : -1;
  const metrics = {
    totalCandidates: aiCandidates.length,
    uniquePositionsFilled: uniquePositions.size,
    stageBreakdown: hiringStages.reduce((acc, stage) => { acc[stage] = aiCandidates.filter(row => row[appColIndices['Last_stage']] === stage).length; return acc; }, {}),
    hasPositionIdColumn: positionIdIndex !== -1,
    aiAvgMatchScore, nonAiAvgMatchScore,
    hasMatchStarsColumn: matchStarsIndex !== -1,
    aiCandidatesCount: aiCandidates.length,
    nonAiCandidatesCount: nonAiCandidates.length,
    aiAvgCandidatesNeeded, nonAiAvgCandidatesNeeded,
    hasPositionApprovedDateColumn: positionApprovedIndex !== -1
  };

  Logger.log(`Hiring Metrics: ${metrics.totalCandidates} AI candidates reached offer stage, ${metrics.uniquePositionsFilled} unique positions filled`);
  Logger.log(`Match Score Comparison: AI avg=${aiAvgMatchScore}, Non-AI avg=${nonAiAvgMatchScore}`);
  return metrics;
}

/**
 * Creates a validation Google Sheet with detailed position-level data for candidate count comparison.
 * @param {Array<Array>} appRows Raw rows from the application sheet.
 * @param {object} appColIndices Column indices map for the application sheet.
 * @returns {string} The URL of the created Google Sheet.
 */
function createCandidateCountValidationSheet(appRows, appColIndices) {
  Logger.log(`--- Starting createCandidateCountValidationSheet ---`);
  
  try {
    const spreadsheet = SpreadsheetApp.create(`AI Interview Candidate Count Validation - ${new Date().toISOString().split('T')[0]}`);
    const sheet = spreadsheet.getActiveSheet();
    
    const headers = [
      'Position ID', 'Position Title', 'Recruiter Name', 'Position Approved Date', 'AI Interview Used',
      'Total Candidates Progressed', 'AI Candidates Progressed', 'Non-AI Candidates Progressed',
      'Candidates Reached Offer', 'AI Candidates Reached Offer', 'Non-AI Candidates Reached Offer',
      'AI Cands-to-Offer Ratio', 'Non-AI Cands-to-Offer Ratio',
      'Hired Candidate Name', 'Included in Calc (>3 Progressed)'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    
    const progressedStages = [
      'Hiring Manager Screen', 'Assessment', 'Onsite Interview', 'Final Interview', 'Candidate Withdrew', 'Candidate Hold',
      'Offer Approvals', 'Offer Extended', 'Offer Declined', 'Pending Start', 'Hired'
    ];
    
    const hiringStages = [
      'Offer Approvals', 'Offer Extended', 'Offer Declined', 'Pending Start', 'Hired'
    ];
    
    const { 
      Position_id: positionIdIndex, Title: positionTitleIndex, 'Recruiter name': recruiterNameIndex,
      'Position approved date': positionApprovedIndex, Ai_interview: aiInterviewIndex,
      Last_stage: lastStageIndex, Name: nameIndex 
    } = appColIndices;

    if (positionIdIndex === undefined || positionApprovedIndex === undefined) {
      throw new Error('Required columns Position_id or Position approved date not found');
    }
    
    const currentYear = new Date().getFullYear();
    const thisYearPositions = appRows.filter(row => {
      const approvedDate = row[positionApprovedIndex];
      if (!approvedDate) return false;
      const date = vsParseDateSafeRB(approvedDate);
      return date && date.getFullYear() === currentYear;
    });
    
    const filteredPositions = thisYearPositions.filter(row => {
      if (recruiterNameIndex === undefined) return true;
      const recruiterName = (row[recruiterNameIndex] || '').toLowerCase();
      return !recruiterName.includes('samrudh') && !recruiterName.includes('simran');
    });
    
    const positionData = {};
    
    filteredPositions.forEach(row => {
      const positionId = row[positionIdIndex];
      if (!positionData[positionId]) {
        positionData[positionId] = {
          positionId, positionTitle: row[positionTitleIndex] || 'N/A',
          recruiterName: row[recruiterNameIndex] || 'N/A', positionApprovedDate: row[positionApprovedIndex],
          totalProgressed: 0, aiProgressed: 0, nonAiProgressed: 0,
          totalReachedOffer: 0, aiReachedOffer: 0, nonAiReachedOffer: 0,
          hasAiCandidates: false, hiredCandidates: []
        };
      }
      
      const stats = positionData[positionId];
      const lastStage = row[lastStageIndex];
      const aiInterview = row[aiInterviewIndex];
      
      if (progressedStages.includes(lastStage)) {
        stats.totalProgressed++;
        const wasOffered = hiringStages.includes(lastStage);

        if (wasOffered) {
          stats.totalReachedOffer++;
        }

        if (aiInterview === 'Y') {
          stats.hasAiCandidates = true;
          stats.aiProgressed++;
          if (wasOffered) stats.aiReachedOffer++;
        } else {
          stats.nonAiProgressed++;
          if (wasOffered) stats.nonAiReachedOffer++;
        }
        
        if (wasOffered && lastStage === 'Hired' && nameIndex !== undefined) {
          stats.hiredCandidates.push(row[nameIndex] || '');
        }
      }
    });
    
    const positionRows = Object.values(positionData)
      .filter(pos => pos.totalProgressed > 0)
      .map(pos => {
        const isAiPosition = pos.hasAiCandidates;
        // Asymmetrical inclusion logic based on user feedback
        const inclusionThresholdMet = isAiPosition 
          ? pos.aiReachedOffer > 0 
          : (pos.nonAiProgressed >= 3 && pos.nonAiReachedOffer > 0);
        return [
          pos.positionId, pos.positionTitle, pos.recruiterName, pos.positionApprovedDate,
          isAiPosition ? 'Yes' : 'No',
          pos.totalProgressed, pos.aiProgressed, pos.nonAiProgressed,
          pos.totalReachedOffer, pos.aiReachedOffer, pos.nonAiReachedOffer,
          pos.aiReachedOffer > 0 ? (pos.aiProgressed / pos.aiReachedOffer).toFixed(1) : 'N/A',
          pos.nonAiReachedOffer > 0 ? (pos.nonAiProgressed / pos.nonAiReachedOffer).toFixed(1) : 'N/A',
          pos.hiredCandidates.join(', '),
          inclusionThresholdMet ? 'Yes' : 'No'
        ];
      });
    
    positionRows.sort((a, b) => b[5] - a[5]);
    
    if (positionRows.length > 0) {
      sheet.getRange(2, 1, positionRows.length, headers.length).setValues(positionRows);
    }
    
    sheet.autoResizeColumns(1, headers.length);
    
    Logger.log(`Validation sheet created with ${positionRows.length} positions.`);
    
    return spreadsheet.getUrl();
    
  } catch (error) {
    Logger.log(`Error creating validation sheet: ${error.toString()} Stack: ${error.stack}`);
    throw error;
  }
}

/**
 * Calculates AI interview coverage metrics by recruiter.
 * Shows how many eligible candidates (not in "New" or "Added" stages) should have had AI interviews.
 * @param {Array<Array>} appRows Rows from the application sheet.
 * @param {object} appColIndices Column indices for the application sheet.
 * @returns {object} Object containing coverage metrics by recruiter.
 */
function calculateAICoverageMetricsRB(appRows, appColIndices) {
  Logger.log(`--- Starting calculateAICoverageMetricsRB ---`);
  Logger.log(`DEBUG: Available column indices: ${JSON.stringify(appColIndices)}`);
  
  // More robust column name matching
  const findColumnIndex = (possibleNames) => {
    for (const name of possibleNames) {
      if (appColIndices[name] !== undefined && appColIndices[name] !== -1) {
        Logger.log(`DEBUG: Found column "${name}" at index ${appColIndices[name]}`);
        return appColIndices[name];
      }
    }
    return -1;
  };
  
  const recruiterNameIdx = findColumnIndex(['Recruiter name', 'Recruiter_name', 'RecruiterName', 'recruiter name', 'recruiter_name']);
  const lastStageIdx = findColumnIndex(['Last_stage', 'Last stage', 'LastStage', 'last_stage', 'last stage']);
  const aiInterviewIdx = findColumnIndex(['Ai_interview', 'AI_interview', 'AI Interview', 'ai_interview', 'ai interview']);
  const applicationTsIdx = findColumnIndex(['Application_ts', 'Application ts', 'ApplicationTs', 'application_ts']);
  
  Logger.log(`DEBUG: Column indices found - Recruiter name: ${recruiterNameIdx}, Last_stage: ${lastStageIdx}, Ai_interview: ${aiInterviewIdx}, Application_ts: ${applicationTsIdx}`);
  
  if (recruiterNameIdx === -1 || lastStageIdx === -1 || aiInterviewIdx === -1) {
    Logger.log(`ERROR: Required columns not found for AI coverage calculation.`);
    Logger.log(`ERROR: Available column names: ${Object.keys(appColIndices).join(', ')}`);
    Logger.log(`ERROR: Recruiter: ${recruiterNameIdx}, Last_stage: ${lastStageIdx}, Ai_interview: ${aiInterviewIdx}`);
    return null;
  }
  
  const recruiterCoverage = {};
  let totalEligible = 0;
  let totalAIInterviews = 0;
  
  // Debug: Log first few rows to see the data structure
  Logger.log(`DEBUG: Processing ${appRows.length} rows`);
  if (appRows.length > 0) {
    Logger.log(`DEBUG: First row sample - Recruiter: "${appRows[0][recruiterNameIdx]}", Last Stage: "${appRows[0][lastStageIdx]}", AI Interview: "${appRows[0][aiInterviewIdx]}"`);
    
    // Log more sample rows to see the data variety
    for (let i = 0; i < Math.min(5, appRows.length); i++) {
      const row = appRows[i];
      if (row && row.length > Math.max(recruiterNameIdx, lastStageIdx, aiInterviewIdx)) {
        const recruiter = String(row[recruiterNameIdx] || '').trim();
        const stage = String(row[lastStageIdx] || '').trim().toUpperCase();
        const aiInterview = String(row[aiInterviewIdx] || '').trim().toUpperCase();
        Logger.log(`DEBUG: Row ${i} - Recruiter: "${recruiter}", Stage: "${stage}", AI Interview: "${aiInterview}"`);
      }
    }
  }
  
  // Set May 1st, 2025 as the cutoff date
  const mayFirst2025 = new Date('2025-05-01');
  mayFirst2025.setHours(0, 0, 0, 0);
  
  appRows.forEach((row, index) => {
    // Basic validation
    if (!row || row.length <= Math.max(recruiterNameIdx, lastStageIdx, aiInterviewIdx)) {
      if (index < 5) Logger.log(`DEBUG: Skipping row ${index} due to incomplete data`);
      return; // Skip incomplete rows
    }
    
    const recruiterName = String(row[recruiterNameIdx] || '').trim();
    const lastStage = String(row[lastStageIdx] || '').trim().toUpperCase();
    const aiInterview = String(row[aiInterviewIdx] || '').trim().toUpperCase();
    
    // Skip if no recruiter name
    if (!recruiterName) {
      if (index < 5) Logger.log(`DEBUG: Skipping row ${index} due to missing recruiter name`);
      return;
    }
    
    // Skip excluded recruiters
    const excludedRecruiters = ['Samrudh J', 'Pavan Kumar', 'Guruprasad Hegde'];
    if (excludedRecruiters.some(excluded => recruiterName.toLowerCase().includes(excluded.toLowerCase()))) {
      if (index < 5) Logger.log(`DEBUG: Skipping row ${index} due to excluded recruiter: ${recruiterName}`);
      return;
    }
    
    // Check Application_ts filter (May 1st, 2025 or later)
    const applicationTs = applicationTsIdx !== -1 ? vsParseDateSafeRB(row[applicationTsIdx]) : null;
    if (!applicationTs || applicationTs < mayFirst2025) {
      if (index < 5) Logger.log(`DEBUG: Skipping row ${index} due to Application_ts before May 1st, 2025: ${applicationTs}`);
      return; // Skip candidates with application timestamp before May 1st, 2025
    }
    
    // Check if candidate is eligible (only specific stages) - CASE INSENSITIVE
    const eligibleStages = [
      'HIRING MANAGER SCREEN',
      'ASSESSMENT', 
      'ONSITE INTERVIEW',
      'FINAL INTERVIEW',
      'OFFER APPROVALS',
      'OFFER EXTENDED',
      'OFFER DECLINED',
      'PENDING START',
      'HIRED'
    ];
    const isEligible = eligibleStages.some(stage => stage.toUpperCase() === lastStage);
    
    if (index < 5) {
      Logger.log(`DEBUG: Row ${index} - Recruiter: "${recruiterName}", Last Stage: "${lastStage}", AI Interview: "${aiInterview}", Eligible: ${isEligible}`);
    }
    
    if (isEligible) {
      totalEligible++;
      
      // Initialize recruiter data if not exists
      if (!recruiterCoverage[recruiterName]) {
        recruiterCoverage[recruiterName] = {
          totalEligible: 0,
          totalAIInterviews: 0,
          percentage: 0
        };
      }
      
      recruiterCoverage[recruiterName].totalEligible++;
      
      // Check if AI interview was conducted
      if (aiInterview === 'Y') {
        totalAIInterviews++;
        recruiterCoverage[recruiterName].totalAIInterviews++;
      }
    }
  });
  
  // Calculate percentages for each recruiter
  Object.keys(recruiterCoverage).forEach(recruiter => {
    const data = recruiterCoverage[recruiter];
    data.percentage = data.totalEligible > 0 ? 
      parseFloat(((data.totalAIInterviews / data.totalEligible) * 100).toFixed(1)) : 0;
  });
  
  // Calculate overall percentage
  const overallPercentage = totalEligible > 0 ? 
    parseFloat(((totalAIInterviews / totalEligible) * 100).toFixed(1)) : 0;
  
  Logger.log(`AI Coverage Metrics: Total eligible candidates = ${totalEligible}, Total AI interviews = ${totalAIInterviews}, Overall percentage = ${overallPercentage}%`);
  Logger.log(`DEBUG: Recruiter coverage data: ${JSON.stringify(recruiterCoverage)}`);
  
  // If no eligible candidates found, return a special object to show the table with a message
  if (totalEligible === 0) {
    Logger.log(`WARNING: No eligible candidates found. This could mean no candidates are in the target stages.`);
    return {
      recruiterCoverage: {},
      totalEligible: 0,
      totalAIInterviews: 0,
      overallPercentage: 0,
      noDataMessage: "No eligible candidates found. No candidates appear to be in the target stages (Hiring Manager Screen, Assessment, Onsite Interview, Final Interview, Offer Approvals, Offer Extended, Offer Declined, Pending Start, Hired), or there may be a data issue."
    };
  }
  
  return {
    recruiterCoverage,
    totalEligible,
    totalAIInterviews,
    overallPercentage
  };
}

/**
 * Test function to debug AI coverage calculation
 */
function testAICoverageCalculation() {
  try {
    Logger.log(`--- Testing AI Coverage Calculation ---`);
    Logger.log(`Filter: Application_ts ≥ May 1st, 2025`);
    
    // Get application data
    const appData = getApplicationDataForChartRB();
    if (!appData || !appData.rows) {
      Logger.log(`ERROR: Could not get application data`);
      return;
    }
    
    Logger.log(`Got ${appData.rows.length} rows from application sheet`);
    Logger.log(`Column indices: ${JSON.stringify(appData.colIndices)}`);
    
    // Test AI coverage calculation
    const aiCoverageMetrics = calculateAICoverageMetricsRB(appData.rows, appData.colIndices);
    
    if (aiCoverageMetrics) {
      Logger.log(`SUCCESS: AI Coverage metrics calculated`);
      Logger.log(`Total eligible: ${aiCoverageMetrics.totalEligible}`);
      Logger.log(`Total AI interviews: ${aiCoverageMetrics.totalAIInterviews}`);
      Logger.log(`Overall percentage: ${aiCoverageMetrics.overallPercentage}%`);
      Logger.log(`Recruiter coverage: ${JSON.stringify(aiCoverageMetrics.recruiterCoverage)}`);
    } else {
      Logger.log(`ERROR: AI Coverage metrics returned null`);
    }
    
  } catch (error) {
    Logger.log(`ERROR in test: ${error.toString()}`);
  }
}

/**
 * Generates HTML for a bar chart showing AI interview coverage by recruiter.
 * Each bar represents a recruiter with Y (AI interview done) and N (AI interview missing) stacked.
 * @param {object} aiCoverageMetrics The AI coverage metrics object from calculateAICoverageMetricsRB.
 * @returns {string} HTML string for the bar chart.
 */
function generateAICoverageBarChartHtml(aiCoverageMetrics) {
  if (!aiCoverageMetrics || !aiCoverageMetrics.recruiterCoverage || Object.keys(aiCoverageMetrics.recruiterCoverage).length === 0) {
    return `
      <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
        <div style="font-weight: 700; font-size: 15px; color: #1a1a1a; margin-bottom: 16px; letter-spacing: -0.3px;">📊 AI Interview Coverage by Recruiter</div>
        <div style="text-align: center; padding: 40px 20px; color: #999; font-size: 13px;">No AI coverage data available.</div>
      </div>
    `;
  }

  // Sort recruiters by total eligible candidates (descending)
  const sortedRecruiters = Object.entries(aiCoverageMetrics.recruiterCoverage)
    .sort(([, a], [, b]) => b.totalEligible - a.totalEligible);

  // Calculate max value for scaling
  const maxEligible = Math.max(...sortedRecruiters.map(([, data]) => data.totalEligible));

  let chartHtml = `
    <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
      <div style="font-weight: 700; font-size: 15px; color: #1a1a1a; margin-bottom: 16px; letter-spacing: -0.3px;">📊 AI Interview Coverage by Recruiter</div>
      <p style="font-size: 11px; color: #999; margin-bottom: 16px; line-height: 1.5;">
        Eligible candidates (Application_ts ≥ May 1, 2025). 
        <span style="color: #4CAF50; font-weight: 600;">Green</span> = Done (Y), 
        <span style="color: #F44336; font-weight: 600;">Red</span> = Missing (N).
      </p>
      
      <div style="margin: 16px 0;">
  `;

  // Generate bars
  sortedRecruiters.forEach(([recruiter, data]) => {
    const totalEligible = data.totalEligible;
    const aiInterviewsDone = data.totalAIInterviews;
    const aiInterviewsMissing = totalEligible - aiInterviewsDone;
    
    // Calculate bar widths (max width 300px)
    const maxBarWidth = 300;
    const barWidth = (totalEligible / maxEligible) * maxBarWidth;
    const doneWidth = (aiInterviewsDone / totalEligible) * barWidth;
    const missingWidth = (aiInterviewsMissing / totalEligible) * barWidth;
    
    // Truncate long recruiter names
    const displayName = recruiter.length > 20 ? recruiter.substring(0, 17) + '...' : recruiter;
    
    chartHtml += `
      <div style="margin-bottom: 12px;">
        <div style="display: flex; align-items: center;">
          <div style="width: 140px; font-size: 12px; font-weight: 600; color: #1a1a1a; text-align: right; padding-right: 12px; overflow: hidden; text-overflow: ellipsis;" title="${recruiter}">
            ${displayName}
          </div>
          <div style="flex: 1; display: flex; align-items: center;">
            <div style="width: ${barWidth}px; height: 28px; display: flex; border-radius: 6px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
              ${aiInterviewsDone > 0 ? `<div style="width: ${doneWidth}px; background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%); display: flex; align-items: center; justify-content: center; color: white; font-size: 11px; font-weight: 700;" title="Done: ${aiInterviewsDone}">${aiInterviewsDone}</div>` : ''}
              ${aiInterviewsMissing > 0 ? `<div style="width: ${missingWidth}px; background: linear-gradient(135deg, #F44336 0%, #d32f2f 100%); display: flex; align-items: center; justify-content: center; color: white; font-size: 11px; font-weight: 700;" title="Missing: ${aiInterviewsMissing}">${aiInterviewsMissing}</div>` : ''}
            </div>
            <div style="margin-left: 12px; font-size: 11px; color: #667eea; font-weight: 600; min-width: 90px;">
              ${data.percentage}% <span style="color: #999; font-weight: 400;">(${aiInterviewsDone}/${totalEligible})</span>
            </div>
          </div>
        </div>
      </div>
    `;
  });

  // Add legend
  chartHtml += `
      </div>
      
      <div style="display: flex; justify-content: center; align-items: center; margin-top: 20px; padding-top: 16px; border-top: 1px solid #f0f0f0; font-size: 11px;">
        <div style="display: flex; align-items: center; margin-right: 24px;">
          <div style="width: 16px; height: 16px; background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%); border-radius: 3px; margin-right: 8px; box-shadow: 0 1px 2px rgba(0,0,0,0.1);"></div>
          <span style="color: #666; font-weight: 500;">Done (Y)</span>
        </div>
        <div style="display: flex; align-items: center;">
          <div style="width: 16px; height: 16px; background: linear-gradient(135deg, #F44336 0%, #d32f2f 100%); border-radius: 3px; margin-right: 8px; box-shadow: 0 1px 2px rgba(0,0,0,0.1);"></div>
          <span style="color: #666; font-weight: 500;">Missing (N)</span>
        </div>
      </div>
      
      <p style="font-size: 11px; color: #999; margin-top: 16px; text-align: center; padding-top: 12px; border-top: 1px solid #f0f0f0;">
        Eligible: ${aiCoverageMetrics.totalEligible} • 
        AI Done: ${aiCoverageMetrics.totalAIInterviews} • 
        Coverage: <strong style="color: #667eea;">${aiCoverageMetrics.overallPercentage}%</strong>
      </p>
    </div>
  `;

  return chartHtml;
}

/**
 * Test function to debug AI coverage bar chart
 */
function testAICoverageBarChart() {
  try {
    Logger.log(`--- Testing AI Coverage Bar Chart ---`);
    
    // Get application data
    const appData = getApplicationDataForChartRB();
    if (!appData || !appData.rows) {
      Logger.log(`ERROR: Could not get application data`);
      return;
    }
    
    Logger.log(`Got ${appData.rows.length} rows from application sheet`);
    
    // Test AI coverage calculation
    const aiCoverageMetrics = calculateAICoverageMetricsRB(appData.rows, appData.colIndices);
    
    if (aiCoverageMetrics) {
      Logger.log(`SUCCESS: AI Coverage metrics calculated`);
      Logger.log(`Total eligible: ${aiCoverageMetrics.totalEligible}`);
      Logger.log(`Total AI interviews: ${aiCoverageMetrics.totalAIInterviews}`);
      Logger.log(`Overall percentage: ${aiCoverageMetrics.overallPercentage}%`);
      
      // Test bar chart generation
      const barChartHtml = generateAICoverageBarChartHtml(aiCoverageMetrics);
      Logger.log(`SUCCESS: Bar chart HTML generated (${barChartHtml.length} characters)`);
      
      // Log sample of the HTML
      Logger.log(`Sample HTML (first 500 chars): ${barChartHtml.substring(0, 500)}...`);
      
      // Log recruiter data for verification
      Object.entries(aiCoverageMetrics.recruiterCoverage).forEach(([recruiter, data]) => {
        const missing = data.totalEligible - data.totalAIInterviews;
        Logger.log(`Recruiter: ${recruiter} | Eligible: ${data.totalEligible} | AI Done: ${data.totalAIInterviews} | Missing: ${missing} | %: ${data.percentage}%`);
      });
      
    } else {
      Logger.log(`ERROR: AI Coverage metrics returned null`);
    }
    
  } catch (error) {
    Logger.log(`ERROR in test: ${error.toString()}`);
  }
}

/**
 * Creates a detailed validation spreadsheet for a specific recruiter showing all their candidates.
 * @param {string} recruiterName The name of the recruiter to analyze.
 * @param {Array<Array>} appRows Raw rows from the application sheet.
 * @param {object} appColIndices Column indices map for the application sheet.
 * @returns {string} The URL of the created Google Sheet.
 */
function createRecruiterValidationSheet(recruiterName, appRows, appColIndices) {
  Logger.log(`--- Creating validation sheet for recruiter: ${recruiterName} ---`);
  
  try {
    const spreadsheet = SpreadsheetApp.create(`AI Interview Validation - ${recruiterName} - ${new Date().toISOString().split('T')[0]}`);
    const sheet = spreadsheet.getActiveSheet();
    
    // Set May 1st, 2025 as the cutoff date
    const mayFirst2025 = new Date('2025-05-01');
    mayFirst2025.setHours(0, 0, 0, 0);
    
    // Find column indices
    const findColumnIndex = (possibleNames) => {
      for (const name of possibleNames) {
        if (appColIndices[name] !== undefined && appColIndices[name] !== -1) {
          return appColIndices[name];
        }
      }
      return -1;
    };
    
    const recruiterNameIdx = findColumnIndex(['Recruiter name', 'Recruiter_name', 'RecruiterName', 'recruiter name', 'recruiter_name']);
    const lastStageIdx = findColumnIndex(['Last_stage', 'Last stage', 'LastStage', 'last_stage', 'last stage']);
    const aiInterviewIdx = findColumnIndex(['Ai_interview', 'AI_interview', 'AI Interview', 'ai_interview', 'ai interview']);
    const applicationTsIdx = findColumnIndex(['Application_ts', 'Application ts', 'ApplicationTs', 'application_ts']);
    const nameIdx = findColumnIndex(['Name', 'name', 'Candidate_name', 'Candidate name']);
    const positionIdIdx = findColumnIndex(['Position_id', 'Position id', 'Position ID']);
    const titleIdx = findColumnIndex(['Title', 'title', 'Position_title', 'Position title']);
    const currentCompanyIdx = findColumnIndex(['Current_company', 'Current company', 'Company', 'company']);
    const applicationStatusIdx = findColumnIndex(['Application_status', 'Application status', 'Status', 'status']);
    const positionStatusIdx = findColumnIndex(['Position_status', 'Position status']);
    const matchStarsIdx = findColumnIndex(['Match_stars', 'Match score', 'Match Stars', 'MatchStars', 'Match_Stars', 'Stars', 'Score']);
    
    if (recruiterNameIdx === -1 || lastStageIdx === -1 || aiInterviewIdx === -1) {
      throw new Error('Required columns not found for validation sheet');
    }
    
    // Define headers for the validation sheet
    const headers = [
      'Candidate Name',
      'Position ID', 
      'Position Title',
      'Current Company',
      'Application Status',
      'Position Status',
      'Last Stage',
      'Application Timestamp',
      'AI Interview (Y/N)',
      'Match Stars/Score',
      'Eligible for AI Interview',
      'AI Interview Status',
      'Days Since Application'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    
    // Filter data for the specific recruiter and apply date filter
    const recruiterData = appRows.filter(row => {
      if (!row || row.length <= Math.max(recruiterNameIdx, lastStageIdx, aiInterviewIdx)) {
        return false;
      }
      
      const rowRecruiterName = String(row[recruiterNameIdx] || '').trim();
      if (rowRecruiterName !== recruiterName) {
        return false;
      }
      
      // Check Application_ts filter (May 1st, 2025 or later)
      const applicationTs = applicationTsIdx !== -1 ? vsParseDateSafeRB(row[applicationTsIdx]) : null;
      if (!applicationTs || applicationTs < mayFirst2025) {
        return false;
      }
      
      return true;
    });
    
    Logger.log(`Found ${recruiterData.length} candidates for ${recruiterName} with Application_ts ≥ May 1st, 2025`);
    
    // Process each candidate
    const validationRows = recruiterData.map(row => {
      const lastStage = String(row[lastStageIdx] || '').trim().toUpperCase();
      const aiInterview = String(row[aiInterviewIdx] || '').trim().toUpperCase();
      const applicationTs = applicationTsIdx !== -1 ? vsParseDateSafeRB(row[applicationTsIdx]) : null;
      
      // Check if candidate is eligible (only specific stages) - CASE INSENSITIVE
      const eligibleStages = [
        'HIRING MANAGER SCREEN',
        'ASSESSMENT', 
        'ONSITE INTERVIEW',
        'FINAL INTERVIEW',
        'OFFER APPROVALS',
        'OFFER EXTENDED',
        'OFFER DECLINED',
        'PENDING START',
        'HIRED'
      ];
      const isEligible = eligibleStages.some(stage => stage.toUpperCase() === lastStage);
      
      // Determine AI interview status
      let aiInterviewStatus = 'N/A';
      if (isEligible) {
        if (aiInterview === 'Y') {
          aiInterviewStatus = '✅ AI Interview Done';
        } else if (aiInterview === 'N') {
          aiInterviewStatus = '❌ AI Interview Missing';
        } else {
          aiInterviewStatus = '❓ Unknown Status';
        }
      } else {
        aiInterviewStatus = '⏭️ Not Eligible (Not in target stages)';
      }
      
      // Calculate days since application
      const daysSinceApplication = applicationTs ? 
        Math.floor((new Date() - applicationTs) / (1000 * 60 * 60 * 24)) : 'N/A';
      
      return [
        nameIdx !== -1 ? row[nameIdx] || 'N/A' : 'N/A',
        positionIdIdx !== -1 ? row[positionIdIdx] || 'N/A' : 'N/A',
        titleIdx !== -1 ? row[titleIdx] || 'N/A' : 'N/A',
        currentCompanyIdx !== -1 ? row[currentCompanyIdx] || 'N/A' : 'N/A',
        applicationStatusIdx !== -1 ? row[applicationStatusIdx] || 'N/A' : 'N/A',
        positionStatusIdx !== -1 ? row[positionStatusIdx] || 'N/A' : 'N/A',
        lastStage,
        applicationTs ? applicationTs.toLocaleDateString() : 'N/A',
        aiInterview,
        matchStarsIdx !== -1 ? row[matchStarsIdx] || 'N/A' : 'N/A',
        isEligible ? 'Yes' : 'No',
        aiInterviewStatus,
        daysSinceApplication
      ];
    });
    
    // Sort by AI interview status (missing first, then done, then not eligible)
    validationRows.sort((a, b) => {
      const statusA = a[11]; // AI Interview Status column
      const statusB = b[11];
      
      if (statusA.includes('Missing') && !statusB.includes('Missing')) return -1;
      if (!statusA.includes('Missing') && statusB.includes('Missing')) return 1;
      if (statusA.includes('Done') && !statusB.includes('Done')) return -1;
      if (!statusA.includes('Done') && statusB.includes('Done')) return 1;
      return 0;
    });
    
    if (validationRows.length > 0) {
      sheet.getRange(2, 1, validationRows.length, headers.length).setValues(validationRows);
    }
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
    
    // Add summary statistics
    const totalCandidates = validationRows.length;
    const eligibleCandidates = validationRows.filter(row => row[10] === 'Yes').length;
    const aiInterviewsDone = validationRows.filter(row => row[11].includes('Done')).length;
    const aiInterviewsMissing = validationRows.filter(row => row[11].includes('Missing')).length;
    const notEligible = validationRows.filter(row => row[11].includes('Not Eligible')).length;
    
    // Add summary at the top
    const summaryRow = [
      `SUMMARY FOR ${recruiterName.toUpperCase()}`,
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      ''
    ];
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([summaryRow]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e3f2fd');
    
    // Add statistics rows
    const statsRows = [
      [`Total Candidates (Application_ts ≥ May 1st, 2025): ${totalCandidates}`],
      [`Eligible Candidates (in target stages): ${eligibleCandidates}`],
      [`AI Interviews Done: ${aiInterviewsDone}`],
      [`AI Interviews Missing: ${aiInterviewsMissing}`],
      [`Not Eligible (not in target stages): ${notEligible}`],
      [`Coverage Rate: ${eligibleCandidates > 0 ? ((aiInterviewsDone / eligibleCandidates) * 100).toFixed(1) : 0}%`]
    ];
    
    sheet.insertRowsBefore(2, statsRows.length);
    for (let i = 0; i < statsRows.length; i++) {
      sheet.getRange(2 + i, 1, 1, 1).setValues([statsRows[i]]);
      sheet.getRange(2 + i, 1, 1, 1).setFontWeight('bold');
    }
    
    // Add conditional formatting for AI interview status
    const statusRange = sheet.getRange(3 + statsRows.length, 12, validationRows.length, 1); // AI Interview Status column
    const rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('✅ AI Interview Done')
      .setBackground('#d4edda')
      .setRanges([statusRange])
      .build();
    
    const rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('❌ AI Interview Missing')
      .setBackground('#f8d7da')
      .setRanges([statusRange])
      .build();
    
    const rule3 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('⏭️ Not Eligible (New/Added)')
      .setBackground('#fff3cd')
      .setRanges([statusRange])
      .build();
    
    sheet.setConditionalFormatRules([rule1, rule2, rule3]);
    
    Logger.log(`Validation sheet created for ${recruiterName} with ${validationRows.length} candidates`);
    Logger.log(`Summary: ${eligibleCandidates} eligible, ${aiInterviewsDone} AI done, ${aiInterviewsMissing} AI missing`);
    
    return spreadsheet.getUrl();
    
  } catch (error) {
    Logger.log(`Error creating validation sheet for ${recruiterName}: ${error.toString()} Stack: ${error.stack}`);
    throw error;
  }
}

/**
 * Test function to create validation sheet for a specific recruiter
 */
function testCreateRecruiterValidationSheet() {
  try {
    Logger.log(`--- Testing Recruiter Validation Sheet Creation ---`);
    
    // Get application data
    const appData = getApplicationDataForChartRB();
    if (!appData || !appData.rows) {
      Logger.log(`ERROR: Could not get application data`);
      return;
    }
    
    Logger.log(`Got ${appData.rows.length} rows from application sheet`);
    
    // Test with a specific recruiter (you can change this name)
    const testRecruiter = 'Akhila Kashyap';
    Logger.log(`Creating validation sheet for: ${testRecruiter}`);
    
    const sheetUrl = createRecruiterValidationSheet(testRecruiter, appData.rows, appData.colIndices);
    Logger.log(`SUCCESS: Validation sheet created: ${sheetUrl}`);
    
  } catch (error) {
    Logger.log(`ERROR in test: ${error.toString()}`);
  }
}

/**
 * Creates validation sheets for all recruiters and returns a summary of URLs.
 * @param {Array<Array>} appRows Raw rows from the application sheet.
 * @param {object} appColIndices Column indices map for the application sheet.
 * @returns {object} Object containing validation sheet URLs and summary.
 */
function createAllRecruiterValidationSheets(appRows, appColIndices) {
  Logger.log(`--- Creating validation sheets for all recruiters ---`);
  
  try {
    // First, get the list of all recruiters from the AI coverage data
    const aiCoverageMetrics = calculateAICoverageMetricsRB(appRows, appColIndices);
    if (!aiCoverageMetrics || !aiCoverageMetrics.recruiterCoverage) {
      Logger.log(`ERROR: Could not calculate AI coverage metrics`);
      return null;
    }
    
    const recruiterNames = Object.keys(aiCoverageMetrics.recruiterCoverage);
    Logger.log(`Found ${recruiterNames.length} recruiters to create validation sheets for`);
    
    const validationSheets = {};
    const failedRecruiters = [];
    
    // Create validation sheet for each recruiter
    recruiterNames.forEach(recruiterName => {
      try {
        Logger.log(`Creating validation sheet for: ${recruiterName}`);
        const sheetUrl = createRecruiterValidationSheet(recruiterName, appRows, appColIndices);
        validationSheets[recruiterName] = {
          url: sheetUrl,
          eligible: aiCoverageMetrics.recruiterCoverage[recruiterName].totalEligible,
          aiDone: aiCoverageMetrics.recruiterCoverage[recruiterName].totalAIInterviews,
          aiMissing: aiCoverageMetrics.recruiterCoverage[recruiterName].totalEligible - aiCoverageMetrics.recruiterCoverage[recruiterName].totalAIInterviews,
          percentage: aiCoverageMetrics.recruiterCoverage[recruiterName].percentage
        };
        Logger.log(`SUCCESS: Validation sheet created for ${recruiterName}`);
      } catch (error) {
        Logger.log(`ERROR creating validation sheet for ${recruiterName}: ${error.toString()}`);
        failedRecruiters.push(recruiterName);
      }
    });
    
    Logger.log(`Created ${Object.keys(validationSheets).length} validation sheets successfully`);
    if (failedRecruiters.length > 0) {
      Logger.log(`Failed to create sheets for: ${failedRecruiters.join(', ')}`);
    }
    
    return {
      validationSheets,
      failedRecruiters,
      totalRecruiters: recruiterNames.length,
      successfulSheets: Object.keys(validationSheets).length
    };
    
  } catch (error) {
    Logger.log(`ERROR in createAllRecruiterValidationSheets: ${error.toString()} Stack: ${error.stack}`);
    return null;
  }
}

/**
 * Generates HTML for validation sheets section in the report.
 * @param {object} validationData The validation data from createAllRecruiterValidationSheets.
 * @returns {string} HTML string for the validation sheets section.
 */
function generateValidationSheetsHtml(validationData) {
  if (!validationData || !validationData.validationSheets || Object.keys(validationData.validationSheets).length === 0) {
    return `
      <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
        <div style="font-weight: 700; font-size: 15px; color: #1a1a1a; margin-bottom: 16px; letter-spacing: -0.3px;">📋 Detailed Validation Sheets</div>
        <div style="text-align: center; padding: 40px 20px; color: #999; font-size: 13px;">No validation sheets available.</div>
      </div>
    `;
  }
  
  // Sort recruiters by AI interviews missing (descending) to highlight those needing attention
  const sortedRecruiters = Object.entries(validationData.validationSheets)
    .sort(([, a], [, b]) => b.aiMissing - a.aiMissing);
  
  let html = `
    <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
      <div style="font-weight: 700; font-size: 15px; color: #1a1a1a; margin-bottom: 16px; letter-spacing: -0.3px;">📋 Detailed Validation Sheets</div>
      <p style="font-size: 11px; color: #999; margin-bottom: 16px; line-height: 1.5;">
        Click recruiter name to view detailed candidate list. Application_ts ≥ May 1, 2025.
      </p>
      
      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
        <thead>
          <tr>
            <th style="padding: 10px 8px; text-align: left; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">Recruiter</th>
            <th style="padding: 10px 8px; text-align: right; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">Eligible</th>
            <th style="padding: 10px 8px; text-align: right; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">AI Done</th>
            <th style="padding: 10px 8px; text-align: right; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">AI Missing</th>
            <th style="padding: 10px 8px; text-align: right; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">Coverage</th>
            <th style="padding: 10px 8px; text-align: center; font-size: 10px; font-weight: 600; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #f0f0f0;">Action</th>
          </tr>
        </thead>
        <tbody>
  `;
  
  sortedRecruiters.forEach(([recruiterName, data], index) => {
    const bgColor = index % 2 === 0 ? '#fafafa' : '#ffffff';
    const coverageColor = data.percentage >= 80 ? '#4CAF50' : data.percentage >= 60 ? '#FF9800' : '#F44336';
    const missingPercentage = 100 - data.percentage;
    
    html += `
      <tr>
        <td style="padding: 12px 8px; font-size: 12px; color: #1a1a1a; font-weight: 600; border-bottom: 1px solid #f5f5f5;">
          ${recruiterName}
        </td>
        <td style="padding: 12px 8px; text-align: right; font-size: 12px; color: #1a1a1a; font-weight: 500; border-bottom: 1px solid #f5f5f5;">${data.eligible}</td>
        <td style="padding: 12px 8px; text-align: right; font-size: 12px; color: #4CAF50; font-weight: 700; border-bottom: 1px solid #f5f5f5;">${data.aiDone}</td>
        <td style="padding: 12px 8px; text-align: right; font-size: 12px; color: #F44336; font-weight: 700; border-bottom: 1px solid #f5f5f5;">${data.aiMissing}</td>
        <td style="padding: 12px 8px; text-align: right; font-size: 12px; color: ${coverageColor}; font-weight: 700; border-bottom: 1px solid #f5f5f5;">${data.percentage}%</td>
        <td style="padding: 12px 8px; text-align: center; font-size: 12px; border-bottom: 1px solid #f5f5f5;">
          <a href="${data.url}" target="_blank" style="color: #667eea; text-decoration: none; font-weight: 600; padding: 4px 12px; background: rgba(102, 126, 234, 0.1); border-radius: 6px; display: inline-block;">View</a>
        </td>
      </tr>
    `;
  });
  
  html += `
        </tbody>
      </table>
      
      <p style="font-size: 10px; color: #999; margin-top: 16px; padding-top: 12px; border-top: 1px solid #f0f0f0; text-align: center; line-height: 1.6;">
        High (>30% missing) | Medium (10-30%) | Low (<10%)<br>
        Created ${validationData.successfulSheets} sheets successfully.
        ${validationData.failedRecruiters.length > 0 ? `<span style="color: #f5576c;">Failed: ${validationData.failedRecruiters.join(', ')}</span>` : ''}
      </p>
    </div>
  `;
  
  return html;
}

/**
 * Comprehensive test function for validation sheet functionality
 */
/**
 * Test function to debug feedback and pending counts for a specific recruiter
 * Tests: Thulan Kumar
 * Expected: 33 rows with Feedback_status = SUBMITTED
 * Checks: Pending count based on status vs Feedback_status = REQUESTED
 */
function testRecruiterFeedbackCount() {
  try {
    Logger.log(`=== Testing Feedback Count for Thulan Kumar ===`);
    
    // Get raw log data
    const logData = getLogSheetDataRB();
    if (!logData || !logData.rows || logData.rows.length === 0) {
      Logger.log(`ERROR: Could not get log data`);
      return;
    }
    
    Logger.log(`Total rows in log sheet: ${logData.rows.length}`);
    
    // Get column indices
    const recruiterIdx = logData.colIndices.hasOwnProperty('Recruiter_name') ? logData.colIndices['Recruiter_name'] : -1;
    const feedbackStatusIdx = logData.colIndices.hasOwnProperty('Feedback_status') ? logData.colIndices['Feedback_status'] : -1;
    const profileIdIdx = logData.colIndices['Profile_id'];
    const positionIdIdx = logData.colIndices['Position_id'];
    const statusIdx = logData.colIndices['STATUS_COLUMN'];
    
    Logger.log(`Column indices - Recruiter_name: ${recruiterIdx}, Feedback_status: ${feedbackStatusIdx}, Profile_id: ${profileIdIdx}, Position_id: ${positionIdIdx}`);
    
    if (recruiterIdx === -1) {
      Logger.log(`ERROR: Recruiter_name column not found`);
      return;
    }
    
    if (feedbackStatusIdx === -1) {
      Logger.log(`ERROR: Feedback_status column not found`);
      return;
    }
    
    // Step 1: Count raw rows for Thulan Kumar with SUBMITTED
    let rawCount = 0;
    let rawRows = [];
    logData.rows.forEach((row, index) => {
      if (row.length > recruiterIdx && row.length > feedbackStatusIdx) {
        const recruiter = row[recruiterIdx] ? String(row[recruiterIdx]).trim() : '';
        const feedbackStatus = row[feedbackStatusIdx] ? String(row[feedbackStatusIdx]).trim() : '';
        
        if (recruiter === 'Thulan Kumar' && feedbackStatus.toUpperCase() === 'SUBMITTED') {
          rawCount++;
          rawRows.push({ rowIndex: index, profileId: row[profileIdIdx], positionId: row[positionIdIdx], status: row[statusIdx] });
        }
      }
    });
    
    Logger.log(`\nStep 1 - Raw count (Thulan Kumar + SUBMITTED): ${rawCount}`);
    Logger.log(`Sample rows: ${JSON.stringify(rawRows.slice(0, 5))}`);
    
    // Step 2: After time range filter
    const filteredData = filterDataByTimeRangeRB(logData.rows, logData.colIndices);
    let timeFilteredCount = 0;
    filteredData.forEach(row => {
      if (row.length > recruiterIdx && row.length > feedbackStatusIdx) {
        const recruiter = row[recruiterIdx] ? String(row[recruiterIdx]).trim() : '';
        const feedbackStatus = row[feedbackStatusIdx] ? String(row[feedbackStatusIdx]).trim() : '';
        if (recruiter === 'Thulan Kumar' && feedbackStatus.toUpperCase() === 'SUBMITTED') {
          timeFilteredCount++;
        }
      }
    });
    Logger.log(`\nStep 2 - After time range filter: ${timeFilteredCount}`);
    
    // Step 3: After template filter
    const feedbackTemplateIndex = logData.colIndices.hasOwnProperty('Feedback_template_name') ? logData.colIndices['Feedback_template_name'] : -1;
    const excludedTemplates = ["AI Coding Interview Metrics Feedback Form", "AI Functional Interview Feedback Form"];
    let templateFilteredData = filteredData;
    if (feedbackTemplateIndex !== -1) {
      templateFilteredData = filteredData.filter(row => {
        if (row.length <= feedbackTemplateIndex) return true;
        const templateName = row[feedbackTemplateIndex] ? String(row[feedbackTemplateIndex]).trim() : '';
        return !excludedTemplates.includes(templateName);
      });
    }
    
    let templateFilteredCount = 0;
    templateFilteredData.forEach(row => {
      if (row.length > recruiterIdx && row.length > feedbackStatusIdx) {
        const recruiter = row[recruiterIdx] ? String(row[recruiterIdx]).trim() : '';
        const feedbackStatus = row[feedbackStatusIdx] ? String(row[feedbackStatusIdx]).trim() : '';
        if (recruiter === 'Thulan Kumar' && feedbackStatus.toUpperCase() === 'SUBMITTED') {
          templateFilteredCount++;
        }
      }
    });
    Logger.log(`\nStep 3 - After template filter: ${templateFilteredCount}`);
    
    // Step 4: After position filter
    const positionNameIndex = logData.colIndices.hasOwnProperty('Position_name') ? logData.colIndices['Position_name'] : -1;
    const positionToExclude = "AIR Testing";
    let finalFilteredData = templateFilteredData;
    if (positionNameIndex !== -1) {
      finalFilteredData = templateFilteredData.filter(row => {
        return !(row.length > positionNameIndex && row[positionNameIndex] === positionToExclude);
      });
    }
    
    let positionFilteredCount = 0;
    finalFilteredData.forEach(row => {
      if (row.length > recruiterIdx && row.length > feedbackStatusIdx) {
        const recruiter = row[recruiterIdx] ? String(row[recruiterIdx]).trim() : '';
        const feedbackStatus = row[feedbackStatusIdx] ? String(row[feedbackStatusIdx]).trim() : '';
        if (recruiter === 'Thulan Kumar' && feedbackStatus.toUpperCase() === 'SUBMITTED') {
          positionFilteredCount++;
        }
      }
    });
    Logger.log(`\nStep 4 - After position filter: ${positionFilteredCount}`);
    
    // Step 5: After deduplication
    const groupedData = {};
    let skippedRowCount = 0;
    
    finalFilteredData.forEach(row => {
      if (!row || row.length <= profileIdIdx || row.length <= positionIdIdx || row.length <= statusIdx) {
        skippedRowCount++;
        return;
      }
      const profileId = row[profileIdIdx];
      const positionId = row[positionIdIdx];
      const status = row[statusIdx] ? String(row[statusIdx]).trim() : 'Unknown';
      
      if (!profileId || !positionId) {
        skippedRowCount++;
        return;
      }
      
      const key = `${profileId}_${positionId}`;
      const statusRank = getStatusRank(status);
      
      // Check if current row has SUBMITTED feedback
      const currentFeedbackStatus = (row.length > feedbackStatusIdx && row[feedbackStatusIdx]) ? String(row[feedbackStatusIdx]).trim().toUpperCase() : '';
      const hasSubmittedFeedback = currentFeedbackStatus === 'SUBMITTED';
      
      if (!groupedData[key]) {
        // First row for this key - store it
        groupedData[key] = { bestRank: statusRank, row: row, hasSubmittedFeedback: hasSubmittedFeedback };
      } else {
        const existing = groupedData[key];
        const existingFeedbackStatus = (existing.row.length > feedbackStatusIdx && existing.row[feedbackStatusIdx]) ? String(existing.row[feedbackStatusIdx]).trim().toUpperCase() : '';
        const existingHasSubmitted = existingFeedbackStatus === 'SUBMITTED';
        
        // Priority: Keep row with SUBMITTED feedback if either has it, otherwise keep better status
        if (hasSubmittedFeedback && !existingHasSubmitted) {
          // Current row has SUBMITTED, existing doesn't - prefer current
          groupedData[key] = { bestRank: statusRank, row: row, hasSubmittedFeedback: true };
        } else if (!hasSubmittedFeedback && existingHasSubmitted) {
          // Existing has SUBMITTED, current doesn't - keep existing
          // Do nothing, keep existing row
        } else if (statusRank < existing.bestRank) {
          // Both have same feedback status, prefer better status rank
          groupedData[key] = { bestRank: statusRank, row: row, hasSubmittedFeedback: hasSubmittedFeedback };
        }
        // Otherwise keep existing row
      }
    });
    
    const deduplicatedRows = Object.values(groupedData).map(item => item.row);
    
    let deduplicatedCount = 0;
    let duplicateDetails = [];
    deduplicatedRows.forEach(row => {
      if (row.length > recruiterIdx && row.length > feedbackStatusIdx) {
        const recruiter = row[recruiterIdx] ? String(row[recruiterIdx]).trim() : '';
        const feedbackStatus = row[feedbackStatusIdx] ? String(row[feedbackStatusIdx]).trim() : '';
        if (recruiter === 'Thulan Kumar' && feedbackStatus.toUpperCase() === 'SUBMITTED') {
          deduplicatedCount++;
          const key = `${row[profileIdIdx]}_${row[positionIdIdx]}`;
          duplicateDetails.push({ profileId: row[profileIdIdx], positionId: row[positionIdIdx], status: row[statusIdx] });
        }
      }
    });
    
    Logger.log(`\nStep 5 - After deduplication: ${deduplicatedCount}`);
    Logger.log(`Skipped rows during deduplication: ${skippedRowCount}`);
    
    // Check for duplicates that might have been removed
    const duplicateKeys = new Set();
    const allKeys = [];
    finalFilteredData.forEach(row => {
      if (row.length > recruiterIdx && row.length > feedbackStatusIdx) {
        const recruiter = row[recruiterIdx] ? String(row[recruiterIdx]).trim() : '';
        const feedbackStatus = row[feedbackStatusIdx] ? String(row[feedbackStatusIdx]).trim() : '';
        if (recruiter === 'Thulan Kumar' && feedbackStatus.toUpperCase() === 'SUBMITTED') {
          if (row.length > profileIdIdx && row.length > positionIdIdx) {
            const key = `${row[profileIdIdx]}_${row[positionIdIdx]}`;
            allKeys.push(key);
            if (allKeys.filter(k => k === key).length > 1) {
              duplicateKeys.add(key);
            }
          }
        }
      }
    });
    
    Logger.log(`\nDuplicate Profile_id + Position_id combinations found: ${duplicateKeys.size}`);
    if (duplicateKeys.size > 0) {
      Logger.log(`Duplicate keys: ${Array.from(duplicateKeys).slice(0, 10).join(', ')}`);
      
      // Show details for each duplicate key
      Array.from(duplicateKeys).slice(0, 5).forEach(key => {
        const duplicateRows = finalFilteredData.filter(row => {
          if (row.length > recruiterIdx && row.length > feedbackStatusIdx && row.length > profileIdIdx && row.length > positionIdIdx) {
            const recruiter = row[recruiterIdx] ? String(row[recruiterIdx]).trim() : '';
            const feedbackStatus = row[feedbackStatusIdx] ? String(row[feedbackStatusIdx]).trim() : '';
            const rowKey = `${row[profileIdIdx]}_${row[positionIdIdx]}`;
            return recruiter === 'Thulan Kumar' && feedbackStatus.toUpperCase() === 'SUBMITTED' && rowKey === key;
          }
          return false;
        });
        
        Logger.log(`\nDuplicate key ${key} details:`);
        duplicateRows.forEach((row, idx) => {
          const status = row.length > statusIdx ? String(row[statusIdx]).trim() : 'N/A';
          const feedbackStatus = row.length > feedbackStatusIdx ? String(row[feedbackStatusIdx]).trim() : 'N/A';
          Logger.log(`  Row ${idx + 1}: Status=${status}, Feedback_status=${feedbackStatus}`);
        });
        
        // Check which one was kept after deduplication
        const keptRow = deduplicatedRows.find(row => {
          if (row.length > profileIdIdx && row.length > positionIdIdx) {
            const rowKey = `${row[profileIdIdx]}_${row[positionIdIdx]}`;
            return rowKey === key;
          }
          return false;
        });
        
        if (keptRow) {
          const keptStatus = keptRow.length > statusIdx ? String(keptRow[statusIdx]).trim() : 'N/A';
          const keptFeedback = keptRow.length > feedbackStatusIdx ? String(keptRow[feedbackStatusIdx]).trim() : 'N/A';
          Logger.log(`  KEPT: Status=${keptStatus}, Feedback_status=${keptFeedback}`);
        }
      });
    }
    
    // PENDING COUNT ANALYSIS
    Logger.log(`\n=== PENDING COUNT ANALYSIS FOR THULAN KUMAR ===`);
    
    const PENDING_STATUSES = ['PENDING', 'INVITED', 'EMAIL SENT'];
    
    // Count rows with Feedback_status = REQUESTED
    let requestedFeedbackCount = 0;
    let requestedRows = [];
    finalFilteredData.forEach(row => {
      if (row.length > recruiterIdx && row.length > feedbackStatusIdx) {
        const recruiter = row[recruiterIdx] ? String(row[recruiterIdx]).trim() : '';
        const feedbackStatus = row[feedbackStatusIdx] ? String(row[feedbackStatusIdx]).trim().toUpperCase() : '';
        if (recruiter === 'Thulan Kumar' && feedbackStatus === 'REQUESTED') {
          requestedFeedbackCount++;
          const status = row.length > statusIdx ? String(row[statusIdx]).trim() : '';
          requestedRows.push({ status: status, profileId: row[profileIdIdx], positionId: row[positionIdIdx] });
        }
      }
    });
    Logger.log(`Rows with Feedback_status = REQUESTED (before deduplication): ${requestedFeedbackCount}`);
    
    // Count rows with status = PENDING/INVITED/EMAIL SENT
    let pendingByStatusCount = 0;
    finalFilteredData.forEach(row => {
      if (row.length > recruiterIdx && row.length > statusIdx) {
        const recruiter = row[recruiterIdx] ? String(row[recruiterIdx]).trim() : '';
        const status = String(row[statusIdx]).trim();
        if (recruiter === 'Thulan Kumar' && PENDING_STATUSES.includes(status)) {
          pendingByStatusCount++;
        }
      }
    });
    Logger.log(`Rows with status in [PENDING, INVITED, EMAIL SENT] (before deduplication): ${pendingByStatusCount}`);
    
    // Count after deduplication
    let deduplicatedPendingCount = 0;
    deduplicatedRows.forEach(row => {
      if (row.length > recruiterIdx && row.length > statusIdx) {
        const recruiter = row[recruiterIdx] ? String(row[recruiterIdx]).trim() : '';
        const status = String(row[statusIdx]).trim();
        if (recruiter === 'Thulan Kumar' && PENDING_STATUSES.includes(status)) {
          deduplicatedPendingCount++;
        }
      }
    });
    Logger.log(`Rows with status in [PENDING, INVITED, EMAIL SENT] (after deduplication): ${deduplicatedPendingCount}`);
    
    // Analyze REQUESTED feedback rows
    let requestedWithPendingStatus = 0;
    let requestedWithOtherStatus = [];
    requestedRows.forEach(r => {
      if (PENDING_STATUSES.includes(r.status)) {
        requestedWithPendingStatus++;
      } else {
        requestedWithOtherStatus.push(r.status);
      }
    });
    Logger.log(`\nOf ${requestedFeedbackCount} REQUESTED feedback rows:`);
    Logger.log(`  - ${requestedWithPendingStatus} have status in [PENDING, INVITED, EMAIL SENT]`);
    Logger.log(`  - ${requestedWithOtherStatus.length} have other statuses`);
    if (requestedWithOtherStatus.length > 0) {
      const statusCounts = {};
      requestedWithOtherStatus.forEach(s => {
        statusCounts[s] = (statusCounts[s] || 0) + 1;
      });
      Logger.log(`  Status breakdown: ${JSON.stringify(statusCounts)}`);
    }
    
    // Final summary
    Logger.log(`\n=== SUMMARY ===`);
    Logger.log(`FEEDBACK COUNT:`);
    Logger.log(`  Expected: 33`);
    Logger.log(`  Raw count: ${rawCount}`);
    Logger.log(`  After deduplication: ${deduplicatedCount}`);
    Logger.log(`  Difference: ${positionFilteredCount - deduplicatedCount} rows lost in deduplication`);
    Logger.log(`\nPENDING COUNT:`);
    Logger.log(`  Feedback_status = REQUESTED (raw): ${requestedFeedbackCount}`);
    Logger.log(`  Status in [PENDING, INVITED, EMAIL SENT] (before dedup): ${pendingByStatusCount}`);
    Logger.log(`  Status in [PENDING, INVITED, EMAIL SENT] (after dedup): ${deduplicatedPendingCount}`);
    Logger.log(`  REQUESTED feedback with pending status: ${requestedWithPendingStatus}`);
    Logger.log(`  REQUESTED feedback with other status: ${requestedWithOtherStatus.length}`);
    
  } catch (error) {
    Logger.log(`ERROR in test: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
  }
}

// Helper function to get status rank (same as in main code)
function getStatusRank(status) {
  const statusUpper = status.toUpperCase();
  if (statusUpper === 'COMPLETED') return 1;
  if (statusUpper === 'SCHEDULED') return 2;
  if (statusUpper === 'PENDING' || statusUpper === 'INVITED' || statusUpper === 'EMAIL SENT') return 3;
  return 99; // Other statuses get lower priority
}

/**
 * Test function to debug pending count discrepancy
 * Tests rows with Feedback_status = REQUESTED vs interview status PENDING/INVITED/EMAIL SENT
 */
function testPendingCount() {
  try {
    Logger.log(`=== Testing Pending Count (Feedback_status = REQUESTED) ===`);
    
    // Get raw log data
    const logData = getLogSheetDataRB();
    if (!logData || !logData.rows || logData.rows.length === 0) {
      Logger.log(`ERROR: Could not get log data`);
      return;
    }
    
    Logger.log(`Total rows in log sheet: ${logData.rows.length}`);
    
    // Get column indices
    const recruiterIdx = logData.colIndices.hasOwnProperty('Recruiter_name') ? logData.colIndices['Recruiter_name'] : -1;
    const feedbackStatusIdx = logData.colIndices.hasOwnProperty('Feedback_status') ? logData.colIndices['Feedback_status'] : -1;
    const statusIdx = logData.colIndices['STATUS_COLUMN'];
    const profileIdIdx = logData.colIndices['Profile_id'];
    const positionIdIdx = logData.colIndices['Position_id'];
    
    Logger.log(`Column indices - Recruiter_name: ${recruiterIdx}, Feedback_status: ${feedbackStatusIdx}, STATUS_COLUMN: ${statusIdx}`);
    
    if (recruiterIdx === -1 || feedbackStatusIdx === -1 || statusIdx === -1) {
      Logger.log(`ERROR: Required columns not found`);
      return;
    }
    
    // Step 1: Count raw rows with Feedback_status = REQUESTED
    let requestedFeedbackCount = 0;
    let requestedRows = [];
    logData.rows.forEach((row, index) => {
      if (row.length > feedbackStatusIdx) {
        const feedbackStatus = row[feedbackStatusIdx] ? String(row[feedbackStatusIdx]).trim().toUpperCase() : '';
        if (feedbackStatus === 'REQUESTED') {
          requestedFeedbackCount++;
          const recruiter = row.length > recruiterIdx ? String(row[recruiterIdx]).trim() : '';
          const status = row.length > statusIdx ? String(row[statusIdx]).trim() : '';
          requestedRows.push({ rowIndex: index, recruiter: recruiter, status: status, profileId: row[profileIdIdx], positionId: row[positionIdIdx] });
        }
      }
    });
    
    Logger.log(`\nStep 1 - Raw count (Feedback_status = REQUESTED): ${requestedFeedbackCount}`);
    
    // Step 2: Count rows with Feedback_status = REQUESTED AND status = PENDING/INVITED/EMAIL SENT
    const PENDING_STATUSES = ['PENDING', 'INVITED', 'EMAIL SENT'];
    let requestedAndPendingCount = 0;
    let requestedButNotPending = [];
    
    requestedRows.forEach(rowData => {
      const row = logData.rows[rowData.rowIndex];
      const status = row.length > statusIdx ? String(row[statusIdx]).trim() : '';
      if (PENDING_STATUSES.includes(status)) {
        requestedAndPendingCount++;
      } else {
        requestedButNotPending.push({ recruiter: rowData.recruiter, status: status, profileId: rowData.profileId, positionId: rowData.positionId });
      }
    });
    
    Logger.log(`\nStep 2 - Count with Feedback_status = REQUESTED AND status in [PENDING, INVITED, EMAIL SENT]: ${requestedAndPendingCount}`);
    Logger.log(`Rows with REQUESTED feedback but different status: ${requestedButNotPending.length}`);
    
    if (requestedButNotPending.length > 0) {
      Logger.log(`\nSample rows with REQUESTED feedback but not PENDING status:`);
      requestedButNotPending.slice(0, 10).forEach((r, idx) => {
        Logger.log(`  ${idx + 1}. Recruiter: ${r.recruiter}, Status: ${r.status}, Profile_id: ${r.profileId}, Position_id: ${r.positionId}`);
      });
    }
    
    // Step 3: After time range filter
    const filteredData = filterDataByTimeRangeRB(logData.rows, logData.colIndices);
    let timeFilteredRequestedCount = 0;
    filteredData.forEach(row => {
      if (row.length > feedbackStatusIdx) {
        const feedbackStatus = row[feedbackStatusIdx] ? String(row[feedbackStatusIdx]).trim().toUpperCase() : '';
        if (feedbackStatus === 'REQUESTED') {
          timeFilteredRequestedCount++;
        }
      }
    });
    Logger.log(`\nStep 3 - After time range filter (REQUESTED): ${timeFilteredRequestedCount}`);
    
    // Step 4: After template filter
    const feedbackTemplateIndex = logData.colIndices.hasOwnProperty('Feedback_template_name') ? logData.colIndices['Feedback_template_name'] : -1;
    const excludedTemplates = ["AI Coding Interview Metrics Feedback Form", "AI Functional Interview Feedback Form"];
    let templateFilteredData = filteredData;
    if (feedbackTemplateIndex !== -1) {
      templateFilteredData = filteredData.filter(row => {
        if (row.length <= feedbackTemplateIndex) return true;
        const templateName = row[feedbackTemplateIndex] ? String(row[feedbackTemplateIndex]).trim() : '';
        return !excludedTemplates.includes(templateName);
      });
    }
    
    let templateFilteredRequestedCount = 0;
    templateFilteredData.forEach(row => {
      if (row.length > feedbackStatusIdx) {
        const feedbackStatus = row[feedbackStatusIdx] ? String(row[feedbackStatusIdx]).trim().toUpperCase() : '';
        if (feedbackStatus === 'REQUESTED') {
          templateFilteredRequestedCount++;
        }
      }
    });
    Logger.log(`\nStep 4 - After template filter (REQUESTED): ${templateFilteredRequestedCount}`);
    
    // Step 5: After position filter
    const positionNameIndex = logData.colIndices.hasOwnProperty('Position_name') ? logData.colIndices['Position_name'] : -1;
    const positionToExclude = "AIR Testing";
    let finalFilteredData = templateFilteredData;
    if (positionNameIndex !== -1) {
      finalFilteredData = templateFilteredData.filter(row => {
        return !(row.length > positionNameIndex && row[positionNameIndex] === positionToExclude);
      });
    }
    
    let positionFilteredRequestedCount = 0;
    finalFilteredData.forEach(row => {
      if (row.length > feedbackStatusIdx) {
        const feedbackStatus = row[feedbackStatusIdx] ? String(row[feedbackStatusIdx]).trim().toUpperCase() : '';
        if (feedbackStatus === 'REQUESTED') {
          positionFilteredRequestedCount++;
        }
      }
    });
    Logger.log(`\nStep 5 - After position filter (REQUESTED): ${positionFilteredRequestedCount}`);
    
    // Step 6: Count pending based on status (current logic)
    let pendingByStatusCount = 0;
    finalFilteredData.forEach(row => {
      if (row.length > statusIdx) {
        const status = String(row[statusIdx]).trim();
        if (PENDING_STATUSES.includes(status)) {
          pendingByStatusCount++;
        }
      }
    });
    Logger.log(`\nStep 6 - Pending count based on status [PENDING, INVITED, EMAIL SENT]: ${pendingByStatusCount}`);
    
    // Step 7: After deduplication
    const groupedData = {};
    finalFilteredData.forEach(row => {
      if (row && row.length > profileIdIdx && row.length > positionIdIdx && row.length > statusIdx) {
        const profileId = row[profileIdIdx];
        const positionId = row[positionIdIdx];
        if (profileId && positionId) {
          const key = `${profileId}_${positionId}`;
          const status = String(row[statusIdx]).trim();
          const statusRank = getStatusRank(status);
          
          if (!groupedData[key] || statusRank < groupedData[key].bestRank) {
            groupedData[key] = { bestRank: statusRank, row: row };
          }
        }
      }
    });
    
    const deduplicatedRows = Object.values(groupedData).map(item => item.row);
    let deduplicatedPendingCount = 0;
    deduplicatedRows.forEach(row => {
      if (row.length > statusIdx) {
        const status = String(row[statusIdx]).trim();
        if (PENDING_STATUSES.includes(status)) {
          deduplicatedPendingCount++;
        }
      }
    });
    Logger.log(`\nStep 7 - After deduplication (Pending by status): ${deduplicatedPendingCount}`);
    
    // Final summary
    Logger.log(`\n=== SUMMARY ===`);
    Logger.log(`Rows with Feedback_status = REQUESTED (raw): ${requestedFeedbackCount}`);
    Logger.log(`Rows with Feedback_status = REQUESTED AND status in [PENDING, INVITED, EMAIL SENT]: ${requestedAndPendingCount}`);
    Logger.log(`Rows with REQUESTED feedback but different status: ${requestedButNotPending.length}`);
    Logger.log(`Pending count by status (after all filters, before deduplication): ${pendingByStatusCount}`);
    Logger.log(`Pending count by status (after deduplication): ${deduplicatedPendingCount}`);
    Logger.log(`\nDifference: ${requestedFeedbackCount - deduplicatedPendingCount} rows`);
    
  } catch (error) {
    Logger.log(`ERROR in test: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
  }
}

function testValidationSheetFunctionality() {
  try {
    Logger.log(`--- Testing Complete Validation Sheet Functionality ---`);
    
    // Get application data
    const appData = getApplicationDataForChartRB();
    if (!appData || !appData.rows) {
      Logger.log(`ERROR: Could not get application data`);
      return;
    }
    
    Logger.log(`Got ${appData.rows.length} rows from application sheet`);
    
    // Test 1: Single recruiter validation sheet
    Logger.log(`\n--- Test 1: Single Recruiter Validation Sheet ---`);
    const testRecruiter = 'Akhila Kashyap';
    const singleSheetUrl = createRecruiterValidationSheet(testRecruiter, appData.rows, appData.colIndices);
    Logger.log(`SUCCESS: Single validation sheet created for ${testRecruiter}: ${singleSheetUrl}`);
    
    // Test 2: All recruiters validation sheets
    Logger.log(`\n--- Test 2: All Recruiters Validation Sheets ---`);
    const allSheetsData = createAllRecruiterValidationSheets(appData.rows, appData.colIndices);
    if (allSheetsData) {
      Logger.log(`SUCCESS: Created ${allSheetsData.successfulSheets} validation sheets out of ${allSheetsData.totalRecruiters} recruiters`);
      Logger.log(`Failed recruiters: ${allSheetsData.failedRecruiters.length > 0 ? allSheetsData.failedRecruiters.join(', ') : 'None'}`);
      
      // Log details for each recruiter
      Object.entries(allSheetsData.validationSheets).forEach(([recruiter, data]) => {
        Logger.log(`${recruiter}: ${data.eligible} eligible, ${data.aiDone} AI done, ${data.aiMissing} AI missing, ${data.percentage}% coverage`);
      });
    } else {
      Logger.log(`ERROR: Could not create all recruiter validation sheets`);
    }
    
    // Test 3: HTML generation
    Logger.log(`\n--- Test 3: HTML Generation ---`);
    if (allSheetsData) {
      const htmlContent = generateValidationSheetsHtml(allSheetsData);
      Logger.log(`SUCCESS: Generated HTML content (${htmlContent.length} characters)`);
      Logger.log(`HTML preview (first 500 chars): ${htmlContent.substring(0, 500)}...`);
    }
    
    Logger.log(`\n--- All Validation Sheet Tests Completed Successfully ---`);
    
  } catch (error) {
    Logger.log(`ERROR in comprehensive test: ${error.toString()}`);
  }
}

/**
 * Detailed debug function to analyze Akhila's data and understand the filtering issue
 */
function debugAkhilaData() {
  try {
    Logger.log(`=== DETAILED DEBUG: AKHILA'S DATA ===`);
    
    // Get application data
    const appData = getApplicationDataForChartRB();
    if (!appData || !appData.rows) {
      Logger.log(`ERROR: Could not get application data`);
      return;
    }
    
    Logger.log(`Total rows in dataset: ${appData.rows.length}`);
    Logger.log(`Column indices: ${JSON.stringify(appData.colIndices)}`);
    
    const { rows, colIndices } = appData;
    const { recruiterNameIdx, lastStageIdx, aiInterviewIdx, applicationTsIdx } = colIndices;
    
    // Set May 1st, 2025 as the cutoff date
    const mayFirst2025 = new Date('2025-05-01');
    mayFirst2025.setHours(0, 0, 0, 0);
    
    // Define eligible stages
    const eligibleStages = [
      'HIRING MANAGER SCREEN',
      'ASSESSMENT', 
      'ONSITE INTERVIEW',
      'FINAL INTERVIEW',
      'OFFER APPROVALS',
      'OFFER EXTENDED',
      'OFFER DECLINED',
      'PENDING START',
      'HIRED'
    ];
    
    // Track Akhila's data specifically
    let akhilaTotal = 0;
    let akhilaAfterMay1st = 0;
    let akhilaEligible = 0;
    let akhilaStages = {};
    let akhilaApplicationDates = [];
    
    // Analyze all rows for Akhila
    rows.forEach((row, index) => {
      const recruiterName = String(row[recruiterNameIdx] || '').trim();
      
      // Only process Akhila's data
      if (recruiterName.toLowerCase().includes('akhila')) {
        akhilaTotal++;
        
        const lastStage = String(row[lastStageIdx] || '').trim().toUpperCase();
        const aiInterview = String(row[aiInterviewIdx] || '').trim().toUpperCase();
        const applicationTs = applicationTsIdx !== -1 ? vsParseDateSafeRB(row[applicationTsIdx]) : null;
        
        // Track stages
        if (!akhilaStages[lastStage]) {
          akhilaStages[lastStage] = 0;
        }
        akhilaStages[lastStage]++;
        
        // Check Application_ts filter
        if (applicationTs && applicationTs >= mayFirst2025) {
          akhilaAfterMay1st++;
          akhilaApplicationDates.push(applicationTs.toISOString().split('T')[0]);
          
          // Check if eligible stage (case insensitive)
          if (eligibleStages.some(stage => stage.toUpperCase() === lastStage)) {
            akhilaEligible++;
            Logger.log(`AKHILA ELIGIBLE: Row ${index}, Stage: "${lastStage}", AI: "${aiInterview}", Date: ${applicationTs.toISOString().split('T')[0]}`);
          } else {
            Logger.log(`AKHILA NOT ELIGIBLE: Row ${index}, Stage: "${lastStage}", AI: "${aiInterview}", Date: ${applicationTs.toISOString().split('T')[0]}`);
          }
        } else {
          Logger.log(`AKHILA BEFORE MAY 1ST: Row ${index}, Stage: "${lastStage}", AI: "${aiInterview}", Date: ${applicationTs ? applicationTs.toISOString().split('T')[0] : 'N/A'}`);
        }
      }
    });
    
    Logger.log(`=== AKHILA SUMMARY ===`);
    Logger.log(`Total Akhila candidates: ${akhilaTotal}`);
    Logger.log(`Akhila candidates after May 1st, 2025: ${akhilaAfterMay1st}`);
    Logger.log(`Akhila eligible candidates: ${akhilaEligible}`);
    Logger.log(`Akhila stages breakdown: ${JSON.stringify(akhilaStages)}`);
    Logger.log(`Akhila application dates (after May 1st): ${akhilaApplicationDates.slice(0, 10).join(', ')}...`);
    
    // Also check for variations of Akhila's name
    Logger.log(`=== CHECKING FOR NAME VARIATIONS ===`);
    const nameVariations = {};
    rows.forEach((row, index) => {
      const recruiterName = String(row[recruiterNameIdx] || '').trim();
      if (recruiterName.toLowerCase().includes('akhila') || recruiterName.toLowerCase().includes('kashyap')) {
        if (!nameVariations[recruiterName]) {
          nameVariations[recruiterName] = 0;
        }
        nameVariations[recruiterName]++;
      }
    });
    Logger.log(`Name variations found: ${JSON.stringify(nameVariations)}`);
    
  } catch (error) {
    Logger.log(`ERROR in debugAkhilaData: ${error.toString()}`);
  }
}

/**
 * Function to list all current filters being applied in AI coverage calculation
 */
function listCurrentFilters() {
  Logger.log(`=== CURRENT FILTERS IN AI COVERAGE CALCULATION ===`);
  Logger.log(`1. Application_ts filter: Must be ≥ May 1st, 2025`);
  Logger.log(`2. Last_stage filter: Must be one of the following:`);
  Logger.log(`   - HIRING MANAGER SCREEN`);
  Logger.log(`   - ASSESSMENT`);
  Logger.log(`   - ONSITE INTERVIEW`);
  Logger.log(`   - FINAL INTERVIEW`);
  Logger.log(`   - OFFER APPROVALS`);
  Logger.log(`   - OFFER EXTENDED`);
  Logger.log(`   - OFFER DECLINED`);
  Logger.log(`   - PENDING START`);
  Logger.log(`   - HIRED`);
  Logger.log(`3. Recruiter name must not be empty`);
  Logger.log(`4. All required columns must have data (recruiter, last stage, AI interview, application timestamp)`);
  Logger.log(`5. Excluded recruiters: Samrudh J, Pavan Kumar, Guruprasad Hegde`);
  Logger.log(`=== END OF FILTERS ===`);
}