# Dynamic Timezone Converter - Setup Guide

## Overview
This Google Apps Script automatically converts UTC timestamps (`Schedule_start_time`) to local time based on the `Location_country` column. It runs automatically when you edit the sheet.

## Setup Instructions

### Step 1: Open Script Editor
1. Open your Google Sheet
2. Go to **Extensions** → **Apps Script**
3. Delete any default code
4. Copy and paste the code from `timezone_converter.gs`

### Step 2: Configure Column Numbers
Update the `CONFIG` object at the top of the script:

```javascript
const CONFIG = {
  SCHEDULE_START_TIME_COL: 13,  // Column M (Schedule_start_time)
  LOCATION_COUNTRY_COL: 21,     // Column U (Location_country)
  LOCAL_TIME_COL: 30,           // Column AD (where Local Time will be written)
  DATA_START_ROW: 2,             // Row where data starts (row 1 = headers)
  SHEET_NAME: 'Enriched'        // Your sheet name
};
```

**To find column numbers:**
- Column A = 1, B = 2, C = 3, ..., M = 13, U = 21, AD = 30
- Or use: `=COLUMN(M1)` in a cell to get column number

### Step 3: Add Country Mappings
Edit the `COUNTRY_TIMEZONE_MAP` object to add your countries:

```javascript
const COUNTRY_TIMEZONE_MAP = {
  'India': 5.5,              // IST (UTC+5:30)
  'United States': -5,       // EST (UTC-5) - adjust for other US timezones
  'United Kingdom': 0,        // GMT (UTC+0)
  // Add more countries...
};
```

**Timezone offsets:**
- Positive = ahead of UTC (e.g., India +5.5)
- Negative = behind UTC (e.g., US -5)
- 0 = same as UTC (e.g., UK)

### Step 4: Run Setup
1. In the script editor, select `setupTimezoneConverter` from the function dropdown
2. Click **Run** (▶️)
3. Authorize permissions when prompted
4. This will:
   - Create a "Local Time" header in your specified column
   - Add a custom menu to your sheet

### Step 5: Initial Conversion
1. In your Google Sheet, you'll see a new menu: **Timezone Converter**
2. Click **Timezone Converter** → **Convert All Timezones**
3. This will convert all existing rows

## How It Works

### Automatic Conversion (onEdit Trigger)
- When you edit `Schedule_start_time` (column M) or `Location_country` (column U)
- The script automatically converts that row's timestamp
- Runs instantly without manual action

### Manual Conversion
- Use the menu: **Timezone Converter** → **Convert All Timezones**
- Useful for bulk updates or after importing new data

## Adding More Countries

Edit the `COUNTRY_TIMEZONE_MAP` in the script:

```javascript
const COUNTRY_TIMEZONE_MAP = {
  // Existing entries...
  
  // Add new countries
  'Philippines': 8,           // UTC+8
  'South Africa': 2,          // UTC+2
  'Argentina': -3,            // UTC-3
  'New Zealand': 12,          // UTC+12
  // ... etc
};
```

**Note:** The script does case-insensitive and partial matching, so "United States", "USA", "US" all work.

## Handling Multiple Timezones in One Country

For countries with multiple timezones (like US), you have options:

### Option 1: Use Most Common Timezone
```javascript
'United States': -5,  // EST (most common)
```

### Option 2: Add State/Region Column
If you have a state/region column, you can enhance the script:

```javascript
function getTimezoneOffset(country, region) {
  if (country === 'United States') {
    if (region === 'California' || region === 'CA') return -8; // PST
    if (region === 'New York' || region === 'NY') return -5;   // EST
    // ... etc
  }
  // ... rest of logic
}
```

## Troubleshooting

### Script Not Running
1. Check that `onEdit` trigger is installed (should be automatic)
2. Go to **Triggers** in Apps Script (clock icon on left)
3. Verify `onEdit` trigger exists

### Wrong Timezone Conversion
1. Check country name in `Location_country` column matches exactly (case-insensitive)
2. Verify offset in `COUNTRY_TIMEZONE_MAP` is correct
3. Check logs: **View** → **Logs** in Apps Script

### Column Numbers Wrong
1. Verify column numbers in `CONFIG` object
2. Use `=COLUMN()` formula to check column numbers

### Performance Issues
- The script processes one row at a time on edit (efficient)
- For bulk updates, use the menu option instead of editing each row

## Testing

Run the test function:
1. In script editor, select `testTimezoneConversion`
2. Click **Run**
3. Check **View** → **Logs** for results

## Example Output

**Input:**
- Schedule_start_time: `2025-12-07 03:30:00` (UTC)
- Location_country: `India`

**Output:**
- Local Time: `2025-12-07 09:00:00` (IST, UTC+5:30)

## Advanced: Using Google's Timezone API

For more accurate timezone handling (including DST), you can enhance the script to use Google's Timezone API:

```javascript
function getTimezoneOffsetAdvanced(lat, lng, timestamp) {
  const apiKey = 'YOUR_API_KEY';
  const url = `https://maps.googleapis.com/maps/api/timezone/json?location=${lat},${lng}&timestamp=${timestamp/1000}&key=${apiKey}`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());
  return data.rawOffset / 3600; // Convert seconds to hours
}
```

This requires latitude/longitude data and a Google Maps API key.

