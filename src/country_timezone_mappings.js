/**
 * Comprehensive Country to Timezone Offset Mapping
 * 
 * Offset values are in hours from UTC
 * Positive = ahead of UTC (e.g., India +5.5)
 * Negative = behind UTC (e.g., US -5)
 * 
 * Add your countries here and copy to timezone_converter.gs
 */

const COMPREHENSIVE_COUNTRY_TIMEZONE_MAP = {
  // Asia
  'India': 5.5,
  'IN': 5.5,
  'IND': 5.5,
  
  'China': 8,
  'CN': 8,
  'CHN': 8,
  
  'Japan': 9,
  'JP': 9,
  'JPN': 9,
  
  'Singapore': 8,
  'SG': 8,
  'SGP': 8,
  
  'South Korea': 9,
  'Korea': 9,
  'KR': 9,
  'KOR': 9,
  
  'Thailand': 7,
  'TH': 7,
  'THA': 7,
  
  'Malaysia': 8,
  'MY': 8,
  'MYS': 8,
  
  'Indonesia': 7, // Western Indonesia (Jakarta)
  'ID': 7,
  'IDN': 7,
  
  'Philippines': 8,
  'PH': 8,
  'PHL': 8,
  
  'Vietnam': 7,
  'VN': 7,
  'VNM': 7,
  
  'Bangladesh': 6,
  'BD': 6,
  'BGD': 6,
  
  'Pakistan': 5,
  'PK': 5,
  'PAK': 5,
  
  'Sri Lanka': 5.5,
  'LK': 5.5,
  'LKA': 5.5,
  
  'Nepal': 5.75,
  'NP': 5.75,
  'NPL': 5.75,
  
  'United Arab Emirates': 4,
  'UAE': 4,
  'AE': 4,
  'ARE': 4,
  
  'Saudi Arabia': 3,
  'SA': 3,
  'SAU': 3,
  
  'Israel': 2,
  'IL': 2,
  'ISR': 2,
  
  'Turkey': 3,
  'TR': 3,
  'TUR': 3,
  
  // Europe
  'United Kingdom': 0,
  'UK': 0,
  'England': 0,
  'GB': 0,
  'GBR': 0,
  
  'Germany': 1,
  'DE': 1,
  'DEU': 1,
  
  'France': 1,
  'FR': 1,
  'FRA': 1,
  
  'Italy': 1,
  'IT': 1,
  'ITA': 1,
  
  'Spain': 1,
  'ES': 1,
  'ESP': 1,
  
  'Netherlands': 1,
  'NL': 1,
  'NLD': 1,
  
  'Belgium': 1,
  'BE': 1,
  'BEL': 1,
  
  'Switzerland': 1,
  'CH': 1,
  'CHE': 1,
  
  'Austria': 1,
  'AT': 1,
  'AUT': 1,
  
  'Sweden': 1,
  'SE': 1,
  'SWE': 1,
  
  'Norway': 1,
  'NO': 1,
  'NOR': 1,
  
  'Denmark': 1,
  'DK': 1,
  'DNK': 1,
  
  'Finland': 2,
  'FI': 2,
  'FIN': 2,
  
  'Poland': 1,
  'PL': 1,
  'POL': 1,
  
  'Portugal': 0,
  'PT': 0,
  'PRT': 0,
  
  'Greece': 2,
  'GR': 2,
  'GRC': 2,
  
  'Ireland': 0,
  'IE': 0,
  'IRL': 0,
  
  'Russia': 3, // Moscow time (Russia has multiple timezones)
  'RU': 3,
  'RUS': 3,
  
  // Americas
  'United States': -5, // EST default (US has multiple timezones)
  'USA': -5,
  'US': -5,
  'America': -5,
  'United States of America': -5,
  
  'Canada': -5, // EST default (Canada has multiple timezones)
  'CA': -5,
  'CAN': 5,
  
  'Mexico': -6, // CST default
  'MX': -6,
  'MEX': -6,
  
  'Brazil': -3, // Brasilia time
  'BR': -3,
  'BRA': -3,
  
  'Argentina': -3,
  'AR': -3,
  'ARG': -3,
  
  'Chile': -3,
  'CL': -3,
  'CHL': -3,
  
  'Colombia': -5,
  'CO': -5,
  'COL': -5,
  
  'Peru': -5,
  'PE': -5,
  'PER': -5,
  
  'Venezuela': -4,
  'VE': -4,
  'VEN': -4,
  
  // Oceania
  'Australia': 10, // AEST default (Australia has multiple timezones)
  'AU': 10,
  'AUS': 10,
  
  'New Zealand': 12,
  'NZ': 12,
  'NZL': 12,
  
  // Africa
  'South Africa': 2,
  'ZA': 2,
  'ZAF': 2,
  
  'Egypt': 2,
  'EG': 2,
  'EGY': 2,
  
  'Nigeria': 1,
  'NG': 1,
  'NGA': 1,
  
  'Kenya': 3,
  'KE': 3,
  'KEN': 3,
  
  // Default fallback
  'Unknown': 0,
  '': 0
};

/**
 * US State-specific timezone offsets (if you have state data)
 * Uncomment and use if you have a state/region column
 */
const US_STATE_TIMEZONE_MAP = {
  // Eastern Time (UTC-5)
  'Alabama': -5, 'AL': -5,
  'Connecticut': -5, 'CT': -5,
  'Delaware': -5, 'DE': -5,
  'Florida': -5, 'FL': -5, // Most of Florida
  'Georgia': -5, 'GA': -5,
  'Indiana': -5, 'IN': -5, // Most of Indiana
  'Kentucky': -5, 'KY': -5, // Eastern Kentucky
  'Maine': -5, 'ME': -5,
  'Maryland': -5, 'MD': -5,
  'Massachusetts': -5, 'MA': -5,
  'Michigan': -5, 'MI': -5,
  'New Hampshire': -5, 'NH': -5,
  'New Jersey': -5, 'NJ': -5,
  'New York': -5, 'NY': -5,
  'North Carolina': -5, 'NC': -5,
  'Ohio': -5, 'OH': -5,
  'Pennsylvania': -5, 'PA': -5,
  'Rhode Island': -5, 'RI': -5,
  'South Carolina': -5, 'SC': -5,
  'Tennessee': -5, 'TN': -5, // Eastern Tennessee
  'Vermont': -5, 'VT': -5,
  'Virginia': -5, 'VA': -5,
  'West Virginia': -5, 'WV': -5,
  'Washington DC': -5, 'DC': -5,
  
  // Central Time (UTC-6)
  'Arkansas': -6, 'AR': -6,
  'Illinois': -6, 'IL': -6,
  'Iowa': -6, 'IA': -6,
  'Kansas': -6, 'KS': -6, // Most of Kansas
  'Louisiana': -6, 'LA': -6,
  'Minnesota': -6, 'MN': -6,
  'Mississippi': -6, 'MS': -6,
  'Missouri': -6, 'MO': -6,
  'Nebraska': -6, 'NE': -6, // Most of Nebraska
  'North Dakota': -6, 'ND': -6, // Most of North Dakota
  'Oklahoma': -6, 'OK': -6,
  'South Dakota': -6, 'SD': -6, // Most of South Dakota
  'Texas': -6, 'TX': -6, // Most of Texas
  'Wisconsin': -6, 'WI': -6,
  
  // Mountain Time (UTC-7)
  'Arizona': -7, 'AZ': -7, // No DST
  'Colorado': -7, 'CO': -7,
  'Idaho': -7, 'ID': -7, // Southern Idaho
  'Montana': -7, 'MT': -7,
  'New Mexico': -7, 'NM': -7,
  'Utah': -7, 'UT': -7,
  'Wyoming': -7, 'WY': -7,
  
  // Pacific Time (UTC-8)
  'California': -8, 'CA': -8,
  'Nevada': -8, 'NV': -8,
  'Oregon': -8, 'OR': -8,
  'Washington': -8, 'WA': -8,
  
  // Alaska Time (UTC-9)
  'Alaska': -9, 'AK': -9,
  
  // Hawaii Time (UTC-10)
  'Hawaii': -10, 'HI': -10
};

/**
 * Canadian Province timezone offsets
 */
const CANADA_PROVINCE_TIMEZONE_MAP = {
  // Eastern Time
  'Ontario': -5, 'ON': -5, // Most of Ontario
  'Quebec': -5, 'QC': -5, // Most of Quebec
  
  // Central Time
  'Manitoba': -6, 'MB': -6,
  'Saskatchewan': -6, 'SK': -6, // Most of Saskatchewan
  
  // Mountain Time
  'Alberta': -7, 'AB': -7,
  'British Columbia': -8, 'BC': -8, // Most of BC
  
  // Atlantic Time
  'New Brunswick': -4, 'NB': -4,
  'Nova Scotia': -4, 'NS': -4,
  'Prince Edward Island': -4, 'PE': -4,
  'Newfoundland': -3.5, 'NL': -3.5
};

/**
 * Australian State timezone offsets
 */
const AUSTRALIA_STATE_TIMEZONE_MAP = {
  'New South Wales': 10, 'NSW': 10,
  'Victoria': 10, 'VIC': 10,
  'Queensland': 10, 'QLD': 10,
  'Tasmania': 10, 'TAS': 10,
  'South Australia': 9.5, 'SA': 9.5,
  'Western Australia': 8, 'WA': 8,
  'Northern Territory': 9.5, 'NT': 9.5,
  'Australian Capital Territory': 10, 'ACT': 10
};

