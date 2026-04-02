/**
 * Convert Microsoft timezone names to IANA timezone names
 */
function convertMicrosoftTimezone(msTimezone) {
  const timezoneMap = {
    // Europe
    'GMT Standard Time': 'Europe/London',
    'Greenwich Standard Time': 'Atlantic/Reykjavik',
    'W. Europe Standard Time': 'Europe/Berlin',
    'Central Europe Standard Time': 'Europe/Warsaw',
    'Central European Standard Time': 'Europe/Warsaw',
    'Romance Standard Time': 'Europe/Paris',
    'GTB Standard Time': 'Europe/Bucharest',
    'FLE Standard Time': 'Europe/Kiev',
    'E. Europe Standard Time': 'Europe/Bucharest',
    'Turkey Standard Time': 'Europe/Istanbul',
    'Kaliningrad Standard Time': 'Europe/Kaliningrad',
    'Russian Standard Time': 'Europe/Moscow',
    'Belarus Standard Time': 'Europe/Minsk',
    'Russia Time Zone 3': 'Europe/Samara',
    'Russia Time Zone 11': 'Asia/Kamchatka',

    // Africa
    'Morocco Standard Time': 'Africa/Casablanca',
    'South Africa Standard Time': 'Africa/Johannesburg',
    'E. Africa Standard Time': 'Africa/Nairobi',
    'W. Central Africa Standard Time': 'Africa/Lagos',
    'Egypt Standard Time': 'Africa/Cairo',
    'Namibia Standard Time': 'Africa/Windhoek',

    // Middle East / West Asia
    'Israel Standard Time': 'Asia/Jerusalem',
    'Jordan Standard Time': 'Asia/Amman',
    'Middle East Standard Time': 'Asia/Beirut',
    'Syria Standard Time': 'Asia/Damascus',
    'West Bank Standard Time': 'Asia/Hebron',
    'Arab Standard Time': 'Asia/Riyadh',
    'Arabian Standard Time': 'Asia/Dubai',
    'Iran Standard Time': 'Asia/Tehran',
    'Afghanistan Standard Time': 'Asia/Kabul',
    'Azerbaijan Standard Time': 'Asia/Baku',
    'Caucasus Standard Time': 'Asia/Yerevan',
    'Georgian Standard Time': 'Asia/Tbilisi',
    'Mauritius Standard Time': 'Indian/Mauritius',

    // Asia
    'Pakistan Standard Time': 'Asia/Karachi',
    'West Asia Standard Time': 'Asia/Tashkent',
    'Ekaterinburg Standard Time': 'Asia/Yekaterinburg',
    'India Standard Time': 'Asia/Kolkata',
    'Sri Lanka Standard Time': 'Asia/Colombo',
    'Nepal Standard Time': 'Asia/Kathmandu',
    'Bangladesh Standard Time': 'Asia/Dhaka',
    'N. Central Asia Standard Time': 'Asia/Novosibirsk',
    'Central Asia Standard Time': 'Asia/Almaty',
    'Myanmar Standard Time': 'Asia/Rangoon',
    'SE Asia Standard Time': 'Asia/Bangkok',
    'North Asia Standard Time': 'Asia/Krasnoyarsk',
    'China Standard Time': 'Asia/Shanghai',
    'Singapore Standard Time': 'Asia/Singapore',
    'Taipei Standard Time': 'Asia/Taipei',
    'Korea Standard Time': 'Asia/Seoul',
    'Tokyo Standard Time': 'Asia/Tokyo',
    'North Asia East Standard Time': 'Asia/Irkutsk',
    'Yakutsk Standard Time': 'Asia/Yakutsk',
    'Vladivostok Standard Time': 'Asia/Vladivostok',

    // Australia / Pacific
    'AUS Eastern Standard Time': 'Australia/Sydney',
    'E. Australia Standard Time': 'Australia/Brisbane',
    'AUS Central Standard Time': 'Australia/Darwin',
    'Cen. Australia Standard Time': 'Australia/Adelaide',
    'Tasmania Standard Time': 'Australia/Hobart',
    'West Pacific Standard Time': 'Pacific/Port_Moresby',
    'Central Pacific Standard Time': 'Pacific/Guadalcanal',
    'New Zealand Standard Time': 'Pacific/Auckland',
    'Fiji Standard Time': 'Pacific/Fiji',
    'Tonga Standard Time': 'Pacific/Tongatapu',
    'Samoa Standard Time': 'Pacific/Apia',

    // Americas
    'Hawaiian Standard Time': 'Pacific/Honolulu',
    'Alaskan Standard Time': 'America/Anchorage',
    'Pacific Standard Time': 'America/Los_Angeles',
    'Pacific Standard Time (Mexico)': 'America/Tijuana',
    'Mountain Standard Time': 'America/Denver',
    'Mountain Standard Time (Mexico)': 'America/Chihuahua',
    'US Mountain Standard Time': 'America/Phoenix',
    'Central Standard Time': 'America/Chicago',
    'Central Standard Time (Mexico)': 'America/Mexico_City',
    'Canada Central Standard Time': 'America/Regina',
    'W. Central America Standard Time': 'America/Guatemala',
    'Eastern Standard Time': 'America/New_York',
    'Eastern Standard Time (Mexico)': 'America/Cancun',
    'US Eastern Standard Time': 'America/Indianapolis',
    'SA Pacific Standard Time': 'America/Bogota',
    'Venezuela Standard Time': 'America/Caracas',
    'SA Western Standard Time': 'America/La_Paz',
    'Central Brazilian Standard Time': 'America/Cuiaba',
    'E. South America Standard Time': 'America/Sao_Paulo',
    'Argentina Standard Time': 'America/Argentina/Buenos_Aires',
    'SA Eastern Standard Time': 'America/Cayenne',
    'Greenland Standard Time': 'America/Godthab',
    'Montevideo Standard Time': 'America/Montevideo',
    'Bahia Standard Time': 'America/Bahia',
    'Newfoundland Standard Time': 'America/St_Johns',
    'Atlantic Standard Time': 'America/Halifax',

    // UTC variants
    'UTC': 'UTC',
    'GMT': 'UTC',
    'Dateline Standard Time': 'Etc/GMT+12',
    'UTC-11': 'Etc/GMT+11',
    'UTC-02': 'Etc/GMT+2',
    'UTC+12': 'Etc/GMT-12',
    'Line Islands Standard Time': 'Pacific/Kiritimati',
    'Azores Standard Time': 'Atlantic/Azores',
    'Cabo Verde Standard Time': 'Atlantic/Cape_Verde',
  };

  if (timezoneMap[msTimezone]) {
    return timezoneMap[msTimezone];
  }

  // Fallback: try the raw value in case it's already a valid IANA/Java timezone ID
  try {
    Utilities.formatDate(new Date(), msTimezone, 'z');
    return msTimezone;
  } catch (e) {
    console.warn(`Unknown timezone: ${msTimezone}, using UTC`);
    return 'UTC';
  }
}
