#!/usr/bin/env node
const fs = require('fs');
const http = require('http');
const https = require('https');
const path = require('path');
const ical = require('node-ical');

const WINDOWS_TIMEZONE_MAP = {
  'FLE Standard Time': 'Europe/Helsinki',
  'E. Europe Standard Time': 'Europe/Bucharest',
  'GTB Standard Time': 'Europe/Athens',
  'Central Europe Standard Time': 'Europe/Budapest',
  'W. Europe Standard Time': 'Europe/Berlin',
  'Pacific Standard Time': 'America/Los_Angeles',
  'Eastern Standard Time': 'America/New_York',
  'Central Standard Time': 'America/Chicago',
};
const IANA_TIMEZONE_ALIASES = {
  'Europe/Kiev': 'Europe/Kyiv',
};
const TIME_FORMAT_LOCALE = 'en-GB';
const formatterCache = new Map();
const weekdayFormatterCache = new Map();
const WEEKDAY_KEYS = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'];

function parseArgs(argv) {
  const args = {
    length: null,
    start: null,
    end: null,
    config: 'config.json',
    configProvided: false,
    debug: null,
    format: 'text',
    help: false,
  };

  for (let i = 2; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--help' || arg === '-h') {
      args.help = true;
      continue;
    }
    if (arg === '--length' || arg === '-l') {
      args.length = Number.parseInt(argv[i + 1], 10);
      i += 1;
      continue;
    }
    if (arg === '--start' || arg === '-s') {
      args.start = argv[i + 1];
      i += 1;
      continue;
    }
    if (arg === '--end' || arg === '-e') {
      args.end = argv[i + 1];
      i += 1;
      continue;
    }
    if (arg === '--config' || arg === '-c') {
      args.config = argv[i + 1];
      args.configProvided = true;
      i += 1;
      continue;
    }
    if (arg === '--format' || arg === '-f') {
      args.format = argv[i + 1];
      i += 1;
      continue;
    }
    if (arg === '--debug' || arg === '-d') {
      const next = argv[i + 1];
      if (next && !next.startsWith('-')) {
        args.debug = next;
        i += 1;
      } else {
        args.debug = true;
      }
      continue;
    }
    throw new Error(`Unknown argument: ${arg}`);
  }

  return args;
}

function resolveDefaultConfigPath() {
  const envPath = process.env.OUTLOOK_FREE_TIME_CONFIG;
  if (envPath) {
    return path.resolve(envPath);
  }
  const cwdConfig = path.resolve(process.cwd(), 'config.json');
  if (fs.existsSync(cwdConfig)) {
    return cwdConfig;
  }
  const home = process.env.USERPROFILE || process.env.HOME;
  if (home) {
    const homeConfig = path.resolve(home, '.outlook-free-time.json');
    if (fs.existsSync(homeConfig)) {
      return homeConfig;
    }
  }
  return cwdConfig;
}

function showHelp() {
  const helpText = [
    'Usage:',
    '  node src/index.js --length 30 --start 14.1 --end 16.1',
    '',
    'Options:',
    '  --length, -l   Meeting length in minutes',
    '  --start, -s    Start date (DD.M, DD.MM.YYYY, or YYYY-MM-DD)',
    '  --end, -e      End date (DD.M, DD.MM.YYYY, or YYYY-MM-DD)',
    '  --config, -c   Path to config.json (default: config.json or OUTLOOK_FREE_TIME_CONFIG)',
    '  --format, -f   Output format: text, list, block, json (default: text)',
    '  --debug, -d    Print busy intervals for each day (optionally pass a date)',
  ];
  console.log(helpText.join('\n'));
}

function readConfig(configPath) {
  const resolvedPath = path.resolve(configPath);
  const raw = fs.readFileSync(resolvedPath, 'utf8');
  const config = JSON.parse(raw);
  return {
    ...config,
    configPath: resolvedPath,
    configDir: path.dirname(resolvedPath),
  };
}

function normalizeTimeZone(rawValue) {
  if (rawValue === undefined || rawValue === null || rawValue === '') {
    return null;
  }
  const input = String(rawValue).trim();
  if (!input) {
    return null;
  }
  const mapped =
    WINDOWS_TIMEZONE_MAP[input] || IANA_TIMEZONE_ALIASES[input] || input;
  try {
    new Intl.DateTimeFormat(TIME_FORMAT_LOCALE, { timeZone: mapped });
  } catch (error) {
    throw new Error(`Unsupported timeZone value: ${input}`);
  }
  return mapped;
}

function resolveEventTimeZone(event) {
  const raw =
    (event && event.start && event.start.tz) ||
    (event && event.rrule && event.rrule.origOptions && event.rrule.origOptions.tzid) ||
    (event && event.rrule && event.rrule.options && event.rrule.options.tzid) ||
    (event && event.tz) ||
    null;
  if (!raw) {
    return null;
  }
  try {
    return normalizeTimeZone(raw);
  } catch (error) {
    return null;
  }
}

function normalizeWeekdayKey(rawKey) {
  if (!rawKey) {
    return null;
  }
  const value = String(rawKey).trim().toUpperCase();
  if (!value) {
    return null;
  }
  if (value.length === 2) {
    const map = {
      MO: 'MON',
      TU: 'TUE',
      WE: 'WED',
      TH: 'THU',
      FR: 'FRI',
      SA: 'SAT',
      SU: 'SUN',
    };
    return map[value] || null;
  }
  if (value.startsWith('MON')) return 'MON';
  if (value.startsWith('TUE')) return 'TUE';
  if (value.startsWith('WED')) return 'WED';
  if (value.startsWith('THU') || value.startsWith('THR')) return 'THU';
  if (value.startsWith('FRI')) return 'FRI';
  if (value.startsWith('SAT')) return 'SAT';
  if (value.startsWith('SUN')) return 'SUN';
  return null;
}

function getFormatter(timeZone) {
  if (!timeZone) {
    return null;
  }
  if (!formatterCache.has(timeZone)) {
    formatterCache.set(
      timeZone,
      new Intl.DateTimeFormat(TIME_FORMAT_LOCALE, {
        timeZone,
        hour12: false,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
      }),
    );
  }
  return formatterCache.get(timeZone);
}

function getWeekdayFormatter(timeZone) {
  const key = timeZone || 'local';
  if (!weekdayFormatterCache.has(key)) {
    weekdayFormatterCache.set(
      key,
      new Intl.DateTimeFormat('en-US', { timeZone, weekday: 'short' }),
    );
  }
  return weekdayFormatterCache.get(key);
}

function getWeekdayKey(date, timeZone) {
  if (!timeZone) {
    return WEEKDAY_KEYS[date.getDay()];
  }
  const formatter = getWeekdayFormatter(timeZone);
  const label = formatter.format(date);
  return normalizeWeekdayKey(label) || WEEKDAY_KEYS[date.getDay()];
}

function getTimeZoneParts(date, timeZone) {
  const formatter = getFormatter(timeZone);
  if (!formatter) {
    return null;
  }
  const parts = formatter.formatToParts(date);
  const values = {};
  for (const part of parts) {
    if (part.type !== 'literal') {
      values[part.type] = part.value;
    }
  }
  return {
    year: Number(values.year),
    month: Number(values.month),
    day: Number(values.day),
    hour: Number(values.hour),
    minute: Number(values.minute),
    second: Number(values.second),
  };
}

function getTimeZoneOffset(date, timeZone) {
  const parts = getTimeZoneParts(date, timeZone);
  if (!parts) {
    return 0;
  }
  const asUtc = Date.UTC(
    parts.year,
    parts.month - 1,
    parts.day,
    parts.hour,
    parts.minute,
    parts.second,
  );
  return asUtc - date.getTime();
}

function makeDateInTimeZone(year, month, day, hour, minute, second, timeZone) {
  const utcGuess = new Date(Date.UTC(year, month - 1, day, hour, minute, second));
  if (!timeZone) {
    return utcGuess;
  }
  const offset = getTimeZoneOffset(utcGuess, timeZone);
  return new Date(utcGuess.getTime() - offset);
}

function alignToEventWallTime(date, eventStart, timeZone) {
  if (!timeZone || !(eventStart instanceof Date)) {
    return date;
  }
  const baseParts = getTimeZoneParts(eventStart, timeZone);
  const occurrenceParts = getTimeZoneParts(date, timeZone);
  if (!baseParts || !occurrenceParts) {
    return date;
  }
  return makeDateInTimeZone(
    occurrenceParts.year,
    occurrenceParts.month,
    occurrenceParts.day,
    baseParts.hour,
    baseParts.minute,
    baseParts.second,
    timeZone,
  );
}

function normalizeCacheMaxAgeMinutes(rawValue) {
  if (rawValue === undefined || rawValue === null || rawValue === '') {
    return 5;
  }
  const parsed = Number.parseInt(rawValue, 10);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error('cacheMaxAgeMinutes must be a positive integer');
  }
  return parsed;
}

function resolveCachePath(config) {
  if (config.icsCacheFile) {
    return path.resolve(config.configDir, config.icsCacheFile);
  }
  return path.resolve(config.configDir, '.cache', 'calendar.ics');
}

function isCacheFresh(filePath, maxAgeMinutes) {
  try {
    const stats = fs.statSync(filePath);
    const ageMs = Date.now() - stats.mtimeMs;
    return ageMs >= 0 && ageMs < maxAgeMinutes * 60 * 1000;
  } catch {
    return false;
  }
}

function downloadFile(url, filePath, redirectCount = 0) {
  if (redirectCount > 5) {
    return Promise.reject(new Error('Too many redirects while downloading ICS'));
  }
  const client = url.startsWith('https:') ? https : http;
  return new Promise((resolve, reject) => {
    const request = client.get(url, (res) => {
      const status = res.statusCode || 0;
      if (status >= 300 && status < 400 && res.headers.location) {
        res.resume();
        const nextUrl = new URL(res.headers.location, url).toString();
        downloadFile(nextUrl, filePath, redirectCount + 1).then(resolve).catch(reject);
        return;
      }
      if (status < 200 || status >= 300) {
        res.resume();
        reject(new Error(`Failed to download ICS (${status})`));
        return;
      }
      fs.mkdirSync(path.dirname(filePath), { recursive: true });
      const tempPath = `${filePath}.tmp`;
      const fileStream = fs.createWriteStream(tempPath);
      res.pipe(fileStream);
      fileStream.on('finish', () => {
        fileStream.close(() => {
          fs.renameSync(tempPath, filePath);
          resolve();
        });
      });
      fileStream.on('error', (error) => {
        res.resume();
        try {
          fs.unlinkSync(tempPath);
        } catch {
          // Ignore cleanup failures.
        }
        reject(error);
      });
    });
    request.on('error', reject);
  });
}

async function ensureCachedIcs(config) {
  const cachePath = resolveCachePath(config);
  const maxAgeMinutes = normalizeCacheMaxAgeMinutes(config.cacheMaxAgeMinutes);
  if (isCacheFresh(cachePath, maxAgeMinutes)) {
    return cachePath;
  }
  await downloadFile(config.icsUrl, cachePath);
  return cachePath;
}

function parseDateInput(input) {
  const value = String(input || '').trim();
  if (!value) {
    throw new Error('Date value is required');
  }

  const isoMatch = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(value);
  if (isoMatch) {
    const year = Number.parseInt(isoMatch[1], 10);
    const month = Number.parseInt(isoMatch[2], 10);
    const day = Number.parseInt(isoMatch[3], 10);
    return buildDate(year, month, day);
  }

  const dotMatch = /^(\d{1,2})\.(\d{1,2})(?:\.(\d{4}))?$/.exec(value);
  if (dotMatch) {
    const day = Number.parseInt(dotMatch[1], 10);
    const month = Number.parseInt(dotMatch[2], 10);
    const year = dotMatch[3]
      ? Number.parseInt(dotMatch[3], 10)
      : new Date().getFullYear();
    return buildDate(year, month, day);
  }

  throw new Error(`Unsupported date format: ${value}`);
}

function buildDate(year, month, day) {
  const date = new Date(year, month - 1, day, 0, 0, 0, 0);
  if (
    Number.isNaN(date.getTime()) ||
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    throw new Error(`Invalid date: ${day}.${month}.${year}`);
  }
  return date;
}

function parseTimeOfDay(value) {
  const match = /^([01]?\d|2[0-3]):([0-5]\d)$/.exec(String(value || '').trim());
  if (!match) {
    throw new Error(`Invalid time value: ${value}`);
  }
  return {
    hours: Number.parseInt(match[1], 10),
    minutes: Number.parseInt(match[2], 10),
  };
}

function withTime(date, time) {
  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    time.hours,
    time.minutes,
    0,
    0,
  );
}

function pad2(value) {
  return String(value).padStart(2, '0');
}

function formatTime(date) {
  return `${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
}

function formatTimeInZone(date, timeZone) {
  if (!timeZone) {
    return formatTime(date);
  }
  const parts = getTimeZoneParts(date, timeZone);
  if (!parts) {
    return formatTime(date);
  }
  return `${pad2(parts.hour)}:${pad2(parts.minute)}`;
}

function formatDateLabel(date) {
  return `${date.getDate()}.${date.getMonth() + 1}`;
}

function formatDateLabelInZone(date, timeZone) {
  if (!timeZone) {
    return formatDateLabel(date);
  }
  const parts = getTimeZoneParts(date, timeZone);
  if (!parts) {
    return formatDateLabel(date);
  }
  return `${parts.day}.${parts.month}`;
}

function formatDateIsoInZone(date, timeZone) {
  if (!timeZone) {
    return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
  }
  const parts = getTimeZoneParts(date, timeZone);
  if (!parts) {
    return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
  }
  return `${parts.year}-${pad2(parts.month)}-${pad2(parts.day)}`;
}

function normalizeTimeGridMinutes(rawValue) {
  if (rawValue === undefined || rawValue === null || rawValue === '') {
    return 30;
  }
  const parsed = Number.parseInt(rawValue, 10);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error('timeGridMinutes must be a positive integer');
  }
  return parsed;
}

function ceilToGrid(date, gridMinutes, timeZone) {
  if (!timeZone) {
    const aligned = new Date(date.getTime());
    aligned.setSeconds(0, 0);
    const minutes = aligned.getHours() * 60 + aligned.getMinutes();
    const remainder = minutes % gridMinutes;
    if (remainder === 0) {
      return aligned;
    }
    aligned.setMinutes(aligned.getMinutes() + (gridMinutes - remainder));
    return aligned;
  }
  const parts = getTimeZoneParts(date, timeZone);
  if (!parts) {
    return ceilToGrid(date, gridMinutes, null);
  }
  const minutes = parts.hour * 60 + parts.minute;
  const remainder = minutes % gridMinutes;
  const delta = remainder === 0 ? 0 : gridMinutes - remainder;
  const base = makeDateInTimeZone(parts.year, parts.month, parts.day, 0, 0, 0, timeZone);
  return new Date(base.getTime() + (minutes + delta) * 60 * 1000);
}

function floorToGrid(date, gridMinutes, timeZone) {
  if (!timeZone) {
    const aligned = new Date(date.getTime());
    aligned.setSeconds(0, 0);
    const minutes = aligned.getHours() * 60 + aligned.getMinutes();
    const remainder = minutes % gridMinutes;
    if (remainder === 0) {
      return aligned;
    }
    aligned.setMinutes(aligned.getMinutes() - remainder);
    return aligned;
  }
  const parts = getTimeZoneParts(date, timeZone);
  if (!parts) {
    return floorToGrid(date, gridMinutes, null);
  }
  const minutes = parts.hour * 60 + parts.minute;
  const remainder = minutes % gridMinutes;
  const base = makeDateInTimeZone(parts.year, parts.month, parts.day, 0, 0, 0, timeZone);
  return new Date(base.getTime() + (minutes - remainder) * 60 * 1000);
}

function withTimeInZone(day, time, timeZone) {
  if (!timeZone) {
    return withTime(day, time);
  }
  return makeDateInTimeZone(
    day.getFullYear(),
    day.getMonth() + 1,
    day.getDate(),
    time.hours,
    time.minutes,
    0,
    timeZone,
  );
}

function isSameDay(left, right) {
  return (
    left.getFullYear() === right.getFullYear() &&
    left.getMonth() === right.getMonth() &&
    left.getDate() === right.getDate()
  );
}

function* iterateDays(startDate, endDate) {
  const cursor = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const last = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  while (cursor <= last) {
    yield new Date(cursor);
    cursor.setDate(cursor.getDate() + 1);
  }
}

async function loadCalendar(config) {
  if (config.icsFile) {
    const filePath = path.resolve(config.configDir, config.icsFile);
    return ical.sync.parseFile(filePath);
  }
  if (config.icsUrl) {
    const cachePath = await ensureCachedIcs(config);
    return ical.sync.parseFile(cachePath);
  }
  throw new Error('config.json must include icsUrl or icsFile');
}

function buildKeyCandidates(date) {
  const utcIso = date.toISOString();
  const utcIsoNoMs = utcIso.replace('.000Z', 'Z');
  const utcY = date.getUTCFullYear();
  const utcM = pad2(date.getUTCMonth() + 1);
  const utcD = pad2(date.getUTCDate());
  const utcH = pad2(date.getUTCHours());
  const utcMin = pad2(date.getUTCMinutes());
  const utcS = pad2(date.getUTCSeconds());
  const utcDateOnly = `${utcY}-${utcM}-${utcD}`;
  const utcCompact = `${utcY}${utcM}${utcD}T${utcH}${utcMin}${utcS}Z`;

  const localY = date.getFullYear();
  const localM = pad2(date.getMonth() + 1);
  const localD = pad2(date.getDate());
  const localH = pad2(date.getHours());
  const localMin = pad2(date.getMinutes());
  const localS = pad2(date.getSeconds());
  const localDateOnly = `${localY}-${localM}-${localD}`;
  const localDateTime = `${localDateOnly}T${localH}:${localMin}:${localS}`;
  const localCompact = `${localY}${localM}${localD}T${localH}${localMin}${localS}`;

  return [
    utcIso,
    utcIsoNoMs,
    utcDateOnly,
    utcCompact,
    localDateOnly,
    localDateTime,
    localCompact,
  ];
}

function normalizeKeyDigits(value) {
  if (!value) {
    return '';
  }
  const text = String(value);
  const digits = text.replace(/\D/g, '');
  if (!digits) {
    return '';
  }
  if (text.includes('T') || digits.length > 8) {
    if (digits.length >= 14) {
      return digits.slice(0, 14);
    }
    if (digits.length >= 12) {
      return digits.slice(0, 12);
    }
  }
  if (digits.length >= 8) {
    return digits.slice(0, 8);
  }
  return digits;
}

function buildCandidateDigitSet(date) {
  const candidates = buildKeyCandidates(date);
  const set = new Set();
  for (const candidate of candidates) {
    const digits = normalizeKeyDigits(candidate);
    if (!digits) {
      continue;
    }
    if (digits.length === 14) {
      set.add(digits);
      if (digits.endsWith('00')) {
        set.add(digits.slice(0, 12));
      }
    } else if (digits.length === 12) {
      set.add(digits);
      set.add(`${digits}00`);
    } else if (digits.length === 8) {
      set.add(digits);
    } else {
      set.add(digits);
    }
  }
  return set;
}

function findMapEntry(map, date) {
  if (!map) {
    return null;
  }
  const keys = buildKeyCandidates(date);
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(map, key)) {
      return map[key];
    }
  }
  const candidateDigits = buildCandidateDigitSet(date);
  for (const [key, value] of Object.entries(map)) {
    const keyDigits = normalizeKeyDigits(key);
    if (!keyDigits) {
      continue;
    }
    if (candidateDigits.has(keyDigits)) {
      return value;
    }
  }
  return null;
}

function mapHasDate(map, date) {
  return Boolean(findMapEntry(map, date));
}

function extractRDates(event) {
  const rdate = event.rdate;
  if (!rdate) {
    return [];
  }
  const dates = [];
  const pushDate = (value) => {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      dates.push(value);
    }
  };

  if (Array.isArray(rdate)) {
    for (const value of rdate) {
      if (Array.isArray(value)) {
        for (const nested of value) {
          pushDate(nested);
        }
      } else {
        pushDate(value);
      }
    }
  } else if (typeof rdate === 'object') {
    for (const value of Object.values(rdate)) {
      if (Array.isArray(value)) {
        for (const nested of value) {
          pushDate(nested);
        }
      } else {
        pushDate(value);
      }
    }
  } else {
    pushDate(rdate);
  }

  return dates;
}

function getEventDurationMs(event) {
  if (event.start instanceof Date && event.end instanceof Date) {
    return event.end.getTime() - event.start.getTime();
  }
  if (event.duration && typeof event.duration.asMilliseconds === 'function') {
    return event.duration.asMilliseconds();
  }
  if (event.duration && typeof event.duration.toSeconds === 'function') {
    return event.duration.toSeconds() * 1000;
  }
  return 0;
}

function resolveEventTime(event) {
  if (!event || !(event.start instanceof Date) || Number.isNaN(event.start.getTime())) {
    return null;
  }
  let end = event.end instanceof Date ? event.end : null;
  if (!end || Number.isNaN(end.getTime())) {
    const durationMs = getEventDurationMs(event);
    if (Number.isFinite(durationMs) && durationMs > 0) {
      end = new Date(event.start.getTime() + durationMs);
    }
  }
  if (!end || Number.isNaN(end.getTime())) {
    return null;
  }
  return { start: event.start, end };
}

function getEventLabel(event) {
  return event.summary || event.description || event.uid || 'Busy';
}

function normalizeIgnoreSummaries(rawIgnore) {
  const defaults = ['vapaa'];
  if (!rawIgnore) {
    return new Set(defaults);
  }
  const items = Array.isArray(rawIgnore) ? rawIgnore : [rawIgnore];
  const normalized = defaults.slice();
  for (const item of items) {
    const value = String(item || '').trim().toLowerCase();
    if (value) {
      normalized.push(value);
    }
  }
  return new Set(normalized);
}

function isIgnoredSummary(summary, ignoreSummaries) {
  if (!ignoreSummaries || ignoreSummaries.size === 0) {
    return false;
  }
  const value = String(summary || '').trim().toLowerCase();
  return value ? ignoreSummaries.has(value) : false;
}

function expandRecurring(event, rangeStart, rangeEnd) {
  const durationMs = getEventDurationMs(event);
  if (!Number.isFinite(durationMs) || durationMs <= 0) {
    return [];
  }
  const dates = event.rrule ? event.rrule.between(rangeStart, rangeEnd, true) : [];
  const rdates = extractRDates(event);
  const allDates = dates.concat(rdates);
  const instances = [];
  const seen = new Set();
  const label = getEventLabel(event);
  const eventTimeZone = resolveEventTimeZone(event);

  for (const date of allDates) {
    if (!(date instanceof Date)) {
      continue;
    }
    const time = date.getTime();
    if (Number.isNaN(time) || seen.has(time)) {
      continue;
    }
    seen.add(time);
    const adjustedDate = alignToEventWallTime(date, event.start, eventTimeZone);
    const override =
      findMapEntry(event.recurrences, adjustedDate) ||
      findMapEntry(event.recurrences, date);
    if (override) {
      if (override.status === 'CANCELLED') {
        continue;
      }
      const overrideRange = resolveEventTime(override);
      if (overrideRange) {
        instances.push({ ...overrideRange, label });
      }
      continue;
    }
    if (mapHasDate(event.exdate, adjustedDate) || mapHasDate(event.exdate, date)) {
      continue;
    }
    instances.push({
      start: adjustedDate,
      end: new Date(adjustedDate.getTime() + durationMs),
      label,
    });
  }

  return instances;
}

function collectEventInstances(calendarData, rangeStart, rangeEnd, ignoreSummaries) {
  const instances = [];
  const entries = Object.values(calendarData || {});

  for (const entry of entries) {
    if (!entry) {
      continue;
    }

    if (entry.type === 'VFREEBUSY') {
      const blocks = Array.isArray(entry.freebusy) ? entry.freebusy : [];
      for (const block of blocks) {
        if (!block || !block.start || !block.end) {
          continue;
        }
        const type = String(block.type || '').toUpperCase();
        if (type && !type.includes('BUSY')) {
          continue;
        }
        if (block.end <= rangeStart || block.start >= rangeEnd) {
          continue;
        }
        const label = type ? `VFREEBUSY:${type}` : 'VFREEBUSY';
        instances.push({ start: block.start, end: block.end, label });
      }
      continue;
    }

    if (entry.type !== 'VEVENT') {
      continue;
    }
    if (entry.status === 'CANCELLED') {
      continue;
    }
    if (isIgnoredSummary(entry.summary, ignoreSummaries)) {
      continue;
    }

    const label = getEventLabel(entry);
    let eventInstances = [];
    if (entry.rrule || entry.rdate) {
      eventInstances = expandRecurring(entry, rangeStart, rangeEnd);
    } else {
      const baseRange = resolveEventTime(entry);
      if (baseRange) {
        eventInstances = [{ ...baseRange, label }];
      }
    }

    for (const instance of eventInstances) {
      if (!instance.start || !instance.end) {
        continue;
      }
      if (instance.end <= rangeStart || instance.start >= rangeEnd) {
        continue;
      }
      instances.push({
        start: instance.start,
        end: instance.end,
        label: instance.label || label,
      });
    }
  }

  return instances;
}

function mergeIntervals(intervals) {
  if (intervals.length === 0) {
    return [];
  }
  const sorted = intervals
    .slice()
    .sort((a, b) => a.start.getTime() - b.start.getTime());
  const merged = [{ start: sorted[0].start, end: sorted[0].end }];

  for (let i = 1; i < sorted.length; i += 1) {
    const current = sorted[i];
    const last = merged[merged.length - 1];
    if (current.start <= last.end) {
      if (current.end > last.end) {
        last.end = current.end;
      }
    } else {
      merged.push({ start: current.start, end: current.end });
    }
  }

  return merged;
}

function collectBusyIntervalsForDay(instances, dayStart, dayEnd, extraBusyIntervals) {
  const busyIntervals = [];
  for (const instance of instances) {
    if (instance.end <= dayStart || instance.start >= dayEnd) {
      continue;
    }
    const start = instance.start > dayStart ? instance.start : dayStart;
    const end = instance.end < dayEnd ? instance.end : dayEnd;
    if (end > start) {
      busyIntervals.push({ start, end, label: instance.label });
    }
  }

  for (const interval of extraBusyIntervals || []) {
    if (interval.end <= dayStart || interval.start >= dayEnd) {
      continue;
    }
    const start = interval.start > dayStart ? interval.start : dayStart;
    const end = interval.end < dayEnd ? interval.end : dayEnd;
    if (end > start) {
      busyIntervals.push({ start, end, label: interval.label || 'excludeTime' });
    }
  }

  return busyIntervals;
}

function getFreeSlotsForDay(instances, dayStart, dayEnd, minSlotMs, extraBusyIntervals) {
  const busyIntervals = collectBusyIntervalsForDay(
    instances,
    dayStart,
    dayEnd,
    extraBusyIntervals,
  );
  const mergedBusy = mergeIntervals(
    busyIntervals.map((interval) => ({ start: interval.start, end: interval.end })),
  );
  const free = [];
  let cursor = dayStart;

  for (const busy of mergedBusy) {
    if (busy.start > cursor && busy.start.getTime() - cursor.getTime() >= minSlotMs) {
      free.push({ start: cursor, end: busy.start });
    }
    if (busy.end > cursor) {
      cursor = busy.end;
    }
  }

  if (dayEnd > cursor && dayEnd.getTime() - cursor.getTime() >= minSlotMs) {
    free.push({ start: cursor, end: dayEnd });
  }

  return free;
}

function alignFreeSlotsToGrid(freeSlots, meetingLengthMs, gridMinutes, timeZone) {
  const aligned = [];
  for (const slot of freeSlots) {
    const earliestStart = ceilToGrid(slot.start, gridMinutes, timeZone);
    const latestStartCandidate = new Date(slot.end.getTime() - meetingLengthMs);
    const latestStart = floorToGrid(latestStartCandidate, gridMinutes, timeZone);
    if (latestStart < earliestStart) {
      continue;
    }
    aligned.push({
      start: earliestStart,
      end: new Date(latestStart.getTime() + meetingLengthMs),
    });
  }
  return aligned;
}

function normalizeExcludeTime(rawExclude) {
  if (!rawExclude) {
    return [];
  }
  const items = Array.isArray(rawExclude) ? rawExclude : [rawExclude];
  return items.map((item, index) => {
    if (!item || typeof item !== 'object') {
      throw new Error(`excludeTime entry ${index + 1} must be an object`);
    }
    if (!item.start || !item.end) {
      throw new Error(`excludeTime entry ${index + 1} must include start and end`);
    }
    const start = parseTimeOfDay(item.start);
    const end = parseTimeOfDay(item.end);
    if (end.hours < start.hours || (end.hours === start.hours && end.minutes <= start.minutes)) {
      throw new Error(`excludeTime entry ${index + 1} end must be after start`);
    }
    return { start, end };
  });
}

function normalizeWeeklyExcludeTime(rawExclude) {
  if (!rawExclude) {
    return {};
  }
  if (typeof rawExclude !== 'object' || Array.isArray(rawExclude)) {
    throw new Error('excludeTimeWeekly must be an object keyed by weekday');
  }
  const result = {};
  for (const [rawKey, rawValue] of Object.entries(rawExclude)) {
    const weekdayKey = normalizeWeekdayKey(rawKey);
    if (!weekdayKey) {
      throw new Error(`Invalid excludeTimeWeekly weekday: ${rawKey}`);
    }
    const items = Array.isArray(rawValue) ? rawValue : [rawValue];
    const windows = items.map((item, index) => {
      if (!item || typeof item !== 'object') {
        throw new Error(
          `excludeTimeWeekly ${weekdayKey} entry ${index + 1} must be an object`,
        );
      }
      if (!item.start || !item.end) {
        throw new Error(
          `excludeTimeWeekly ${weekdayKey} entry ${index + 1} must include start and end`,
        );
      }
      const start = parseTimeOfDay(item.start);
      const end = parseTimeOfDay(item.end);
      if (end.hours < start.hours || (end.hours === start.hours && end.minutes <= start.minutes)) {
        throw new Error(
          `excludeTimeWeekly ${weekdayKey} entry ${index + 1} end must be after start`,
        );
      }
      return { start, end };
    });
    result[weekdayKey] = windows;
  }
  return result;
}

async function main() {
  const args = parseArgs(process.argv);
  if (args.help) {
    showHelp();
    return;
  }

  if (!args.length || !args.start || !args.end) {
    showHelp();
    throw new Error('Missing required arguments');
  }

  const meetingLengthMinutes = args.length;
  if (!Number.isFinite(meetingLengthMinutes) || meetingLengthMinutes <= 0) {
    throw new Error('Meeting length must be a positive number of minutes');
  }

  const configPath = args.configProvided
    ? path.resolve(args.config)
    : resolveDefaultConfigPath();
  if (!fs.existsSync(configPath)) {
    throw new Error(
      `Config not found at ${configPath}. Use --config or set OUTLOOK_FREE_TIME_CONFIG.`,
    );
  }
  const config = readConfig(configPath);
  const startDate = parseDateInput(args.start);
  const endDate = parseDateInput(args.end);
  const timeZone = normalizeTimeZone(config.timeZone);

  if (startDate > endDate) {
    throw new Error('Start date must be before or equal to end date');
  }

  const outputFormat = String(args.format || 'text').trim().toLowerCase();
  const allowedFormats = new Set(['text', 'list', 'block', 'json']);
  if (!allowedFormats.has(outputFormat)) {
    throw new Error('format must be one of: text, list, json');
  }

  const workDayStart = parseTimeOfDay(config.workDayStart || '08:00');
  const workDayEnd = parseTimeOfDay(config.workDayEnd || '16:00');
  const excludeTime = normalizeExcludeTime(config.excludeTime);
  const excludeTimeWeekly = normalizeWeeklyExcludeTime(config.excludeTimeWeekly);
  const ignoreSummaries = normalizeIgnoreSummaries(config.ignoreSummaries);
  const timeGridMinutes = normalizeTimeGridMinutes(config.timeGridMinutes);

  const rangeStart = makeDateInTimeZone(
    startDate.getFullYear(),
    startDate.getMonth() + 1,
    startDate.getDate(),
    0,
    0,
    0,
    timeZone,
  );
  const rangeEnd = makeDateInTimeZone(
    endDate.getFullYear(),
    endDate.getMonth() + 1,
    endDate.getDate(),
    23,
    59,
    59,
    timeZone,
  );

  const calendarData = await loadCalendar(config);
  const instances = collectEventInstances(
    calendarData,
    rangeStart,
    rangeEnd,
    ignoreSummaries,
  );
  const minSlotMs = meetingLengthMinutes * 60 * 1000;
  const debugDate =
    args.debug && args.debug !== true ? parseDateInput(args.debug) : null;

  const lines = [];
  const results = [];
  for (const day of iterateDays(startDate, endDate)) {
    const dayStart = withTimeInZone(day, workDayStart, timeZone);
    const dayEnd = withTimeInZone(day, workDayEnd, timeZone);
    if (dayEnd <= dayStart) {
      throw new Error('workDayEnd must be after workDayStart');
    }

    const weekdayKey = getWeekdayKey(dayStart, timeZone);
    const weeklyWindows = excludeTimeWeekly[weekdayKey] || [];
    const excludeIntervals = [
      ...excludeTime.map((window) => ({
        start: withTimeInZone(day, window.start, timeZone),
        end: withTimeInZone(day, window.end, timeZone),
        label: 'excludeTime',
      })),
      ...weeklyWindows.map((window) => ({
        start: withTimeInZone(day, window.start, timeZone),
        end: withTimeInZone(day, window.end, timeZone),
        label: `excludeTimeWeekly:${weekdayKey}`,
      })),
    ];
    const freeSlots = getFreeSlotsForDay(
      instances,
      dayStart,
      dayEnd,
      minSlotMs,
      excludeIntervals,
    );
    if (args.debug && (!debugDate || isSameDay(day, debugDate))) {
      const busyIntervals = collectBusyIntervalsForDay(
        instances,
        dayStart,
        dayEnd,
        excludeIntervals,
      ).sort((a, b) => a.start.getTime() - b.start.getTime());
      const dayLabel = formatDateLabelInZone(dayStart, timeZone);
      console.log(`${dayLabel} busy:`);
      if (busyIntervals.length === 0) {
        console.log('  (none)');
      } else {
        for (const interval of busyIntervals) {
          console.log(
            `  ${formatTimeInZone(interval.start, timeZone)}-${formatTimeInZone(
              interval.end,
              timeZone,
            )}${interval.label ? ` ${interval.label}` : ''}`,
          );
        }
      }
    }
    const alignedSlots = alignFreeSlotsToGrid(
      freeSlots,
      minSlotMs,
      timeGridMinutes,
      timeZone,
    );
    const label = formatDateLabelInZone(dayStart, timeZone);
    const dateIso = formatDateIsoInZone(dayStart, timeZone);
    const slots = alignedSlots.map((slot) => ({
      start: formatTimeInZone(slot.start, timeZone),
      end: formatTimeInZone(slot.end, timeZone),
    }));
    results.push({ date: dateIso, label, slots });
    if (outputFormat === 'text') {
      if (slots.length === 0) {
        lines.push(`${label}: (no availability)`);
        continue;
      }
      const formatted = slots.map((slot) => `${slot.start}-${slot.end}`).join(' & ');
      lines.push(`${label}: ${formatted}`);
    } else if (outputFormat === 'list') {
      if (slots.length === 0) {
        lines.push(`${label}: (no availability)`);
        continue;
      }
      for (const slot of slots) {
        lines.push(`${label} ${slot.start}-${slot.end}`);
      }
    } else if (outputFormat === 'block') {
      lines.push(`${label}:`);
      if (slots.length === 0) {
        lines.push('  (no availability)');
      } else {
        for (const slot of slots) {
          lines.push(`  ${slot.start}-${slot.end}`);
        }
      }
    }
  }

  if (outputFormat === 'json') {
    console.log(JSON.stringify(results, null, 2));
  } else {
    console.log(lines.join('\n'));
  }
}

main().catch((error) => {
  console.error(`Error: ${error.message}`);
  process.exit(1);
});
