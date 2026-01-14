# Outlook Free Time

Small Node.js CLI that pulls an Outlook `.ics` calendar and prints free slots
between two dates.

## License

MIT. See `LICENSE`.

## Setup

1. Configure `config.json`:
   - Copy `config.example.json` to `config.json`
   - `icsUrl` points to your Outlook calendar `.ics` URL
   - `workDayStart` and `workDayEnd` set the daily availability window
   - `timeZone` sets the display and workday timezone (IANA name)
   - `timeGridMinutes` controls the allowed meeting start grid (default: 30)
   - `icsCacheFile` and `cacheMaxAgeMinutes` control local caching of the feed
   - `excludeTimeWeekly` blocks weekly windows keyed by weekday (MON..SUN)
   - `excludeTime` blocks daily windows like lunch
   - `ignoreSummaries` blocks events whose summary matches (case-insensitive)
2. Install dependencies:
   - `npm install`

## Get Outlook .ics URL (Web)

1. Open Outlook on the web.
2. Go to Settings (gear icon) → View all Outlook settings.
3. Open Calendar → Shared calendars.
4. Under Publish a calendar:
   - Choose your calendar.
   - Set permission to "Can view when I'm busy" (or "Can view all details" if
     you want event titles for debugging/filters).
   - Click Publish.
5. Copy the ICS link and paste it into `config.json` as `icsUrl`.

## Install As CLI

From this folder, run:

```bash
npm install
npm link
```

Then you can run:

```bash
outlook-free-time --length 30 --start 14.1 --end 16.1 --format block
```

Config lookup order (when `--config` is not provided):
1. `OUTLOOK_FREE_TIME_CONFIG`
2. `config.json` in the current directory
3. `~/.outlook-free-time.json`

To remove the global link:

```bash
npm unlink
```

## Usage

```bash
node src/index.js --length 30 --start 14.1 --end 16.1
```

Supported date formats:
- `DD.M` (uses current year)
- `DD.MM.YYYY`
- `YYYY-MM-DD`
 
Output formats:
- `--format text` (default): one line per day with `&`-separated slots
- `--format list`: one slot per line
- `--format block`: day header followed by indented slots
- `--format json`: structured output with `date`, `label`, and `slots`

Example output:

```
14.1: 08:00-09:30 & 13:00-16:00
15.1: 09:30-10:00 & 12:00-13:00
16.1: 14:00-16:00
```

## Notes

- Times are shown in your local system timezone.
- If you set `timeZone`, times and day boundaries use that timezone instead.
- To use a local `.ics` file, add `icsFile` to `config.json` instead of `icsUrl`.
- When using `icsUrl`, the feed is downloaded to `icsCacheFile` and refreshed
  when older than `cacheMaxAgeMinutes`.
- `excludeTime` accepts an object or array of objects with `start` and `end`
  in `HH:MM` 24-hour format.
- `excludeTimeWeekly` accepts an object keyed by weekday with values in the same
  shape as `excludeTime`. Weekday keys can be `MO`/`MON` through `SU`/`SUN`.
- `ignoreSummaries` is a list of summary strings to ignore. `Vapaa` is ignored
  by default.
- Slots are aligned to the configured `timeGridMinutes` for meeting starts.
- Use `--debug` (optionally with a date) to list busy intervals per day.
