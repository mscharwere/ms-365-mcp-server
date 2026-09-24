/**
 * UTC-offset normalization for calendarView query windows.
 *
 * Microsoft Graph interprets a `startDateTime` / `endDateTime` query value that carries no
 * UTC offset (e.g. "2026-09-22T00:00:00") as UTC. The separate `timezone` tool parameter only
 * sets the `Prefer: outlook.timezone` header, which changes how RETURNED event times are
 * rendered — it does NOT reinterpret the input window. A caller who passes a bare datetime
 * intending local time therefore silently gets a window shifted by the zone's offset, and
 * evening events drop out of the result with no error.
 *
 * This module closes that gap for the calendarView tools. Resolution order for each value:
 *   1. datetime already carrying `Z` or `±HH:MM` → passed through byte-for-byte unchanged
 *   2. bare datetime + `timezone` param → inject that zone's real (DST-aware) offset for that date
 *   3. bare datetime + no `timezone`, server default configured (MS365_MCP_DEFAULT_TIMEZONE, see
 *      default-timezone.ts) → same as 2, using the default zone
 *   4. bare datetime + neither → validation error (never send an ambiguous window)
 * The default also drives the `Prefer: outlook.timezone` display header when no explicit `timezone`
 * param is passed (set in graph-tools.ts), so returned event times come back in the default zone;
 * an explicit `timezone` param always wins for both the query window and the display header.
 *
 * Offsets are computed with the built-in `Intl` API (full ICU ships with Node 18+), so DST
 * transitions are handled from the tz database rather than a fixed offset table.
 */

import { invalidDefaultTimeZoneMessage } from './default-timezone.js';

/**
 * Tools whose startDateTime/endDateTime window is normalized. Deliberately explicit: every entry
 * is a Graph endpoint whose required startDateTime/endDateTime query params are documented as
 * "interpreted as UTC if no offset" and which accepts `Prefer: outlook.timezone` (so it exposes the
 * `timezone` tool param via `supportsTimezone` in endpoints.json).
 */
export const CALENDAR_VIEW_OFFSET_TOOLS: ReadonlySet<string> = new Set([
  'get-calendar-view', // /me/calendarView
  'get-specific-calendar-view', // /me/calendars/{id}/calendarView
  'get-shared-calendar-view', // /users/{id}/calendarView
  'get-group-calendar-view', // /groups/{id}/calendarView
  'list-calendar-event-instances', // /me/calendars/{id}/events/{id}/instances
  'list-calendar-view-delta', // /me/calendarView/delta()
]);

export const CALENDAR_WINDOW_PARAMS = ['startDateTime', 'endDateTime'] as const;

export class CalendarDateTimeError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'CalendarDateTimeError';
  }
}

// YYYY-MM-DD, optionally followed by THH:MM[:SS[.fraction]] and an optional offset designator.
const ISO_RE =
  /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(\.\d+)?)?)?(Z|[+-]\d{2}:?\d{2})?$/i;

interface ParsedDateTime {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
  hasTime: boolean;
  hasOffset: boolean;
}

function parseIso(value: string): ParsedDateTime | null {
  const m = ISO_RE.exec(value);
  if (!m) return null;
  return {
    year: Number(m[1]),
    month: Number(m[2]),
    day: Number(m[3]),
    hour: m[4] !== undefined ? Number(m[4]) : 0,
    minute: m[5] !== undefined ? Number(m[5]) : 0,
    second: m[6] !== undefined ? Number(m[6]) : 0,
    hasTime: m[4] !== undefined,
    hasOffset: m[8] !== undefined,
  };
}

/** True when the value ends in `Z` or a `±HH:MM` / `±HHMM` offset. */
export function hasUtcOffset(value: string): boolean {
  const parsed = parseIso(value.trim());
  return parsed !== null && parsed.hasOffset;
}

/** True when `timeZone` is a zone name the runtime's ICU knows. */
export function isValidTimeZone(timeZone: string): boolean {
  try {
    new Intl.DateTimeFormat('en-US', { timeZone });
    return true;
  } catch {
    return false;
  }
}

/** Throws CalendarDateTimeError if `timeZone` is not a zone name the runtime's ICU knows. */
function assertValidTimeZone(timeZone: string): void {
  if (!isValidTimeZone(timeZone)) {
    throw new CalendarDateTimeError(
      `Unrecognized timezone "${timeZone}". Pass an IANA timezone name (e.g. "America/Los_Angeles") ` +
        'so the offset can be inferred, or include an explicit UTC offset in startDateTime/endDateTime ' +
        '(e.g. "2026-09-22T00:00:00-07:00").'
    );
  }
}

/**
 * Offset of `timeZone` from UTC at the given instant, in minutes (e.g. -420 for PDT).
 *
 * Exported for reuse by mail-response-timezone.ts, which converts an already-known UTC instant
 * (Graph's mail DateTimeOffset fields) to local time — unlike offsetMinutesForWallTime below,
 * there's no wall-clock ambiguity to resolve since the instant is unambiguous.
 */
export function offsetMinutesAt(timeZone: string, epochMs: number): number {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone,
    hourCycle: 'h23',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  }).formatToParts(new Date(epochMs));
  const get = (type: string): number => Number(parts.find((p) => p.type === type)?.value);
  const wallAsUtc = Date.UTC(
    get('year'),
    get('month') - 1,
    get('day'),
    get('hour'),
    get('minute'),
    get('second')
  );
  // Drop sub-second precision from the instant so the difference is a whole number of minutes.
  return Math.round((wallAsUtc - Math.floor(epochMs / 1000) * 1000) / 60000);
}

/**
 * Offset (minutes) that `timeZone` applies to the given wall-clock time.
 *
 * Two-pass resolution handles DST: the first guess uses the offset at the naive-UTC instant,
 * the second re-checks at the corrected instant. For an ambiguous fall-back hour the earlier
 * (pre-transition) offset is kept; for a nonexistent spring-forward time one of the two offsets
 * in effect that day is used. Either way the result is a real offset of that zone on that date
 * (the only hours affected are the ~1h transition window itself, never a query at midnight).
 */
function offsetMinutesForWallTime(timeZone: string, p: ParsedDateTime): number {
  const naive = Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second);
  const first = offsetMinutesAt(timeZone, naive);
  const second = offsetMinutesAt(timeZone, naive - first * 60000);
  if (first === second) return first;
  const third = offsetMinutesAt(timeZone, naive - second * 60000);
  return third === second ? second : first;
}

/** Format a UTC-offset in minutes as `±HH:MM` (e.g. -420 -> "-07:00"). Exported for reuse. */
export function formatOffset(minutes: number): string {
  const sign = minutes < 0 ? '-' : '+';
  const abs = Math.abs(minutes);
  const hh = String(Math.floor(abs / 60)).padStart(2, '0');
  const mm = String(abs % 60).padStart(2, '0');
  return `${sign}${hh}:${mm}`;
}

/**
 * Normalize a single calendarView window value.
 *
 * - Already has an offset → returned unchanged.
 * - Bare, timezone given → offset for that zone/date appended (date-only gets `T00:00:00`).
 * - Bare, no timezone → CalendarDateTimeError.
 * - Not an ISO 8601 date/datetime this module recognizes → returned unchanged; Graph's own
 *   validation rejects malformed values with a 400, so nothing is silently re-interpreted here.
 */
export function normalizeCalendarDateTime(
  paramName: string,
  value: string,
  timeZone: string | undefined
): string {
  const trimmed = value.trim();
  const parsed = parseIso(trimmed);
  if (parsed === null || parsed.hasOffset) return value;

  if (!timeZone) {
    throw new CalendarDateTimeError(
      `${paramName} "${value}" has no UTC offset. Microsoft Graph treats an offset-less ` +
        'startDateTime/endDateTime as UTC, which silently shifts the query window. ' +
        'startDateTime/endDateTime must include a UTC offset, or pass `timezone` so it can be ' +
        'inferred (e.g. "2026-09-22T00:00:00-07:00", or timezone "America/Los_Angeles").'
    );
  }

  assertValidTimeZone(timeZone);
  const offset = formatOffset(offsetMinutesForWallTime(timeZone, parsed));
  const base = parsed.hasTime ? trimmed : `${trimmed}T00:00:00`;
  return `${base}${offset}`;
}

/**
 * Return a copy of `params` with startDateTime/endDateTime normalized. Only string values are
 * touched; `params` itself is not mutated. Throws CalendarDateTimeError on an unfixable window.
 *
 * `defaultTimeZone` is the server-level fallback (the caller reads MS365_MCP_DEFAULT_TIMEZONE); it is
 * used only when a value is offset-less AND no non-blank `timezone` param was passed. An explicit
 * `timezone` param always wins. Omitted/undefined means no default (bare values are rejected).
 */
export function normalizeCalendarWindowParams(
  params: Record<string, unknown>,
  defaultTimeZone?: string
): Record<string, unknown> {
  const explicitTimeZone =
    typeof params.timezone === 'string' && params.timezone.trim() !== ''
      ? params.timezone.trim()
      : undefined;
  const fallbackTimeZone = defaultTimeZone?.trim() || undefined;
  const timeZone = explicitTimeZone ?? fallbackTimeZone;
  const usingDefault = explicitTimeZone === undefined && fallbackTimeZone !== undefined;
  const out: Record<string, unknown> = { ...params };
  for (const name of CALENDAR_WINDOW_PARAMS) {
    const v = params[name];
    if (typeof v === 'string' && v !== '') {
      // A misconfigured default is only an error when it would actually be consulted.
      const parsed = parseIso(v.trim());
      const needsZone = parsed !== null && !parsed.hasOffset;
      if (needsZone && usingDefault && !isValidTimeZone(fallbackTimeZone!)) {
        throw new CalendarDateTimeError(invalidDefaultTimeZoneMessage(fallbackTimeZone!));
      }
      out[name] = normalizeCalendarDateTime(name, v, timeZone);
    }
  }
  return out;
}
