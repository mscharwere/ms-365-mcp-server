/**
 * Offset guard for date comparisons inside a mail `$filter`.
 *
 * Message datetime properties (receivedDateTime, sentDateTime, createdDateTime,
 * lastModifiedDateTime) are Edm.DateTimeOffset values that Graph stores and compares in UTC.
 * Graph's own $filter docs show a date-only literal as valid
 * (`receivedDateTime ge 2017-04-01 and receivedDateTime lt 2017-05-01`), and nothing in a mail
 * request carries the caller's time zone, so a literal without an offset can only be resolved
 * against a zone the caller didn't choose (UTC). "Today's mail" in Pacific time written as
 * `receivedDateTime ge 2026-09-22` therefore starts at 5pm PDT the previous day — the same
 * silent window shift fixed for calendarView in calendar-datetime.ts.
 *
 * Unlike calendarView, the mail tools have no `timezone` parameter (their endpoints don't honor
 * `Prefer: outlook.timezone` for filtering), so the only zone that can be used to auto-correct is
 * the optional server default (MS365_MCP_DEFAULT_TIMEZONE, see default-timezone.ts). Resolution
 * order for each date literal in a comparison:
 *   1. literal already carries `Z`/`±HH:MM` → passed through untouched (the request builder
 *      percent-encodes `+`, so `+05:30` survives)
 *   2. offset-less, server default configured → that zone's DST-aware offset for that date is
 *      appended (a date-only literal becomes local midnight, `YYYY-MM-DDT00:00:00±HH:MM`)
 *   3. offset-less, no default → rejected before Graph is called, with an example of the
 *      explicit form (fail loud rather than silently shift the window)
 *
 * Scope is deliberately narrow: only `<property> <op> <literal>` comparisons on the four message
 * DateTimeOffset properties are inspected. Anything else in the filter is left alone.
 */

import { hasUtcOffset, isValidTimeZone, normalizeCalendarDateTime } from './calendar-datetime.js';
import { invalidDefaultTimeZoneMessage } from './default-timezone.js';

/** Mail list tools whose `$filter` is checked. Deliberately explicit. */
export const MAIL_DATETIME_FILTER_TOOLS: ReadonlySet<string> = new Set([
  'list-mail-messages', // /me/messages
  'list-mail-folder-messages', // /me/mailFolders/{id}/messages
  'list-mail-folder-messages-delta', // /me/mailFolders/{id}/messages/delta() (receivedDateTime ge/gt only)
  'list-shared-mailbox-messages', // /users/{id}/messages
  'list-shared-mailbox-folder-messages', // /users/{id}/mailFolders/{id}/messages
]);

export class MailFilterDateTimeError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'MailFilterDateTimeError';
  }
}

// <property> <op> <literal>, literal optionally single-quoted. Separators may be spaces or a
// literal '+' (URL-style, as Graph's docs sometimes write it). The literal alternation stops at
// the first character that can't be part of an ISO date/time/offset, so `and`/`)` are never eaten.
const COMPARISON_RE =
  /\b(receivedDateTime|sentDateTime|createdDateTime|lastModifiedDateTime)[\s+]+(eq|ne|gt|ge|lt|le)[\s+]+'?(\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?(?:Z|[+-]\d{2}:?\d{2})?)(?![\dT:.])/gi;

function offsetMissingError(
  property: string,
  op: string,
  literal: string
): MailFilterDateTimeError {
  const isDateOnly = !literal.includes('T');
  const example = isDateOnly ? `${literal}T00:00:00` : literal;
  return new MailFilterDateTimeError(
    `$filter compares ${property} to "${literal}", which has no UTC offset. Graph stores and ` +
      `compares ${property} in UTC, and mail tools have no timezone parameter, so an ` +
      'offset-less value silently shifts the window by your UTC offset. Include an explicit ' +
      `offset, e.g. "${property} ${op} ${example}-07:00" (Pacific daylight time) or ` +
      `"${property} ${op} ${example}Z" (UTC).`
  );
}

/**
 * Return `filter` with every offset-less message-datetime literal resolved in `defaultTimeZone`.
 *
 * - No offending literal → `filter` returned unchanged (same string).
 * - Offending literal, `defaultTimeZone` set → offset appended in place (quotes and everything else
 *   in the filter are preserved).
 * - Offending literal, no `defaultTimeZone` → MailFilterDateTimeError (unchanged pre-default text).
 * - Offending literal, `defaultTimeZone` not a recognized zone → MailFilterDateTimeError naming
 *   the misconfigured env var. The default is never validated unless it is actually consulted.
 */
export function normalizeMailFilterDateTimes(filter: string, defaultTimeZone?: string): string {
  const zone = defaultTimeZone?.trim() || undefined;
  return filter.replace(
    COMPARISON_RE,
    (match: string, property: string, op: string, literal: string) => {
      if (hasUtcOffset(literal)) return match;
      if (zone === undefined) throw offsetMissingError(property, op, literal);
      if (!isValidTimeZone(zone)) {
        throw new MailFilterDateTimeError(invalidDefaultTimeZoneMessage(zone));
      }
      const resolved = normalizeCalendarDateTime(property, literal, zone);
      // The literal is always the tail of the match (the trailing lookahead is zero-width).
      return match.slice(0, match.length - literal.length) + resolved;
    }
  );
}

/**
 * Throws MailFilterDateTimeError if `filter` compares a message datetime property against a
 * literal with no UTC offset. Returns normally (and never modifies the filter) otherwise.
 * Equivalent to normalizeMailFilterDateTimes with no default timezone configured.
 */
export function assertMailFilterDateTimesHaveOffset(filter: string): void {
  normalizeMailFilterDateTimes(filter, undefined);
}

/** Pull the $filter string out of tool params, whichever key the client used. */
export function getFilterParam(params: Record<string, unknown>): string | undefined {
  for (const key of ['filter', '$filter']) {
    const v = params[key];
    if (typeof v === 'string' && v.trim() !== '') return v;
  }
  return undefined;
}

/**
 * Return a copy of `params` with its $filter (whichever key the client used) normalized per
 * normalizeMailFilterDateTimes. `params` itself is not mutated. Throws MailFilterDateTimeError.
 */
export function normalizeMailFilterParams(
  params: Record<string, unknown>,
  defaultTimeZone?: string
): Record<string, unknown> {
  const out: Record<string, unknown> = { ...params };
  for (const key of ['filter', '$filter']) {
    const v = params[key];
    if (typeof v === 'string' && v.trim() !== '') {
      out[key] = normalizeMailFilterDateTimes(v, defaultTimeZone);
    }
  }
  return out;
}
