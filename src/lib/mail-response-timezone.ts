/**
 * Local-time enrichment for mail message datetimes.
 *
 * Microsoft Graph always returns message DateTimeOffset properties (receivedDateTime,
 * sentDateTime, createdDateTime, lastModifiedDateTime) in UTC. Unlike calendar endpoints, mail
 * endpoints do not honor `Prefer: outlook.timezone` — there is no way to ask Graph to render a
 * mail timestamp in local time (see calendar-datetime.ts for the calendar-side equivalent).
 * Printing the raw UTC value as if it were local time silently mislabels it: the incident that
 * prompted this module was FRIDAY printing mail `receivedDateTime` "2026-09-24T00:41:48Z" as
 * "12:41 AM" when the message actually arrived at 5:41 PM PST on 9/23.
 *
 * This module adds a `<field>Local` sibling next to each present UTC field, converted to
 * MS365_MCP_DEFAULT_TIMEZONE (see default-timezone.ts) and rendered as an ISO 8601 string with an
 * explicit UTC offset (e.g. "2026-09-23T17:41:48-07:00"). It NEVER overwrites the original UTC
 * field: mail-datetime-filter.ts's $filter normalization, and any downstream sort/compare logic,
 * depend on that field staying UTC.
 *
 * Opt-in only, same as the calendar/mail-filter fixes: with no default timezone configured, this
 * is a no-op and responses pass through completely unchanged.
 */

import { offsetMinutesAt, formatOffset } from './calendar-datetime.js';

/** Message DateTimeOffset properties Graph always returns in UTC. */
export const MAIL_DATETIME_FIELDS = [
  'receivedDateTime',
  'sentDateTime',
  'lastModifiedDateTime',
  'createdDateTime',
] as const;

export type MailDateTimeField = (typeof MAIL_DATETIME_FIELDS)[number];

/**
 * Mail tools whose responses get `<field>Local` siblings added. Deliberately explicit, and
 * matching the set mail-datetime-filter.ts already guards on the request side (the $filter tools)
 * plus get-mail-message, which returns a single message rather than a `{ value: [...] }` list.
 */
export const MAIL_LOCAL_FIELD_TOOLS: ReadonlySet<string> = new Set([
  'list-mail-messages',
  'list-mail-folder-messages',
  'list-mail-folder-messages-delta',
  'list-shared-mailbox-messages',
  'list-shared-mailbox-folder-messages',
  'get-mail-message',
]);

/**
 * Convert a UTC datetime string (Graph's DateTimeOffset wire format, e.g.
 * "2026-09-24T00:41:48Z" or with fractional seconds) to an ISO 8601 string in `timeZone`, with an
 * explicit `±HH:MM` offset appended (e.g. "2026-09-23T17:41:48-07:00").
 *
 * DST-correct: the offset is computed for the actual instant via Intl (offsetMinutesAt), not a
 * fixed table, so the same field converts correctly whether the message landed in daylight or
 * standard time.
 *
 * Returns null if `value` doesn't parse as a datetime or `timeZone` isn't a zone the runtime's
 * ICU recognizes — defensive, so an unexpected Graph payload shape or misconfigured env var never
 * throws mid-response; the caller simply omits that `*Local` field.
 */
export function convertUtcToLocalIso(value: string, timeZone: string): string | null {
  const epochMs = Date.parse(value);
  if (Number.isNaN(epochMs)) return null;

  let offsetMinutes: number;
  try {
    offsetMinutes = offsetMinutesAt(timeZone, epochMs);
  } catch {
    return null;
  }

  // Shifting the instant by the offset and reading it back with the UTC getters yields the
  // local wall-clock time (the same trick offsetMinutesForWallTime uses in reverse).
  const shifted = new Date(epochMs + offsetMinutes * 60000);
  const pad = (n: number): string => String(n).padStart(2, '0');

  const datePart = `${shifted.getUTCFullYear()}-${pad(shifted.getUTCMonth() + 1)}-${pad(shifted.getUTCDate())}`;
  const timePart = `${pad(shifted.getUTCHours())}:${pad(shifted.getUTCMinutes())}:${pad(shifted.getUTCSeconds())}`;

  return `${datePart}T${timePart}${formatOffset(offsetMinutes)}`;
}

/**
 * Return a copy of a single mail message object with a `<field>Local` sibling added for each
 * MAIL_DATETIME_FIELDS entry actually present as a string (so `$select` subsets are honored —
 * a field never requested by the caller was never fetched, so no Local sibling is invented for
 * it). The original message object is not mutated.
 */
function addLocalFieldsToMessage(
  msg: Record<string, unknown>,
  timeZone: string
): Record<string, unknown> {
  const out: Record<string, unknown> = { ...msg };
  for (const field of MAIL_DATETIME_FIELDS) {
    const value = msg[field];
    if (typeof value !== 'string' || value === '') continue;
    const local = convertUtcToLocalIso(value, timeZone);
    if (local !== null) {
      out[`${field}Local`] = local;
    }
  }
  return out;
}

/**
 * Enrich a parsed mail-tool JSON response with `<field>Local` siblings, in `timeZone`, for every
 * present UTC datetime field. Handles both response shapes:
 *   - list tools: `{ value: Message[], @odata.nextLink?, ... }` — every item in `value` is
 *     projected; non-object items are passed through unchanged.
 *   - get-mail-message: a bare Message object.
 *
 * Returns `raw` unchanged (same reference) if it isn't an object, or — for the bare-object case —
 * if it doesn't look like a message (none of the four tracked fields present), so an unrelated
 * single-object Graph response is never accidentally rewritten. Never mutates `raw`.
 */
export function addMailResponseLocalFields(raw: unknown, timeZone: string): unknown {
  if (!raw || typeof raw !== 'object') return raw;
  const obj = raw as Record<string, unknown>;

  if (Array.isArray(obj.value)) {
    return {
      ...obj,
      value: obj.value.map((item) =>
        item && typeof item === 'object'
          ? addLocalFieldsToMessage(item as Record<string, unknown>, timeZone)
          : item
      ),
    };
  }

  const looksLikeMessage = MAIL_DATETIME_FIELDS.some((field) => typeof obj[field] === 'string');
  return looksLikeMessage ? addLocalFieldsToMessage(obj, timeZone) : raw;
}
