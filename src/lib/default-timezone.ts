/**
 * Optional server-level default timezone for offset-less datetimes.
 *
 * When `MS365_MCP_DEFAULT_TIMEZONE` is set to an IANA zone name (e.g. "Europe/Berlin"), a bare
 * datetime (no `Z`/`±HH:MM`) that the caller did not pair with an explicit `timezone` param is
 * interpreted in this zone instead of being rejected:
 *   - calendarView window params (see calendar-datetime.ts)
 *   - mail `$filter` date literals (see mail-datetime-filter.ts)
 *
 * Opt-in only. Unset or blank means no default: offset-less values keep being rejected exactly as
 * before. There is deliberately no built-in fallback zone.
 *
 * Read on every call (not cached at import) to match how other MS365_MCP_* env vars are consumed.
 */

export const DEFAULT_TIMEZONE_ENV = 'MS365_MCP_DEFAULT_TIMEZONE';

/** The configured default zone name (trimmed), or undefined when unset/blank. Not validated. */
export function getDefaultTimeZone(): string | undefined {
  const raw = process.env.MS365_MCP_DEFAULT_TIMEZONE?.trim();
  return raw || undefined;
}

/** Error text for a configured default that the runtime's ICU does not recognize. */
export function invalidDefaultTimeZoneMessage(value: string): string {
  return (
    `${DEFAULT_TIMEZONE_ENV} is set to "${value}", which is not a recognized IANA timezone name ` +
    '(e.g. "Europe/Berlin"). Fix the server configuration, or include an explicit UTC offset in the ' +
    'datetime (e.g. "2026-09-22T00:00:00-07:00").'
  );
}
