import { describe, expect, it } from 'vitest';
import {
  CALENDAR_VIEW_OFFSET_TOOLS,
  CalendarDateTimeError,
  hasUtcOffset,
  normalizeCalendarDateTime,
  normalizeCalendarWindowParams,
} from '../src/lib/calendar-datetime.js';

const LA = 'America/Los_Angeles';

describe('hasUtcOffset', () => {
  it.each([
    ['2026-09-22T00:00:00Z', true],
    ['2026-09-22T00:00:00z', true],
    ['2026-09-22T00:00:00-07:00', true],
    ['2026-09-22T00:00:00+0530', true],
    ['2026-09-22T00:00:00.0000000+00:00', true],
    ['2026-09-22T00:00:00', false],
    ['2026-09-22T00:00', false],
    ['2026-09-22T00:00:00.0000000', false],
    ['2026-09-22', false],
  ])('%s -> %s', (value, expected) => {
    expect(hasUtcOffset(value)).toBe(expected);
  });
});

describe('normalizeCalendarDateTime — bare datetime + timezone', () => {
  it('injects PDT (-07:00) on a DST-active date', () => {
    expect(normalizeCalendarDateTime('startDateTime', '2026-09-22T00:00:00', LA)).toBe(
      '2026-09-22T00:00:00-07:00'
    );
    expect(normalizeCalendarDateTime('endDateTime', '2026-09-22T23:59:59', LA)).toBe(
      '2026-09-22T23:59:59-07:00'
    );
  });

  it('injects PST (-08:00) on a DST-inactive date', () => {
    expect(normalizeCalendarDateTime('startDateTime', '2026-01-15T00:00:00', LA)).toBe(
      '2026-01-15T00:00:00-08:00'
    );
    expect(normalizeCalendarDateTime('endDateTime', '2026-12-01T18:30:00', LA)).toBe(
      '2026-12-01T18:30:00-08:00'
    );
  });

  it('resolves the correct side of both 2026 US DST transitions', () => {
    // Spring forward 2026-03-08 02:00 PST -> 03:00 PDT
    expect(normalizeCalendarDateTime('s', '2026-03-08T00:00:00', LA)).toBe(
      '2026-03-08T00:00:00-08:00'
    );
    expect(normalizeCalendarDateTime('s', '2026-03-08T12:00:00', LA)).toBe(
      '2026-03-08T12:00:00-07:00'
    );
    // Fall back 2026-11-01 02:00 PDT -> 01:00 PST
    expect(normalizeCalendarDateTime('s', '2026-11-01T00:00:00', LA)).toBe(
      '2026-11-01T00:00:00-07:00'
    );
    expect(normalizeCalendarDateTime('s', '2026-11-01T12:00:00', LA)).toBe(
      '2026-11-01T12:00:00-08:00'
    );
  });

  it('keeps the earlier offset for the ambiguous fall-back hour', () => {
    expect(normalizeCalendarDateTime('s', '2026-11-01T01:30:00', LA)).toBe(
      '2026-11-01T01:30:00-07:00'
    );
  });

  it('handles other zones, including southern-hemisphere DST and non-hour offsets', () => {
    expect(normalizeCalendarDateTime('s', '2026-01-15T00:00:00', 'Australia/Sydney')).toBe(
      '2026-01-15T00:00:00+11:00'
    );
    expect(normalizeCalendarDateTime('s', '2026-07-15T00:00:00', 'Australia/Sydney')).toBe(
      '2026-07-15T00:00:00+10:00'
    );
    expect(normalizeCalendarDateTime('s', '2026-07-15T00:00:00', 'Asia/Kolkata')).toBe(
      '2026-07-15T00:00:00+05:30'
    );
    expect(normalizeCalendarDateTime('s', '2026-07-15T00:00:00', 'UTC')).toBe(
      '2026-07-15T00:00:00+00:00'
    );
  });

  it('preserves fractional seconds and minute-only precision', () => {
    expect(normalizeCalendarDateTime('s', '2026-09-22T00:00:00.0000000', LA)).toBe(
      '2026-09-22T00:00:00.0000000-07:00'
    );
    expect(normalizeCalendarDateTime('s', '2026-09-22T18:30', LA)).toBe('2026-09-22T18:30-07:00');
  });

  it('expands a bare date to local midnight with offset', () => {
    expect(normalizeCalendarDateTime('s', '2026-09-22', LA)).toBe('2026-09-22T00:00:00-07:00');
  });

  it('throws a clear error for an unrecognized timezone name', () => {
    expect(() => normalizeCalendarDateTime('s', '2026-09-22T00:00:00', 'Not/AZone')).toThrow(
      CalendarDateTimeError
    );
    expect(() => normalizeCalendarDateTime('s', '2026-09-22T00:00:00', 'Not/AZone')).toThrow(
      /Unrecognized timezone "Not\/AZone"/
    );
  });
});

describe('normalizeCalendarDateTime — bare datetime without timezone', () => {
  it('throws CalendarDateTimeError naming the param and the fix', () => {
    const call = () => normalizeCalendarDateTime('startDateTime', '2026-09-22T00:00:00', undefined);
    expect(call).toThrow(CalendarDateTimeError);
    expect(call).toThrow(/startDateTime "2026-09-22T00:00:00" has no UTC offset/);
    expect(call).toThrow(/must include a UTC offset/);
    expect(call).toThrow(/pass `timezone`/);
  });
});

describe('normalizeCalendarDateTime — already offset-qualified', () => {
  it.each([
    '2026-09-22T00:00:00Z',
    '2026-09-22T00:00:00-07:00',
    '2026-09-22T00:00:00-08:00', // deliberately "wrong" for Sept — caller's choice is respected
    '2026-09-22T00:00:00+0530',
    '2026-09-22T00:00:00.0000000Z',
  ])('%s is returned unchanged with or without timezone', (value) => {
    expect(normalizeCalendarDateTime('s', value, LA)).toBe(value);
    expect(normalizeCalendarDateTime('s', value, undefined)).toBe(value);
    expect(normalizeCalendarDateTime('s', value, 'Not/AZone')).toBe(value);
  });

  it('leaves unrecognized formats for Graph to validate', () => {
    expect(normalizeCalendarDateTime('s', 'yesterday', undefined)).toBe('yesterday');
  });
});

describe('normalizeCalendarWindowParams', () => {
  it('normalizes both window params and does not mutate the input', () => {
    const input = {
      startDateTime: '2026-09-22T00:00:00',
      endDateTime: '2026-09-23T00:00:00',
      timezone: LA,
      top: 10,
    };
    const snapshot = { ...input };
    const out = normalizeCalendarWindowParams(input);
    expect(out).toEqual({
      startDateTime: '2026-09-22T00:00:00-07:00',
      endDateTime: '2026-09-23T00:00:00-07:00',
      timezone: LA,
      top: 10,
    });
    expect(input).toEqual(snapshot);
  });

  it('fixes a bare value even when the other bound already has an offset', () => {
    const out = normalizeCalendarWindowParams({
      startDateTime: '2026-09-22T00:00:00-07:00',
      endDateTime: '2026-09-23T00:00:00',
      timezone: LA,
    });
    expect(out.startDateTime).toBe('2026-09-22T00:00:00-07:00');
    expect(out.endDateTime).toBe('2026-09-23T00:00:00-07:00');
  });

  it('treats a blank timezone as absent', () => {
    expect(() =>
      normalizeCalendarWindowParams({
        startDateTime: '2026-09-22T00:00:00',
        endDateTime: '2026-09-23T00:00:00Z',
        timezone: '  ',
      })
    ).toThrow(CalendarDateTimeError);
  });

  it('is scoped to exactly the six startDateTime/endDateTime window tools', () => {
    expect([...CALENDAR_VIEW_OFFSET_TOOLS].sort()).toEqual([
      'get-calendar-view',
      'get-group-calendar-view',
      'get-shared-calendar-view',
      'get-specific-calendar-view',
      'list-calendar-event-instances',
      'list-calendar-view-delta',
    ]);
  });

  it('every scoped tool exposes a timezone param (supportsTimezone) in endpoints.json', async () => {
    const endpoints = (await import('../src/endpoints.json', { with: { type: 'json' } }))
      .default as Array<{ toolName: string; supportsTimezone?: boolean }>;
    for (const name of CALENDAR_VIEW_OFFSET_TOOLS) {
      const ep = endpoints.find((e) => e.toolName === name);
      expect(ep, name).toBeDefined();
      expect(ep!.supportsTimezone, name).toBe(true);
    }
  });
});
