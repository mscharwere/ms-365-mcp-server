import { describe, expect, it } from 'vitest';
import {
  MAIL_DATETIME_FIELDS,
  MAIL_LOCAL_FIELD_TOOLS,
  addMailResponseLocalFields,
  convertUtcToLocalIso,
} from '../src/lib/mail-response-timezone.js';

const LA = 'America/Los_Angeles';

describe('convertUtcToLocalIso', () => {
  it('converts a September (PDT, -07:00) UTC timestamp — the 2026-09-24 incident value', () => {
    // FRIDAY incident: "2026-09-24T00:41:48Z" printed as "12:41 AM" when it was 5:41 PM PST(sic,
    // actually PDT) the prior day. Confirm the correct conversion.
    expect(convertUtcToLocalIso('2026-09-24T00:41:48Z', LA)).toBe('2026-09-23T17:41:48-07:00');
  });

  it('converts a December (PST, -08:00) UTC timestamp', () => {
    expect(convertUtcToLocalIso('2026-12-15T04:30:00Z', LA)).toBe('2026-12-14T20:30:00-08:00');
  });

  it('is DST-correct across both 2026 US transitions', () => {
    // Spring forward: 2026-03-08 09:59:59Z is 01:59:59 PST; 10:00:00Z is 03:00:00 PDT.
    expect(convertUtcToLocalIso('2026-03-08T09:59:59Z', LA)).toBe('2026-03-08T01:59:59-08:00');
    expect(convertUtcToLocalIso('2026-03-08T10:00:00Z', LA)).toBe('2026-03-08T03:00:00-07:00');
    // Fall back: 2026-11-01 08:59:59Z is 01:59:59 PDT; 09:00:00Z is 01:00:00 PST.
    expect(convertUtcToLocalIso('2026-11-01T08:59:59Z', LA)).toBe('2026-11-01T01:59:59-07:00');
    expect(convertUtcToLocalIso('2026-11-01T09:00:00Z', LA)).toBe('2026-11-01T01:00:00-08:00');
  });

  it('handles fractional-second Graph timestamps', () => {
    expect(convertUtcToLocalIso('2026-09-24T00:41:48.1234567Z', LA)).toBe(
      '2026-09-23T17:41:48-07:00'
    );
  });

  it('works for a non-US zone (positive offset, no DST)', () => {
    expect(convertUtcToLocalIso('2026-09-24T00:41:48Z', 'Asia/Kolkata')).toBe(
      '2026-09-24T06:11:48+05:30'
    );
  });

  it('returns null for an unparseable value', () => {
    expect(convertUtcToLocalIso('not-a-date', LA)).toBeNull();
    expect(convertUtcToLocalIso('', LA)).toBeNull();
  });

  it('returns null for an unrecognized IANA zone rather than throwing', () => {
    expect(convertUtcToLocalIso('2026-09-24T00:41:48Z', 'Not/AZone')).toBeNull();
  });
});

describe('addMailResponseLocalFields', () => {
  const rawMessage = () => ({
    id: 'msg-1',
    subject: 'Test',
    receivedDateTime: '2026-09-24T00:41:48Z',
    sentDateTime: '2026-09-23T23:40:00Z',
    createdDateTime: '2026-09-23T23:40:01Z',
    lastModifiedDateTime: '2026-09-24T00:41:49Z',
  });

  it('adds a Local sibling for each of the four fields, in a list response', () => {
    const raw = { value: [rawMessage()], '@odata.nextLink': 'https://graph/next' };
    const out = addMailResponseLocalFields(raw, LA) as any;

    expect(out.value[0].receivedDateTimeLocal).toBe('2026-09-23T17:41:48-07:00');
    expect(out.value[0].sentDateTimeLocal).toBe('2026-09-23T16:40:00-07:00');
    expect(out.value[0].createdDateTimeLocal).toBe('2026-09-23T16:40:01-07:00');
    expect(out.value[0].lastModifiedDateTimeLocal).toBe('2026-09-23T17:41:49-07:00');
    // pagination token survives untouched
    expect(out['@odata.nextLink']).toBe('https://graph/next');
  });

  it('never overwrites the original UTC fields', () => {
    const raw = { value: [rawMessage()] };
    const out = addMailResponseLocalFields(raw, LA) as any;
    expect(out.value[0].receivedDateTime).toBe('2026-09-24T00:41:48Z');
    expect(out.value[0].sentDateTime).toBe('2026-09-23T23:40:00Z');
    expect(out.value[0].createdDateTime).toBe('2026-09-23T23:40:01Z');
    expect(out.value[0].lastModifiedDateTime).toBe('2026-09-24T00:41:49Z');
  });

  it('adds Local fields for a single get-mail-message-shaped object (no value wrapper)', () => {
    const out = addMailResponseLocalFields(rawMessage(), LA) as any;
    expect(out.receivedDateTimeLocal).toBe('2026-09-23T17:41:48-07:00');
    expect(out.id).toBe('msg-1');
  });

  it('honors $select — only adds Local for a field actually present', () => {
    const selected = { id: 'msg-1', receivedDateTime: '2026-09-24T00:41:48Z' };
    const out = addMailResponseLocalFields({ value: [selected] }, LA) as any;
    expect(out.value[0].receivedDateTimeLocal).toBe('2026-09-23T17:41:48-07:00');
    expect(out.value[0]).not.toHaveProperty('sentDateTimeLocal');
    expect(out.value[0]).not.toHaveProperty('createdDateTimeLocal');
    expect(out.value[0]).not.toHaveProperty('lastModifiedDateTimeLocal');
  });

  it('does not mutate the input', () => {
    const raw = { value: [rawMessage()] };
    const before = JSON.stringify(raw);
    addMailResponseLocalFields(raw, LA);
    expect(JSON.stringify(raw)).toBe(before);
  });

  it('leaves a non-message single object untouched (does not misclassify arbitrary payloads)', () => {
    const other = { id: 'not-a-message', foo: 'bar' };
    const out = addMailResponseLocalFields(other, LA);
    expect(out).toBe(other);
  });

  it('passes through non-object and empty-list responses unchanged', () => {
    expect(addMailResponseLocalFields(null, LA)).toBeNull();
    expect(addMailResponseLocalFields('str', LA)).toBe('str');
    const emptyList = { value: [] };
    expect(addMailResponseLocalFields(emptyList, LA)).toEqual({ value: [] });
  });

  it('covers all four MAIL_DATETIME_FIELDS constants', () => {
    expect([...MAIL_DATETIME_FIELDS].sort()).toEqual(
      ['createdDateTime', 'lastModifiedDateTime', 'receivedDateTime', 'sentDateTime'].sort()
    );
  });

  it('MAIL_LOCAL_FIELD_TOOLS includes the mail-listing tools plus get-mail-message', () => {
    for (const tool of [
      'list-mail-messages',
      'list-mail-folder-messages',
      'list-mail-folder-messages-delta',
      'list-shared-mailbox-messages',
      'list-shared-mailbox-folder-messages',
      'get-mail-message',
    ]) {
      expect(MAIL_LOCAL_FIELD_TOOLS.has(tool)).toBe(true);
    }
    expect(MAIL_LOCAL_FIELD_TOOLS.has('send-mail')).toBe(false);
  });
});
