import { beforeEach, describe, expect, it, vi } from 'vitest';
import { z } from 'zod';
import {
  MAIL_DATETIME_FILTER_TOOLS,
  MailFilterDateTimeError,
  assertMailFilterDateTimesHaveOffset,
  getFilterParam,
} from '../src/lib/mail-datetime-filter.js';
import { registerGraphTools } from '../src/graph-tools.js';
import type { GraphClient } from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

vi.mock('../src/generated/client.js', () => {
  const listParams = [
    { name: 'filter', type: 'Query', schema: z.string().optional() },
    { name: 'top', type: 'Query', schema: z.number().int().optional() },
    { name: 'select', type: 'Query', schema: z.array(z.string()).optional() },
  ];
  return {
    api: {
      endpoints: [
        {
          alias: 'list-mail-messages',
          method: 'get',
          path: '/me/messages',
          description: 'List messages.',
          parameters: listParams,
        },
        {
          alias: 'list-mail-folder-messages',
          method: 'get',
          path: '/me/mailFolders/:mailFolderId/messages',
          description: 'List folder messages.',
          parameters: [{ name: 'mailFolderId', type: 'Path', schema: z.string() }, ...listParams],
        },
        {
          alias: 'list-mail-folder-messages-delta',
          method: 'get',
          path: '/me/mailFolders/:mailFolderId/messages/delta()',
          description: 'Delta of folder messages.',
          parameters: [{ name: 'mailFolderId', type: 'Path', schema: z.string() }, ...listParams],
        },
        {
          alias: 'list-shared-mailbox-messages',
          method: 'get',
          path: '/users/:userId/messages',
          description: 'List shared mailbox messages.',
          parameters: [{ name: 'userId', type: 'Path', schema: z.string() }, ...listParams],
        },
        {
          alias: 'list-shared-mailbox-folder-messages',
          method: 'get',
          path: '/users/:userId/mailFolders/:mailFolderId/messages',
          description: 'List shared mailbox folder messages.',
          parameters: [
            { name: 'userId', type: 'Path', schema: z.string() },
            { name: 'mailFolderId', type: 'Path', schema: z.string() },
            ...listParams,
          ],
        },
        {
          alias: 'list-chat-messages',
          method: 'get',
          path: '/chats/:chatId/messages',
          description: 'Teams chat messages (not a mail tool; not guarded).',
          parameters: [{ name: 'chatId', type: 'Path', schema: z.string() }, ...listParams],
        },
      ],
    },
  };
});

// Mail tools with the path params each needs.
const MAIL_TOOLS: Array<[string, Record<string, string>]> = [
  ['list-mail-messages', {}],
  ['list-mail-folder-messages', { mailFolderId: 'inbox' }],
  ['list-mail-folder-messages-delta', { mailFolderId: 'inbox' }],
  ['list-shared-mailbox-messages', { userId: 'shared@example.com' }],
  ['list-shared-mailbox-folder-messages', { userId: 'shared@example.com', mailFolderId: 'inbox' }],
];

describe('assertMailFilterDateTimesHaveOffset', () => {
  it.each([
    'receivedDateTime ge 2026-09-22T00:00:00-07:00',
    'receivedDateTime ge 2026-09-22T07:00:00Z',
    'receivedDateTime ge 2026-09-22T07:00:00z',
    'receivedDateTime ge 2026-09-22T00:00:00.0000000-08:00',
    'receivedDateTime ge 2026-09-22T05:30:00+05:30',
    'receivedDateTime ge 2026-09-22T00:00:00-0700',
    'ReceivedDateTime ge 2026-09-22T00:00:00-07:00 and receivedDateTime lt 2026-09-23T00:00:00-07:00',
    "sentDateTime lt '2026-09-22T00:00:00Z'",
    'receivedDateTime+ge+2026-09-22T07:00:00Z',
    // No datetime comparison at all
    "from/emailAddress/address eq 'someone@example.com'",
    'isRead eq false',
    // Other date-ish content that isn't a message DateTimeOffset comparison is left alone
    "contains(subject, '2026-09-22')",
    "flag/flagStatus eq 'flagged'",
  ])('accepts %s', (filter) => {
    expect(() => assertMailFilterDateTimesHaveOffset(filter)).not.toThrow();
  });

  it.each([
    ['receivedDateTime ge 2026-09-22T00:00:00', 'receivedDateTime', '2026-09-22T00:00:00'],
    ['receivedDateTime ge 2026-09-22', 'receivedDateTime', '2026-09-22'],
    // The exact form from Graph's own $filter docs: accepted by Graph, silently evaluated in UTC
    [
      'ReceivedDateTime ge 2017-04-01 and receivedDateTime lt 2017-05-01',
      'ReceivedDateTime',
      '2017-04-01',
    ],
    ['sentDateTime lt 2026-09-22T18:30', 'sentDateTime', '2026-09-22T18:30'],
    [
      'createdDateTime gt 2026-09-22T00:00:00.0000000',
      'createdDateTime',
      '2026-09-22T00:00:00.0000000',
    ],
    ['lastModifiedDateTime le 2026-09-22', 'lastModifiedDateTime', '2026-09-22'],
    ["receivedDateTime ge '2026-09-22T00:00:00'", 'receivedDateTime', '2026-09-22T00:00:00'],
    // Second clause is the bad one
    [
      'receivedDateTime ge 2026-09-22T00:00:00-07:00 and receivedDateTime lt 2026-09-23',
      'receivedDateTime',
      '2026-09-23',
    ],
    ['isRead eq false and (receivedDateTime ge 2026-09-22)', 'receivedDateTime', '2026-09-22'],
  ])('rejects %s', (filter, property, literal) => {
    let caught: unknown;
    try {
      assertMailFilterDateTimesHaveOffset(filter);
    } catch (err) {
      caught = err;
    }
    expect(caught).toBeInstanceOf(MailFilterDateTimeError);
    const msg = (caught as Error).message;
    expect(msg).toContain(`compares ${property} to "${literal}"`);
    expect(msg).toContain('has no UTC offset');
    expect(msg).toMatch(/-07:00" \(Pacific daylight time\)/);
  });

  it('suggests a full datetime (not a date-only+offset) for date-only literals', () => {
    expect(() => assertMailFilterDateTimesHaveOffset('receivedDateTime ge 2026-09-22')).toThrow(
      /"receivedDateTime ge 2026-09-22T00:00:00-07:00"/
    );
  });
});

describe('getFilterParam', () => {
  it('reads filter or $filter and ignores blanks/non-strings', () => {
    expect(getFilterParam({ filter: 'isRead eq false' })).toBe('isRead eq false');
    expect(getFilterParam({ $filter: 'isRead eq false' })).toBe('isRead eq false');
    expect(getFilterParam({ filter: '  ' })).toBeUndefined();
    expect(getFilterParam({ filter: 3 })).toBeUndefined();
    expect(getFilterParam({})).toBeUndefined();
  });
});

describe('MAIL_DATETIME_FILTER_TOOLS', () => {
  it('is scoped to exactly the five mail list/delta tools', () => {
    expect([...MAIL_DATETIME_FILTER_TOOLS].sort()).toEqual(MAIL_TOOLS.map(([n]) => n).sort());
  });
});

describe('mail tool execution with a date $filter', () => {
  let mockServer: { tool: ReturnType<typeof vi.fn> };
  let mockGraphClient: GraphClient;

  beforeEach(() => {
    vi.clearAllMocks();
    mockServer = { tool: vi.fn() };
    mockGraphClient = {
      graphRequest: vi.fn().mockResolvedValue({
        content: [{ type: 'text', text: JSON.stringify({ value: [] }) }],
      }),
    } as unknown as GraphClient;
  });

  function getToolHandler(toolName: string) {
    // orgMode=true so the shared-mailbox (work-scope) tools register
    registerGraphTools(mockServer, mockGraphClient, false, undefined, true);
    const call = mockServer.tool.mock.calls.find((c: unknown[]) => c[0] === toolName);
    expect(call).toBeDefined();
    return call![call!.length - 1] as (params: Record<string, unknown>) => Promise<unknown>;
  }

  const calledPath = () =>
    (mockGraphClient.graphRequest as ReturnType<typeof vi.fn>).mock.calls[0][0] as string;

  it.each(MAIL_TOOLS)(
    '%s: rejects an offset-less date literal and never calls Graph',
    async (toolName, pathParams) => {
      const handler = getToolHandler(toolName);
      const result = (await handler({
        ...pathParams,
        filter: 'receivedDateTime ge 2026-09-22T00:00:00',
      })) as { isError?: boolean; content: { text: string }[] };
      expect(result.isError).toBe(true);
      expect(JSON.parse(result.content[0].text).error).toMatch(/has no UTC offset/);
      expect(mockGraphClient.graphRequest).not.toHaveBeenCalled();
    }
  );

  it.each(MAIL_TOOLS)(
    '%s: rejects a date-only literal (Graph accepts it but evaluates it in UTC)',
    async (toolName, pathParams) => {
      const handler = getToolHandler(toolName);
      const result = (await handler({
        ...pathParams,
        filter: 'receivedDateTime ge 2026-09-22',
      })) as { isError?: boolean };
      expect(result.isError).toBe(true);
      expect(mockGraphClient.graphRequest).not.toHaveBeenCalled();
    }
  );

  it.each(MAIL_TOOLS)(
    '%s: passes an offset-qualified filter through byte-for-byte',
    async (toolName, pathParams) => {
      const handler = getToolHandler(toolName);
      const filter =
        'receivedDateTime ge 2026-09-22T00:00:00-07:00 and receivedDateTime lt 2026-09-23T00:00:00+05:30';
      await handler({ ...pathParams, filter });
      // '+' must be percent-encoded or Graph would read it as a space and mangle the offset
      expect(calledPath()).toContain(`$filter=${encodeURIComponent(filter)}`);
      expect(calledPath()).toContain('%2B05%3A30');
    }
  );

  it('leaves non-date mail filters alone', async () => {
    const handler = getToolHandler('list-mail-messages');
    await handler({ filter: 'isRead eq false' });
    expect(calledPath()).toContain(`$filter=${encodeURIComponent('isRead eq false')}`);
  });

  it('does not guard non-mail tools', async () => {
    const handler = getToolHandler('list-chat-messages');
    const result = (await handler({
      chatId: 'chat-1',
      filter: 'lastModifiedDateTime gt 2026-09-22T00:00:00',
    })) as { isError?: boolean };
    expect(result.isError).toBeUndefined();
    expect(mockGraphClient.graphRequest).toHaveBeenCalled();
  });
});
