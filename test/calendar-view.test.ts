import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { z } from 'zod';
import { registerGraphTools } from '../src/graph-tools.js';
import type { GraphClient } from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
    warn: vi.fn(),
  },
}));

vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [
      {
        alias: 'get-calendar-view',
        method: 'get',
        path: '/me/calendarView',
        description: 'The calendar view for the calendar.',
        parameters: [
          { name: 'startDateTime', type: 'Query', schema: z.string() },
          { name: 'endDateTime', type: 'Query', schema: z.string() },
          { name: 'top', type: 'Query', schema: z.number().int().optional() },
          { name: 'skip', type: 'Query', schema: z.number().int().optional() },
          { name: 'select', type: 'Query', schema: z.array(z.string()).optional() },
          { name: 'orderby', type: 'Query', schema: z.array(z.string()).optional() },
          { name: 'filter', type: 'Query', schema: z.string().optional() },
          { name: 'expand', type: 'Query', schema: z.array(z.string()).optional() },
        ],
      },
      {
        alias: 'get-specific-calendar-view',
        method: 'get',
        path: '/me/calendars/:calendarId/calendarView',
        description: 'The calendar view for a specific calendar.',
        parameters: [
          { name: 'calendarId', type: 'Path', schema: z.string() },
          { name: 'startDateTime', type: 'Query', schema: z.string() },
          { name: 'endDateTime', type: 'Query', schema: z.string() },
          { name: 'top', type: 'Query', schema: z.number().int().optional() },
          { name: 'skip', type: 'Query', schema: z.number().int().optional() },
          { name: 'select', type: 'Query', schema: z.array(z.string()).optional() },
          { name: 'orderby', type: 'Query', schema: z.array(z.string()).optional() },
          { name: 'filter', type: 'Query', schema: z.string().optional() },
          { name: 'expand', type: 'Query', schema: z.array(z.string()).optional() },
        ],
      },
      {
        alias: 'list-calendar-event-instances',
        method: 'get',
        path: '/me/calendars/:calendarId/events/:eventId/instances',
        description: 'Expand recurring event instances.',
        parameters: [
          { name: 'calendarId', type: 'Path', schema: z.string() },
          { name: 'eventId', type: 'Path', schema: z.string() },
          { name: 'startDateTime', type: 'Query', schema: z.string() },
          { name: 'endDateTime', type: 'Query', schema: z.string() },
        ],
      },
      {
        alias: 'get-shared-calendar-view',
        method: 'get',
        path: '/users/:userId/calendarView',
        description: "The calendar view for another user's calendar.",
        parameters: [
          { name: 'userId', type: 'Path', schema: z.string() },
          { name: 'startDateTime', type: 'Query', schema: z.string() },
          { name: 'endDateTime', type: 'Query', schema: z.string() },
        ],
      },
      {
        alias: 'get-group-calendar-view',
        method: 'get',
        path: '/groups/:groupId/calendarView',
        description: "The calendar view for a group's calendar.",
        parameters: [
          { name: 'groupId', type: 'Path', schema: z.string() },
          { name: 'startDateTime', type: 'Query', schema: z.string() },
          { name: 'endDateTime', type: 'Query', schema: z.string() },
        ],
      },
      {
        alias: 'list-calendar-view-delta',
        method: 'get',
        path: '/me/calendarView/delta()',
        description: 'Delta of a calendar view.',
        parameters: [
          { name: 'startDateTime', type: 'Query', schema: z.string() },
          { name: 'endDateTime', type: 'Query', schema: z.string() },
        ],
      },
      {
        alias: 'list-calendar-events',
        method: 'get',
        path: '/me/events',
        description: 'List events (no window params; not normalized).',
        parameters: [{ name: 'filter', type: 'Query', schema: z.string().optional() }],
      },
    ],
  },
}));

// Every tool whose startDateTime/endDateTime window is offset-normalized, with the path
// params each one needs so a single parametrized test can drive all of them.
const WINDOW_TOOLS: Array<[string, Record<string, string>]> = [
  ['get-calendar-view', {}],
  ['get-specific-calendar-view', { calendarId: 'cal-abc-123' }],
  ['get-shared-calendar-view', { userId: 'someone@example.com' }],
  ['get-group-calendar-view', { groupId: 'group-123' }],
  ['list-calendar-event-instances', { calendarId: 'cal-abc-123', eventId: 'event-xyz-456' }],
  ['list-calendar-view-delta', {}],
];

describe('Calendar View Tools', () => {
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
    // orgMode=true so work-scope tools (get-shared-calendar-view, get-group-calendar-view) register
    registerGraphTools(mockServer, mockGraphClient, false, undefined, true);
    const call = mockServer.tool.mock.calls.find((c: unknown[]) => c[0] === toolName);
    expect(call).toBeDefined();
    return call![call!.length - 1] as (params: Record<string, unknown>) => Promise<unknown>;
  }

  describe('tool registration', () => {
    it('should register all three calendar view/instances tools', () => {
      registerGraphTools(mockServer, mockGraphClient, false);

      const toolNames = mockServer.tool.mock.calls.map((call: unknown[]) => call[0]);
      expect(toolNames).toContain('get-calendar-view');
      expect(toolNames).toContain('get-specific-calendar-view');
      expect(toolNames).toContain('list-calendar-event-instances');
    });

    it('should include timezone parameter for calendar view tools', () => {
      registerGraphTools(mockServer, mockGraphClient, false);

      for (const call of mockServer.tool.mock.calls) {
        const toolName = call[0] as string;
        const paramSchema = call[2] as Record<string, z.ZodTypeAny>;

        if (WINDOW_TOOLS.map(([name]) => name).includes(toolName)) {
          expect(paramSchema).toHaveProperty('timezone');
        }
      }
    });

    it('should include expandExtendedProperties parameter for calendar view tools', () => {
      registerGraphTools(mockServer, mockGraphClient, false);

      for (const call of mockServer.tool.mock.calls) {
        const toolName = call[0] as string;
        const paramSchema = call[2] as Record<string, z.ZodTypeAny>;

        if (
          [
            'get-calendar-view',
            'get-specific-calendar-view',
            'list-calendar-event-instances',
          ].includes(toolName)
        ) {
          expect(paramSchema).toHaveProperty('expandExtendedProperties');
        }
      }
    });

    it('should include fetchAllPages parameter for GET tools', () => {
      registerGraphTools(mockServer, mockGraphClient, false);

      for (const call of mockServer.tool.mock.calls) {
        const toolName = call[0] as string;
        // Skip utility tools that are not Graph API endpoints
        if (toolName === 'parse-teams-url' || toolName === 'download-bytes') continue;
        const paramSchema = call[2] as Record<string, z.ZodTypeAny>;
        expect(paramSchema).toHaveProperty('fetchAllPages');
      }
    });

    it('should append llmTip to tool descriptions', () => {
      registerGraphTools(mockServer, mockGraphClient, false);

      for (const call of mockServer.tool.mock.calls) {
        const toolName = call[0] as string;
        const description = call[1] as string;

        if (toolName === 'get-calendar-view') {
          expect(description).toContain('TIP:');
          expect(description).toContain('recurring event instances');
          expect(description).toContain('WITH a UTC offset');
          expect(description).toContain('-07:00');
        }
        if (toolName === 'get-specific-calendar-view') {
          expect(description).toContain('TIP:');
          expect(description).toContain('recurring event instances');
          expect(description).toContain('WITH a UTC offset');
          expect(description).toContain('-07:00');
        }
        if (
          [
            'list-calendar-event-instances',
            'get-shared-calendar-view',
            'get-group-calendar-view',
            'list-calendar-view-delta',
          ].includes(toolName)
        ) {
          expect(description).toContain('TIP:');
          expect(description).toContain('WITH a UTC offset');
          expect(description).toContain('-07:00');
        }
      }
    });
  });

  describe('tool execution', () => {
    it('should call graphRequest with correct path for specific calendar view', async () => {
      const handler = getToolHandler('get-specific-calendar-view');

      await handler({
        calendarId: 'cal-abc-123',
        startDateTime: '2024-01-01T00:00:00Z',
        endDateTime: '2024-01-31T23:59:59Z',
      });

      expect(mockGraphClient.graphRequest).toHaveBeenCalledWith(
        expect.stringContaining('/me/calendars/cal-abc-123/calendarView'),
        expect.objectContaining({ method: 'GET' })
      );

      // Verify startDateTime and endDateTime are in the path as query params
      const calledPath = (mockGraphClient.graphRequest as ReturnType<typeof vi.fn>).mock
        .calls[0][0] as string;
      expect(calledPath).toContain('startDateTime=2024-01-01T00%3A00%3A00Z');
      expect(calledPath).toContain('endDateTime=2024-01-31T23%3A59%3A59Z');
    });

    it('should set timezone header when timezone param is provided', async () => {
      const handler = getToolHandler('get-specific-calendar-view');

      await handler({
        calendarId: 'cal-abc-123',
        startDateTime: '2024-01-01T00:00:00Z',
        endDateTime: '2024-01-31T23:59:59Z',
        timezone: 'Australia/Sydney',
      });

      expect(mockGraphClient.graphRequest).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          headers: expect.objectContaining({
            Prefer: expect.stringContaining('outlook.timezone="Australia/Sydney"'),
          }),
        })
      );
    });

    it('should add $expand for extended properties when requested', async () => {
      const handler = getToolHandler('get-specific-calendar-view');

      await handler({
        calendarId: 'cal-abc-123',
        startDateTime: '2024-01-01T00:00:00Z',
        endDateTime: '2024-01-31T23:59:59Z',
        expandExtendedProperties: true,
      });

      const calledPath = (mockGraphClient.graphRequest as ReturnType<typeof vi.fn>).mock
        .calls[0][0] as string;
      expect(calledPath).toContain('$expand=singleValueExtendedProperties');
    });

    it('should append to existing $expand when expandExtendedProperties is set', async () => {
      const handler = getToolHandler('get-specific-calendar-view');

      await handler({
        calendarId: 'cal-abc-123',
        startDateTime: '2024-01-01T00:00:00Z',
        endDateTime: '2024-01-31T23:59:59Z',
        expand: ['extensions'],
        expandExtendedProperties: true,
      });

      const calledPath = (mockGraphClient.graphRequest as ReturnType<typeof vi.fn>).mock
        .calls[0][0] as string;
      expect(calledPath).toContain('$expand=extensions,singleValueExtendedProperties');
    });

    it('should pass $top query parameter when provided', async () => {
      const handler = getToolHandler('get-specific-calendar-view');

      await handler({
        calendarId: 'cal-abc-123',
        startDateTime: '2024-01-01T00:00:00Z',
        endDateTime: '2024-01-31T23:59:59Z',
        top: 50,
      });

      const calledPath = (mockGraphClient.graphRequest as ReturnType<typeof vi.fn>).mock
        .calls[0][0] as string;
      expect(calledPath).toContain('$top=50');
    });

    describe('UTC offset normalization (bare datetime is otherwise read as UTC by Graph)', () => {
      const calledPath = () =>
        (mockGraphClient.graphRequest as ReturnType<typeof vi.fn>).mock.calls[0][0] as string;

      // Hermetic: these cases describe the no-default behavior, whatever the runner's env holds.
      beforeEach(() => {
        vi.stubEnv('MS365_MCP_DEFAULT_TIMEZONE', '');
      });
      afterEach(() => {
        vi.unstubAllEnvs();
      });

      it.each(WINDOW_TOOLS)(
        '%s: injects PDT offset from timezone on a DST-active date',
        async (toolName, pathParams) => {
          const handler = getToolHandler(toolName);
          await handler({
            ...pathParams,
            startDateTime: '2026-09-22T00:00:00',
            endDateTime: '2026-09-22T23:59:59',
            timezone: 'America/Los_Angeles',
          });
          expect(calledPath()).toContain('startDateTime=2026-09-22T00%3A00%3A00-07%3A00');
          expect(calledPath()).toContain('endDateTime=2026-09-22T23%3A59%3A59-07%3A00');
          // Display preference is still sent alongside the corrected window
          expect(mockGraphClient.graphRequest).toHaveBeenCalledWith(
            expect.any(String),
            expect.objectContaining({
              headers: expect.objectContaining({
                Prefer: expect.stringContaining('outlook.timezone="America/Los_Angeles"'),
              }),
            })
          );
        }
      );

      it.each(WINDOW_TOOLS)(
        '%s: injects PST offset on a DST-inactive date',
        async (toolName, pathParams) => {
          const handler = getToolHandler(toolName);
          await handler({
            ...pathParams,
            startDateTime: '2026-01-15T00:00:00',
            endDateTime: '2026-01-16T00:00:00',
            timezone: 'America/Los_Angeles',
          });
          expect(calledPath()).toContain('startDateTime=2026-01-15T00%3A00%3A00-08%3A00');
          expect(calledPath()).toContain('endDateTime=2026-01-16T00%3A00%3A00-08%3A00');
        }
      );

      it.each(WINDOW_TOOLS)(
        '%s: rejects a bare datetime with no timezone and never calls Graph',
        async (toolName, pathParams) => {
          const handler = getToolHandler(toolName);
          const result = (await handler({
            ...pathParams,
            startDateTime: '2026-09-22T00:00:00',
            endDateTime: '2026-09-23T00:00:00Z',
          })) as { isError?: boolean; content: { text: string }[] };
          expect(result.isError).toBe(true);
          expect(JSON.parse(result.content[0].text).error).toMatch(
            /startDateTime\/endDateTime must include a UTC offset, or pass `timezone`/
          );
          expect(mockGraphClient.graphRequest).not.toHaveBeenCalled();
        }
      );

      it.each(WINDOW_TOOLS)(
        '%s: passes offset-qualified datetimes through unchanged even with timezone',
        async (toolName, pathParams) => {
          const handler = getToolHandler(toolName);
          await handler({
            ...pathParams,
            startDateTime: '2026-09-22T00:00:00-07:00',
            endDateTime: '2026-09-23T07:00:00Z',
            timezone: 'America/Los_Angeles',
          });
          expect(calledPath()).toContain('startDateTime=2026-09-22T00%3A00%3A00-07%3A00');
          expect(calledPath()).toContain('endDateTime=2026-09-23T07%3A00%3A00Z');
          expect(calledPath()).not.toContain('-07%3A00-07%3A00');
        }
      );

      it('does not touch tools outside the window-param scope', async () => {
        const handler = getToolHandler('list-calendar-events');
        const result = (await handler({
          filter: "start/dateTime ge '2026-09-22T00:00:00'",
        })) as { isError?: boolean };
        expect(result.isError).toBeUndefined();
        expect(calledPath()).toContain(
          encodeURIComponent("start/dateTime ge '2026-09-22T00:00:00'")
        );
      });
    });

    describe('server default timezone fallback (MS365_MCP_DEFAULT_TIMEZONE)', () => {
      const calledPath = () =>
        (mockGraphClient.graphRequest as ReturnType<typeof vi.fn>).mock.calls[0][0] as string;
      const calledOptions = () =>
        (mockGraphClient.graphRequest as ReturnType<typeof vi.fn>).mock.calls[0][1] as {
          headers?: Record<string, string>;
        };

      afterEach(() => {
        vi.unstubAllEnvs();
      });

      it.each(WINDOW_TOOLS)(
        '%s: default set, no offset, no timezone param -> PDT on a DST-active date',
        async (toolName, pathParams) => {
          vi.stubEnv('MS365_MCP_DEFAULT_TIMEZONE', 'America/Los_Angeles');
          const handler = getToolHandler(toolName);
          const result = (await handler({
            ...pathParams,
            startDateTime: '2026-09-22T00:00:00',
            endDateTime: '2026-09-22T23:59:59',
          })) as { isError?: boolean };
          expect(result.isError).toBeUndefined();
          expect(calledPath()).toContain('startDateTime=2026-09-22T00%3A00%3A00-07%3A00');
          expect(calledPath()).toContain('endDateTime=2026-09-22T23%3A59%3A59-07%3A00');
          // The default now also drives the display-timezone preference (fix/response-timezone-local):
          // previously it fixed the query window only and results still came back in UTC.
          expect(calledOptions().headers?.Prefer ?? '').toContain(
            'outlook.timezone="America/Los_Angeles"'
          );
        }
      );

      it.each(WINDOW_TOOLS)(
        '%s: default set, no offset, no timezone param -> PST on a DST-inactive date',
        async (toolName, pathParams) => {
          vi.stubEnv('MS365_MCP_DEFAULT_TIMEZONE', 'America/Los_Angeles');
          const handler = getToolHandler(toolName);
          await handler({
            ...pathParams,
            startDateTime: '2026-01-15T00:00:00',
            endDateTime: '2026-01-16',
          });
          expect(calledPath()).toContain('startDateTime=2026-01-15T00%3A00%3A00-08%3A00');
          expect(calledPath()).toContain('endDateTime=2026-01-16T00%3A00%3A00-08%3A00');
        }
      );

      it.each(WINDOW_TOOLS)(
        '%s: explicit timezone param beats the default',
        async (toolName, pathParams) => {
          vi.stubEnv('MS365_MCP_DEFAULT_TIMEZONE', 'America/Los_Angeles');
          const handler = getToolHandler(toolName);
          await handler({
            ...pathParams,
            startDateTime: '2026-09-22T00:00:00',
            endDateTime: '2026-09-23T00:00:00',
            timezone: 'Asia/Kolkata',
          });
          expect(calledPath()).toContain('startDateTime=2026-09-22T00%3A00%3A00%2B05%3A30');
          expect(calledPath()).toContain('endDateTime=2026-09-23T00%3A00%3A00%2B05%3A30');
          expect(calledPath()).not.toContain('-07%3A00');
          // The explicit `timezone` param wins for the DISPLAY header too, not just the window.
          expect(calledOptions().headers?.Prefer ?? '').toContain(
            'outlook.timezone="Asia/Kolkata"'
          );
          expect(calledOptions().headers?.Prefer ?? '').not.toContain('America/Los_Angeles');
        }
      );

      it.each(WINDOW_TOOLS)(
        '%s: no default configured, no timezone param -> no outlook.timezone Prefer header',
        async (toolName, pathParams) => {
          vi.stubEnv('MS365_MCP_DEFAULT_TIMEZONE', undefined);
          const handler = getToolHandler(toolName);
          const result = (await handler({
            ...pathParams,
            startDateTime: '2026-09-22T00:00:00-07:00',
            endDateTime: '2026-09-23T00:00:00-07:00',
          })) as { isError?: boolean };
          expect(result.isError).toBeUndefined();
          expect(calledOptions().headers?.Prefer ?? '').not.toContain('outlook.timezone');
        }
      );

      it.each(WINDOW_TOOLS)(
        '%s: default set -> exactly one outlook.timezone entry in Prefer, not duplicated',
        async (toolName, pathParams) => {
          vi.stubEnv('MS365_MCP_DEFAULT_TIMEZONE', 'America/Los_Angeles');
          const handler = getToolHandler(toolName);
          await handler({
            ...pathParams,
            startDateTime: '2026-09-22T00:00:00',
            endDateTime: '2026-09-22T23:59:59',
          });
          const prefer = calledOptions().headers?.Prefer ?? '';
          expect(prefer.split('outlook.timezone').length - 1).toBe(1);
        }
      );

      it.each(WINDOW_TOOLS)(
        '%s: offset already present -> passthrough, default never consulted',
        async (toolName, pathParams) => {
          // An invalid default would error if it were consulted at all.
          vi.stubEnv('MS365_MCP_DEFAULT_TIMEZONE', 'Not/AZone');
          const handler = getToolHandler(toolName);
          const result = (await handler({
            ...pathParams,
            startDateTime: '2026-09-22T00:00:00-08:00',
            endDateTime: '2026-09-23T07:00:00Z',
          })) as { isError?: boolean };
          expect(result.isError).toBeUndefined();
          expect(calledPath()).toContain('startDateTime=2026-09-22T00%3A00%3A00-08%3A00');
          expect(calledPath()).toContain('endDateTime=2026-09-23T07%3A00%3A00Z');
        }
      );

      it.each(WINDOW_TOOLS)(
        '%s: REGRESSION GUARD - default unset/blank keeps the original hard error',
        async (toolName, pathParams) => {
          for (const unset of [undefined, '', '   ']) {
            vi.clearAllMocks();
            vi.stubEnv('MS365_MCP_DEFAULT_TIMEZONE', unset);
            const handler = getToolHandler(toolName);
            const result = (await handler({
              ...pathParams,
              startDateTime: '2026-09-22T00:00:00',
              endDateTime: '2026-09-23T00:00:00Z',
            })) as { isError?: boolean; content: { text: string }[] };
            expect(result.isError).toBe(true);
            // Byte-for-byte the pre-default message.
            expect(JSON.parse(result.content[0].text).error).toBe(
              'startDateTime "2026-09-22T00:00:00" has no UTC offset. Microsoft Graph treats an ' +
                'offset-less startDateTime/endDateTime as UTC, which silently shifts the query window. ' +
                'startDateTime/endDateTime must include a UTC offset, or pass `timezone` so it can be ' +
                'inferred (e.g. "2026-09-22T00:00:00-07:00", or timezone "America/Los_Angeles").'
            );
            expect(mockGraphClient.graphRequest).not.toHaveBeenCalled();
          }
        }
      );

      it('misconfigured default + bare datetime -> error naming the env var, Graph not called', async () => {
        vi.stubEnv('MS365_MCP_DEFAULT_TIMEZONE', 'Not/AZone');
        const handler = getToolHandler('get-calendar-view');
        const result = (await handler({
          startDateTime: '2026-09-22T00:00:00',
          endDateTime: '2026-09-23T00:00:00',
        })) as { isError?: boolean; content: { text: string }[] };
        expect(result.isError).toBe(true);
        expect(JSON.parse(result.content[0].text).error).toMatch(
          /MS365_MCP_DEFAULT_TIMEZONE is set to "Not\/AZone"/
        );
        expect(mockGraphClient.graphRequest).not.toHaveBeenCalled();
      });

      it('does not touch tools outside the window-param scope even with a default', async () => {
        vi.stubEnv('MS365_MCP_DEFAULT_TIMEZONE', 'America/Los_Angeles');
        const handler = getToolHandler('list-calendar-events');
        await handler({ filter: "start/dateTime ge '2026-09-22T00:00:00'" });
        expect(calledPath()).toContain(
          encodeURIComponent("start/dateTime ge '2026-09-22T00:00:00'")
        );
      });
    });

    it('should call graphRequest with correct path for event instances', async () => {
      const handler = getToolHandler('list-calendar-event-instances');

      await handler({
        calendarId: 'cal-abc-123',
        eventId: 'event-xyz-456',
        startDateTime: '2024-01-01T00:00:00Z',
        endDateTime: '2024-12-31T23:59:59Z',
      });

      expect(mockGraphClient.graphRequest).toHaveBeenCalledWith(
        expect.stringContaining('/me/calendars/cal-abc-123/events/event-xyz-456/instances'),
        expect.objectContaining({ method: 'GET' })
      );
    });
  });
});
