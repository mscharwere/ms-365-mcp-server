import { beforeEach, describe, expect, it, vi } from 'vitest';
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
    ],
  },
}));

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
    registerGraphTools(mockServer, mockGraphClient, false);
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

        if (
          [
            'get-calendar-view',
            'get-specific-calendar-view',
            'list-calendar-event-instances',
          ].includes(toolName)
        ) {
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
        if (toolName === 'list-calendar-event-instances') {
          expect(description).toContain('TIP:');
          expect(description).toContain('startDateTime and endDateTime');
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

      it.each(['get-calendar-view', 'get-specific-calendar-view'])(
        '%s: injects PDT offset from timezone on a DST-active date',
        async (toolName) => {
          const handler = getToolHandler(toolName);
          await handler({
            calendarId: 'cal-abc-123',
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

      it('injects PST offset on a DST-inactive date', async () => {
        const handler = getToolHandler('get-calendar-view');
        await handler({
          startDateTime: '2026-01-15T00:00:00',
          endDateTime: '2026-01-16T00:00:00',
          timezone: 'America/Los_Angeles',
        });
        expect(calledPath()).toContain('startDateTime=2026-01-15T00%3A00%3A00-08%3A00');
        expect(calledPath()).toContain('endDateTime=2026-01-16T00%3A00%3A00-08%3A00');
      });

      it.each(['get-calendar-view', 'get-specific-calendar-view'])(
        '%s: rejects a bare datetime with no timezone and never calls Graph',
        async (toolName) => {
          const handler = getToolHandler(toolName);
          const result = (await handler({
            calendarId: 'cal-abc-123',
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

      it('passes offset-qualified datetimes through unchanged even with timezone', async () => {
        const handler = getToolHandler('get-specific-calendar-view');
        await handler({
          calendarId: 'cal-abc-123',
          startDateTime: '2026-09-22T00:00:00-07:00',
          endDateTime: '2026-09-23T07:00:00Z',
          timezone: 'America/Los_Angeles',
        });
        expect(calledPath()).toContain('startDateTime=2026-09-22T00%3A00%3A00-07%3A00');
        expect(calledPath()).toContain('endDateTime=2026-09-23T07%3A00%3A00Z');
        expect(calledPath()).not.toContain('-07%3A00-07%3A00');
      });

      it('does not change behavior for tools outside the calendarView scope', async () => {
        const handler = getToolHandler('list-calendar-event-instances');
        await handler({
          calendarId: 'cal-abc-123',
          eventId: 'event-xyz-456',
          startDateTime: '2026-09-22T00:00:00',
          endDateTime: '2026-09-23T00:00:00',
        });
        expect(calledPath()).toContain('startDateTime=2026-09-22T00%3A00%3A00&');
        expect(calledPath()).toMatch(/endDateTime=2026-09-23T00%3A00%3A00($|&)/);
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
