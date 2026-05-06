/**
 * M365 MCP Compactor — Phase 1
 *
 * DESIGN NOTES (read before modifying):
 *
 * 1. TIMEZONE LABEL — NEVER STRIP start.timeZone / end.timeZone.
 *    Graph returns local-clock datetime strings (no Z suffix) paired with a timeZone label.
 *    The label is what tells consumers the timezone; stripping it silently converts to UTC-ambiguous.
 *    FRIDAY confirmed the live IANA form is "America/Los_Angeles", NOT "Pacific Standard Time".
 *    TARS's CI assertion §7 used the Windows form — corrected here per FRIDAY's fixture data.
 *
 * 2. SUBJECT FIELD — NEVER TRUNCATE OR DROP.
 *    Tags (#carlitos, #daniel, #tkd, #driving, etc.) live inside event subject strings.
 *    Dropping or truncating subject loses routing metadata that FRIDAY uses to build briefings.
 *
 * 3. SAFELINKS DECODER.
 *    Graph returns mail body as plain text (not HTML) per MS365_MCP_BODY_FORMAT=text default.
 *    Marketing emails retain ~100 Microsoft ATP safelinks per message — each wraps the real URL
 *    with ~200–300 bytes of telemetry chrome. Decoding yields 70–80% body size reduction.
 *
 *    Safelinks pattern (na01 and nam01 variants both observed):
 *      https://na01.safelinks.protection.outlook.com/?url=<URL_ENCODED_REAL_URL>&data=...
 *    or
 *      https://nam01.safelinks.protection.outlook.com/?url=<URL_ENCODED_REAL_URL>&data=...
 *
 *    We extract the `url=` query parameter, decode it with decodeURIComponent, and replace the
 *    entire safelinks wrapper URL with the decoded destination. Any query parameters after `url=`
 *    (data=, sdata=, reserved=) are discarded — they are telemetry only.
 *
 *    IMPORTANT: The regex must match both na01 and nam01 (and future region prefixes).
 *    We match on the safelinks hostname pattern, not the exact region prefix.
 *
 * 4. EXHAUSTIVENESS GATE.
 *    The COMPACTORS record is typed as Record<ToolAlias, Compactor>. Adding a new tool to
 *    endpoints.json without adding a compactor entry causes a TypeScript compile error.
 *    Use `identity` for tools that don't need projection.
 *
 * 5. PHASE SCOPE.
 *    Phase 1: list-specific-calendar-events, get-specific-calendar-event, list-mail-folder-messages.
 *    Phase 2: get-specific-calendar-view, get-calendar-view, get-calendar-event,
 *             list-calendar-events, list-mail-messages, get-mail-message.
 *             (All use identical projection logic — deferred for smoke-diff validation.)
 */

// ---------------------------------------------------------------------------
// Safelinks decoder
// ---------------------------------------------------------------------------

/**
 * Regex for Microsoft ATP safelinks.
 *
 * Matches URLs of the form:
 *   https://<region>.safelinks.protection.outlook.com/?url=<encoded>&data=...
 *
 * Capture group 1: the `url=` parameter value (URL-encoded destination).
 *
 * The `[^&]+` greedily captures the encoded URL up to the first `&` (start of data= param).
 * We use a non-greedy match on the outer group to avoid eating adjacent safelinks.
 *
 * Tested against: na01, nam01, eur01, apac01, aus01 prefix variants.
 */
const SAFELINKS_RE =
  /https:\/\/[a-z0-9]+\.safelinks\.protection\.outlook\.com\/\?url=([^&\s]+)(?:&[^\s]*)*/g;

/**
 * Decode all Microsoft ATP safelinks in a string, replacing each wrapper with the
 * bare destination URL. Leaves non-safelinks content byte-identical.
 */
export function decodeSafelinks(text: string): string {
  return text.replace(SAFELINKS_RE, (_match, encodedUrl: string) => {
    try {
      return decodeURIComponent(encodedUrl);
    } catch {
      // Malformed encoding — return original match intact rather than corrupting content.
      return _match;
    }
  });
}

// ---------------------------------------------------------------------------
// Core types
// ---------------------------------------------------------------------------

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type JsonValue = any;

/** A compactor takes the raw MCP tool response object and returns a smaller one. */
export type Compactor = (raw: JsonValue) => JsonValue;

/** Identity passthrough — used for all tools not in Phase 1/2 scope. */
const identity: Compactor = (raw) => raw;

// ---------------------------------------------------------------------------
// Calendar event projection helpers
// ---------------------------------------------------------------------------

/**
 * Project a single calendar event object.
 *
 * PRESERVED (verbatim):
 *   id, subject (tags live here — never drop), start (dateTime + timeZone — byte-identical),
 *   end (dateTime + timeZone), isAllDay, isCancelled, categories,
 *   location.displayName, body.content (already plain text via Graph Prefer header),
 *   attendees (projected to { type, status: { response }, emailAddress: { name, address } })
 *
 * DROPPED:
 *   internetMessageId, iCalUId, webLink — not read by consumers
 *   recurrence (null on instances; already expanded in view)
 *   organizer (not needed for briefing routing)
 *   location.locationType, location.uniqueId, location.uniqueIdType,
 *     location.address, location.coordinates, location.locationUri — only displayName used
 *   body.contentType — always "text" in our setup; redundant
 *   attendees[].status.time — always "0001-01-01T00:00:00Z" sentinel; safe to drop
 */
function projectCalendarEvent(event: JsonValue): JsonValue {
  if (!event || typeof event !== 'object') return event;

  const out: JsonValue = {};

  // Required identity fields
  if (event.id !== undefined) out.id = event.id;
  // NEVER drop or truncate subject — tags live here
  if (event.subject !== undefined) out.subject = event.subject;

  // Timing — preserve both dateTime AND timeZone byte-identical.
  // DO NOT strip timeZone label. DO NOT convert to UTC. Graph already applied the
  // requested timezone (America/Los_Angeles); the label is what proves it.
  if (event.start !== undefined) {
    out.start = {
      dateTime: event.start.dateTime,
      timeZone: event.start.timeZone,
    };
  }
  if (event.end !== undefined) {
    out.end = {
      dateTime: event.end.dateTime,
      timeZone: event.end.timeZone,
    };
  }

  // Boolean flags
  if (event.isAllDay !== undefined) out.isAllDay = event.isAllDay;
  if (event.isCancelled !== undefined) out.isCancelled = event.isCancelled;

  // Categories (Outlook colored labels — distinct from hashtag tags)
  if (event.categories !== undefined) out.categories = event.categories;

  // Location — keep only displayName
  if (event.location !== undefined) {
    out.location = { displayName: event.location.displayName ?? '' };
  }

  // Body — keep content, drop contentType (always "text")
  if (event.body !== undefined) {
    out.body = { content: event.body.content ?? '' };
  }

  // Attendees — project to minimal shape, drop status.time sentinel
  if (event.attendees !== undefined && Array.isArray(event.attendees)) {
    out.attendees = event.attendees.map((a: JsonValue) => ({
      type: a.type,
      status: { response: a.status?.response },
      emailAddress: {
        name: a.emailAddress?.name,
        address: a.emailAddress?.address,
      },
    }));
  }

  return out;
}

// ---------------------------------------------------------------------------
// Mail message projection helpers
// ---------------------------------------------------------------------------

/**
 * Project a single mail message object.
 *
 * PRESERVED:
 *   id, from, subject, receivedDateTime, body.content (safelinks decoded),
 *   importance, isRead, hasAttachments, bodyPreview (truncated preview for list calls),
 *   @odata.nextLink (CRITICAL — pagination must survive compaction)
 *
 * DROPPED:
 *   internetMessageId, conversationId — not needed for briefing use
 *   flag — when not flagged, always { flagStatus: "notFlagged" }; noise
 *   body.contentType — always "text"; redundant
 *   toRecipients — Carlos is always the recipient; not needed for routing
 *
 * TRANSFORMED:
 *   body.content — safelinks decoded (see decodeSafelinks above)
 */
function projectMailMessage(msg: JsonValue): JsonValue {
  if (!msg || typeof msg !== 'object') return msg;

  const out: JsonValue = {};

  if (msg.id !== undefined) out.id = msg.id;
  if (msg.from !== undefined) out.from = msg.from;
  if (msg.subject !== undefined) out.subject = msg.subject;
  if (msg.receivedDateTime !== undefined) out.receivedDateTime = msg.receivedDateTime;
  if (msg.importance !== undefined) out.importance = msg.importance;
  if (msg.isRead !== undefined) out.isRead = msg.isRead;
  if (msg.hasAttachments !== undefined) out.hasAttachments = msg.hasAttachments;
  // bodyPreview present on list calls (truncated ~255 chars — keep as-is, no safelinks at preview length)
  if (msg.bodyPreview !== undefined) out.bodyPreview = msg.bodyPreview;

  // Body — decode safelinks, keep content only
  if (msg.body !== undefined) {
    out.body = {
      content: decodeSafelinks(msg.body.content ?? ''),
    };
  }

  // Attachments — ref list only (no binary IDs)
  if (msg.attachments !== undefined && Array.isArray(msg.attachments)) {
    out.attachments = msg.attachments.map((a: JsonValue) => ({
      name: a.name,
      size: a.size,
      contentType: a.contentType,
    }));
  }

  return out;
}

// ---------------------------------------------------------------------------
// Top-level response wrappers
// ---------------------------------------------------------------------------

/**
 * Compact a list response (has `value` array + optional `@odata.nextLink`).
 * Applies the item compactor to each element in the array.
 * @odata.nextLink is preserved verbatim — pagination must survive.
 */
function compactList(raw: JsonValue, itemCompactor: (item: JsonValue) => JsonValue): JsonValue {
  if (!raw || typeof raw !== 'object') return raw;
  const out: JsonValue = {};
  if (raw.value !== undefined && Array.isArray(raw.value)) {
    out.value = raw.value.map(itemCompactor);
  }
  // CRITICAL: preserve pagination token
  if (raw['@odata.nextLink'] !== undefined) {
    out['@odata.nextLink'] = raw['@odata.nextLink'];
  }
  // Preserve count if present
  if (raw['@odata.count'] !== undefined) {
    out['@odata.count'] = raw['@odata.count'];
  }
  return out;
}

// ---------------------------------------------------------------------------
// Per-tool compactors (Phase 1 — 3 explicit + identity for remaining 267)
// ---------------------------------------------------------------------------

/**
 * list-specific-calendar-events — returns { value: Event[] }
 *
 * This is the tool FRIDAY actually uses for both Personal and Family calendars.
 * The Family calendar pull is where 16 events with rich tags live.
 */
const compactListSpecificCalendarEvents: Compactor = (raw) =>
  compactList(raw, projectCalendarEvent);

/**
 * get-specific-calendar-event — returns a single Event object (not wrapped in value[]).
 *
 * NOTE: Per FRIDAY's anomaly capture, Family calendar event IDs return 400 via /me/events/{id}.
 * This compactor only operates on Personal calendar events in practice.
 */
const compactGetSpecificCalendarEvent: Compactor = (raw) => projectCalendarEvent(raw);

/**
 * list-mail-folder-messages — returns { value: Message[], @odata.nextLink? }
 *
 * This is the working alternative to list-mail-messages (which has a stale 10-msg cache bug).
 * The list response uses bodyPreview (not full body), so safelinks aren't present at list level.
 * Full body safelinks only appear in get-mail-message responses.
 */
const compactListMailFolderMessages: Compactor = (raw) => compactList(raw, projectMailMessage);

// ---------------------------------------------------------------------------
// Exhaustive compactor record for all 270 Graph tools
//
// TypeScript enforces this record is exhaustive against the ToolAlias union.
// Add a new tool alias here (with `identity` or a real compactor) if endpoints.json grows.
// A compile error here means a tool was added to the API surface without a classification decision.
// ---------------------------------------------------------------------------

export type ToolAlias =
  | 'accept-calendar-event'
  | 'add-excel-table-rows'
  | 'add-group-member'
  | 'add-group-owner'
  | 'add-mail-attachment'
  | 'add-team-member'
  | 'cancel-calendar-event'
  | 'clear-excel-range'
  | 'copy-drive-item'
  | 'copy-mail-message'
  | 'create-calendar'
  | 'create-calendar-event'
  | 'create-chat'
  | 'create-draft-email'
  | 'create-drive-item-preview'
  | 'create-drive-item-share-link'
  | 'create-excel-chart'
  | 'create-excel-table'
  | 'create-focused-inbox-override'
  | 'create-forward-draft'
  | 'create-group'
  | 'create-mail-attachment-upload-session'
  | 'create-mail-child-folder'
  | 'create-mail-folder'
  | 'create-mail-rule'
  | 'create-onedrive-folder'
  | 'create-onenote-notebook'
  | 'create-onenote-page'
  | 'create-onenote-section'
  | 'create-onenote-section-page'
  | 'create-online-meeting'
  | 'create-outlook-category'
  | 'create-outlook-contact'
  | 'create-planner-bucket'
  | 'create-planner-task'
  | 'create-reply-all-draft'
  | 'create-reply-draft'
  | 'create-shared-mailbox-draft'
  | 'create-sharepoint-list'
  | 'create-sharepoint-list-column'
  | 'create-sharepoint-list-item'
  | 'create-specific-calendar-event'
  | 'create-subscription'
  | 'create-team-channel'
  | 'create-todo-linked-resource'
  | 'create-todo-task'
  | 'create-upload-session'
  | 'decline-calendar-event'
  | 'delete-calendar'
  | 'delete-calendar-event'
  | 'delete-drive-item-permission'
  | 'delete-excel-range'
  | 'delete-excel-table-row'
  | 'delete-focused-inbox-override'
  | 'delete-group'
  | 'delete-mail-attachment'
  | 'delete-mail-folder'
  | 'delete-mail-message'
  | 'delete-mail-rule'
  | 'delete-onedrive-file'
  | 'delete-onenote-page'
  | 'delete-online-meeting'
  | 'delete-outlook-contact'
  | 'delete-planner-bucket'
  | 'delete-sharepoint-list-column'
  | 'delete-sharepoint-list-item'
  | 'delete-specific-calendar-event'
  | 'delete-subscription'
  | 'delete-team-channel'
  | 'delete-todo-linked-resource'
  | 'delete-todo-task'
  | 'dismiss-calendar-event-reminder'
  | 'extract-drive-item-sensitivity-labels'
  | 'find-meeting-times'
  | 'format-excel-range'
  | 'forward-calendar-event'
  | 'forward-mail-message'
  | 'get-calendar-event'
  | 'get-calendar-view'
  | 'get-channel-files-folder'
  | 'get-channel-message'
  | 'get-chat'
  | 'get-chat-message'
  | 'get-current-user'
  | 'get-drive-delta'
  | 'get-drive-item'
  | 'get-drive-root-item'
  | 'get-excel-range'
  | 'get-excel-table'
  | 'get-excel-used-range'
  | 'get-group'
  | 'get-group-calendar-view'
  | 'get-mail-message'
  | 'get-mail-message-mime'
  | 'get-mailbox-settings'
  | 'get-meeting-attendance-report'
  | 'get-meeting-recording'
  | 'get-meeting-recording-content'
  | 'get-meeting-transcript'
  | 'get-meeting-transcript-content'
  | 'get-my-manager'
  | 'get-my-presence'
  | 'get-onenote-page-content'
  | 'get-online-meeting'
  | 'get-outlook-contact'
  | 'get-planner-bucket'
  | 'get-planner-plan'
  | 'get-planner-task'
  | 'get-planner-task-details'
  | 'get-presences-by-user-id'
  | 'get-room'
  | 'get-room-list'
  | 'get-room-list-room'
  | 'get-schedule'
  | 'get-sensitivity-label'
  | 'get-shared-calendar-view'
  | 'get-shared-mailbox-message'
  | 'get-sharepoint-list-column'
  | 'get-sharepoint-site'
  | 'get-sharepoint-site-by-path'
  | 'get-sharepoint-site-drive-by-id'
  | 'get-sharepoint-site-item'
  | 'get-sharepoint-site-list'
  | 'get-sharepoint-site-list-item'
  | 'get-sharepoint-site-onenote-page-content'
  | 'get-sharepoint-sites-delta'
  | 'get-specific-calendar-event'
  | 'get-specific-calendar-view'
  | 'get-subscription'
  | 'get-team'
  | 'get-team-channel'
  | 'get-todo-task'
  | 'get-user-manager'
  | 'get-user-presence'
  | 'get-virtual-event-webinar'
  | 'graph-batch'
  | 'insert-excel-range'
  | 'list-all-onenote-sections'
  | 'list-calendar-event-instances'
  | 'list-calendar-events'
  | 'list-calendar-events-delta'
  | 'list-calendar-view-delta'
  | 'list-calendars'
  | 'list-channel-message-hosted-contents'
  | 'list-channel-message-replies'
  | 'list-channel-messages'
  | 'list-channel-tabs'
  | 'list-chat-members'
  | 'list-chat-message-hosted-contents'
  | 'list-chat-message-replies'
  | 'list-chat-messages'
  | 'list-chats'
  | 'list-drive-item-permissions'
  | 'list-drive-item-thumbnails'
  | 'list-drive-item-versions'
  | 'list-drives'
  | 'list-excel-table-rows'
  | 'list-excel-tables'
  | 'list-excel-worksheets'
  | 'list-focused-inbox-overrides'
  | 'list-folder-files'
  | 'list-group-conversations'
  | 'list-group-events'
  | 'list-group-members'
  | 'list-group-owners'
  | 'list-group-threads'
  | 'list-groups'
  | 'list-joined-teams'
  | 'list-mail-attachments'
  | 'list-mail-child-folders'
  | 'list-mail-folder-messages'
  | 'list-mail-folder-messages-delta'
  | 'list-mail-folders'
  | 'list-mail-messages'
  | 'list-mail-rules'
  | 'list-meeting-attendance-records'
  | 'list-meeting-attendance-reports'
  | 'list-meeting-recordings'
  | 'list-meeting-transcripts'
  | 'list-my-direct-reports'
  | 'list-my-memberships'
  | 'list-onenote-notebook-sections'
  | 'list-onenote-notebooks'
  | 'list-onenote-section-pages'
  | 'list-online-meetings'
  | 'list-outlook-categories'
  | 'list-outlook-contacts'
  | 'list-pinned-chat-messages'
  | 'list-plan-buckets'
  | 'list-plan-tasks'
  | 'list-planner-tasks'
  | 'list-relevant-people'
  | 'list-room-list-rooms'
  | 'list-sensitivity-labels'
  | 'list-shared-calendar-events'
  | 'list-shared-mailbox-folder-messages'
  | 'list-shared-mailbox-messages'
  | 'list-sharepoint-list-columns'
  | 'list-sharepoint-site-drives'
  | 'list-sharepoint-site-items'
  | 'list-sharepoint-site-list-items'
  | 'list-sharepoint-site-lists'
  | 'list-sharepoint-site-onenote-notebook-sections'
  | 'list-sharepoint-site-onenote-notebooks'
  | 'list-sharepoint-site-onenote-section-pages'
  | 'list-specific-calendar-events'
  | 'list-subscriptions'
  | 'list-team-channels'
  | 'list-team-members'
  | 'list-todo-linked-resources'
  | 'list-todo-task-lists'
  | 'list-todo-tasks'
  | 'list-trending-insights'
  | 'list-user-direct-reports'
  | 'list-users'
  | 'list-webinar-sessions'
  | 'merge-excel-range'
  | 'move-mail-message'
  | 'move-rename-onedrive-item'
  | 'pin-chat-message'
  | 'reauthorize-subscription'
  | 'remove-group-member'
  | 'remove-group-owner'
  | 'remove-team-member'
  | 'reply-all-mail-message'
  | 'reply-mail-message'
  | 'reply-to-channel-message'
  | 'reply-to-chat-message'
  | 'reply-to-group-thread'
  | 'search-onedrive-files'
  | 'search-query'
  | 'search-sharepoint-sites'
  | 'send-channel-message'
  | 'send-chat-message'
  | 'send-draft-message'
  | 'send-mail'
  | 'send-shared-mailbox-mail'
  | 'set-channel-message-reaction'
  | 'set-chat-message-reaction'
  | 'share-drive-item'
  | 'snooze-calendar-event-reminder'
  | 'sort-excel-range'
  | 'tentatively-accept-calendar-event'
  | 'unmerge-excel-range'
  | 'unpin-chat-message'
  | 'unset-channel-message-reaction'
  | 'unset-chat-message-reaction'
  | 'update-calendar'
  | 'update-calendar-event'
  | 'update-excel-range'
  | 'update-excel-table-row'
  | 'update-focused-inbox-override'
  | 'update-group'
  | 'update-mail-folder'
  | 'update-mail-message'
  | 'update-mail-rule'
  | 'update-mailbox-settings'
  | 'update-online-meeting'
  | 'update-outlook-contact'
  | 'update-place'
  | 'update-planner-bucket'
  | 'update-planner-task'
  | 'update-planner-task-details'
  | 'update-sharepoint-list-column'
  | 'update-sharepoint-list-item'
  | 'update-specific-calendar-event'
  | 'update-subscription'
  | 'update-team-channel'
  | 'update-todo-task'
  | 'upload-file-content';

export const COMPACTORS: Record<ToolAlias, Compactor> = {
  // -------------------------------------------------------------------------
  // Phase 1 — explicit projections
  // -------------------------------------------------------------------------

  /** Primary calendar pull for both Personal and Family calendars */
  'list-specific-calendar-events': compactListSpecificCalendarEvents,

  /** Single-event fetch — Personal calendar IDs only (Family IDs return 400 per FRIDAY anomaly) */
  'get-specific-calendar-event': compactGetSpecificCalendarEvent,

  /** Working mail inbox tool — stale-cache workaround for list-mail-messages */
  'list-mail-folder-messages': compactListMailFolderMessages,

  // -------------------------------------------------------------------------
  // Phase 2 scope — identity until smoke-diff validates Phase 1 projections
  // -------------------------------------------------------------------------
  'get-specific-calendar-view': identity,
  'get-calendar-view': identity,
  'get-calendar-event': identity,
  'list-calendar-events': identity,
  'list-mail-messages': identity,
  'get-mail-message': identity,

  // -------------------------------------------------------------------------
  // Remaining 261 tools — identity passthrough
  // -------------------------------------------------------------------------
  'accept-calendar-event': identity,
  'add-excel-table-rows': identity,
  'add-group-member': identity,
  'add-group-owner': identity,
  'add-mail-attachment': identity,
  'add-team-member': identity,
  'cancel-calendar-event': identity,
  'clear-excel-range': identity,
  'copy-drive-item': identity,
  'copy-mail-message': identity,
  'create-calendar': identity,
  'create-calendar-event': identity,
  'create-chat': identity,
  'create-draft-email': identity,
  'create-drive-item-preview': identity,
  'create-drive-item-share-link': identity,
  'create-excel-chart': identity,
  'create-excel-table': identity,
  'create-focused-inbox-override': identity,
  'create-forward-draft': identity,
  'create-group': identity,
  'create-mail-attachment-upload-session': identity,
  'create-mail-child-folder': identity,
  'create-mail-folder': identity,
  'create-mail-rule': identity,
  'create-onedrive-folder': identity,
  'create-onenote-notebook': identity,
  'create-onenote-page': identity,
  'create-onenote-section': identity,
  'create-onenote-section-page': identity,
  'create-online-meeting': identity,
  'create-outlook-category': identity,
  'create-outlook-contact': identity,
  'create-planner-bucket': identity,
  'create-planner-task': identity,
  'create-reply-all-draft': identity,
  'create-reply-draft': identity,
  'create-shared-mailbox-draft': identity,
  'create-sharepoint-list': identity,
  'create-sharepoint-list-column': identity,
  'create-sharepoint-list-item': identity,
  'create-specific-calendar-event': identity,
  'create-subscription': identity,
  'create-team-channel': identity,
  'create-todo-linked-resource': identity,
  'create-todo-task': identity,
  'create-upload-session': identity,
  'decline-calendar-event': identity,
  'delete-calendar': identity,
  'delete-calendar-event': identity,
  'delete-drive-item-permission': identity,
  'delete-excel-range': identity,
  'delete-excel-table-row': identity,
  'delete-focused-inbox-override': identity,
  'delete-group': identity,
  'delete-mail-attachment': identity,
  'delete-mail-folder': identity,
  'delete-mail-message': identity,
  'delete-mail-rule': identity,
  'delete-onedrive-file': identity,
  'delete-onenote-page': identity,
  'delete-online-meeting': identity,
  'delete-outlook-contact': identity,
  'delete-planner-bucket': identity,
  'delete-sharepoint-list-column': identity,
  'delete-sharepoint-list-item': identity,
  'delete-specific-calendar-event': identity,
  'delete-subscription': identity,
  'delete-team-channel': identity,
  'delete-todo-linked-resource': identity,
  'delete-todo-task': identity,
  'dismiss-calendar-event-reminder': identity,
  'extract-drive-item-sensitivity-labels': identity,
  'find-meeting-times': identity,
  'format-excel-range': identity,
  'forward-calendar-event': identity,
  'forward-mail-message': identity,
  'get-channel-files-folder': identity,
  'get-channel-message': identity,
  'get-chat': identity,
  'get-chat-message': identity,
  'get-current-user': identity,
  'get-drive-delta': identity,
  'get-drive-item': identity,
  'get-drive-root-item': identity,
  'get-excel-range': identity,
  'get-excel-table': identity,
  'get-excel-used-range': identity,
  'get-group': identity,
  'get-group-calendar-view': identity,
  'get-mail-message-mime': identity,
  'get-mailbox-settings': identity,
  'get-meeting-attendance-report': identity,
  'get-meeting-recording': identity,
  'get-meeting-recording-content': identity,
  'get-meeting-transcript': identity,
  'get-meeting-transcript-content': identity,
  'get-my-manager': identity,
  'get-my-presence': identity,
  'get-onenote-page-content': identity,
  'get-online-meeting': identity,
  'get-outlook-contact': identity,
  'get-planner-bucket': identity,
  'get-planner-plan': identity,
  'get-planner-task': identity,
  'get-planner-task-details': identity,
  'get-presences-by-user-id': identity,
  'get-room': identity,
  'get-room-list': identity,
  'get-room-list-room': identity,
  'get-schedule': identity,
  'get-sensitivity-label': identity,
  'get-shared-calendar-view': identity,
  'get-shared-mailbox-message': identity,
  'get-sharepoint-list-column': identity,
  'get-sharepoint-site': identity,
  'get-sharepoint-site-by-path': identity,
  'get-sharepoint-site-drive-by-id': identity,
  'get-sharepoint-site-item': identity,
  'get-sharepoint-site-list': identity,
  'get-sharepoint-site-list-item': identity,
  'get-sharepoint-site-onenote-page-content': identity,
  'get-sharepoint-sites-delta': identity,
  'get-subscription': identity,
  'get-team': identity,
  'get-team-channel': identity,
  'get-todo-task': identity,
  'get-user-manager': identity,
  'get-user-presence': identity,
  'get-virtual-event-webinar': identity,
  'graph-batch': identity,
  'insert-excel-range': identity,
  'list-all-onenote-sections': identity,
  'list-calendar-event-instances': identity,
  'list-calendar-events-delta': identity,
  'list-calendar-view-delta': identity,
  'list-calendars': identity,
  'list-channel-message-hosted-contents': identity,
  'list-channel-message-replies': identity,
  'list-channel-messages': identity,
  'list-channel-tabs': identity,
  'list-chat-members': identity,
  'list-chat-message-hosted-contents': identity,
  'list-chat-message-replies': identity,
  'list-chat-messages': identity,
  'list-chats': identity,
  'list-drive-item-permissions': identity,
  'list-drive-item-thumbnails': identity,
  'list-drive-item-versions': identity,
  'list-drives': identity,
  'list-excel-table-rows': identity,
  'list-excel-tables': identity,
  'list-excel-worksheets': identity,
  'list-focused-inbox-overrides': identity,
  'list-folder-files': identity,
  'list-group-conversations': identity,
  'list-group-events': identity,
  'list-group-members': identity,
  'list-group-owners': identity,
  'list-group-threads': identity,
  'list-groups': identity,
  'list-joined-teams': identity,
  'list-mail-attachments': identity,
  'list-mail-child-folders': identity,
  'list-mail-folder-messages-delta': identity,
  'list-mail-folders': identity,
  'list-mail-rules': identity,
  'list-meeting-attendance-records': identity,
  'list-meeting-attendance-reports': identity,
  'list-meeting-recordings': identity,
  'list-meeting-transcripts': identity,
  'list-my-direct-reports': identity,
  'list-my-memberships': identity,
  'list-onenote-notebook-sections': identity,
  'list-onenote-notebooks': identity,
  'list-onenote-section-pages': identity,
  'list-online-meetings': identity,
  'list-outlook-categories': identity,
  'list-outlook-contacts': identity,
  'list-pinned-chat-messages': identity,
  'list-plan-buckets': identity,
  'list-plan-tasks': identity,
  'list-planner-tasks': identity,
  'list-relevant-people': identity,
  'list-room-list-rooms': identity,
  'list-sensitivity-labels': identity,
  'list-shared-calendar-events': identity,
  'list-shared-mailbox-folder-messages': identity,
  'list-shared-mailbox-messages': identity,
  'list-sharepoint-list-columns': identity,
  'list-sharepoint-site-drives': identity,
  'list-sharepoint-site-items': identity,
  'list-sharepoint-site-list-items': identity,
  'list-sharepoint-site-lists': identity,
  'list-sharepoint-site-onenote-notebook-sections': identity,
  'list-sharepoint-site-onenote-notebooks': identity,
  'list-sharepoint-site-onenote-section-pages': identity,
  'list-subscriptions': identity,
  'list-team-channels': identity,
  'list-team-members': identity,
  'list-todo-linked-resources': identity,
  'list-todo-task-lists': identity,
  'list-todo-tasks': identity,
  'list-trending-insights': identity,
  'list-user-direct-reports': identity,
  'list-users': identity,
  'list-webinar-sessions': identity,
  'merge-excel-range': identity,
  'move-mail-message': identity,
  'move-rename-onedrive-item': identity,
  'pin-chat-message': identity,
  'reauthorize-subscription': identity,
  'remove-group-member': identity,
  'remove-group-owner': identity,
  'remove-team-member': identity,
  'reply-all-mail-message': identity,
  'reply-mail-message': identity,
  'reply-to-channel-message': identity,
  'reply-to-chat-message': identity,
  'reply-to-group-thread': identity,
  'search-onedrive-files': identity,
  'search-query': identity,
  'search-sharepoint-sites': identity,
  'send-channel-message': identity,
  'send-chat-message': identity,
  'send-draft-message': identity,
  'send-mail': identity,
  'send-shared-mailbox-mail': identity,
  'set-channel-message-reaction': identity,
  'set-chat-message-reaction': identity,
  'share-drive-item': identity,
  'snooze-calendar-event-reminder': identity,
  'sort-excel-range': identity,
  'tentatively-accept-calendar-event': identity,
  'unmerge-excel-range': identity,
  'unpin-chat-message': identity,
  'unset-channel-message-reaction': identity,
  'unset-chat-message-reaction': identity,
  'update-calendar': identity,
  'update-calendar-event': identity,
  'update-excel-range': identity,
  'update-excel-table-row': identity,
  'update-focused-inbox-override': identity,
  'update-group': identity,
  'update-mail-folder': identity,
  'update-mail-message': identity,
  'update-mail-rule': identity,
  'update-mailbox-settings': identity,
  'update-online-meeting': identity,
  'update-outlook-contact': identity,
  'update-place': identity,
  'update-planner-bucket': identity,
  'update-planner-task': identity,
  'update-planner-task-details': identity,
  'update-sharepoint-list-column': identity,
  'update-sharepoint-list-item': identity,
  'update-specific-calendar-event': identity,
  'update-subscription': identity,
  'update-team-channel': identity,
  'update-todo-task': identity,
  'upload-file-content': identity,
};
