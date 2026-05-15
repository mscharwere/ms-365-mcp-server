/**
 * M365 MCP Compactor — Phase 2
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
 *             All Phase 2 tools use the same projectCalendarEvent / projectMailMessage helpers
 *             as Phase 1 — the projection logic is already validated.
 *
 * 6. STRUCTURAL SURPRISES (Phase 2 — from FRIDAY's real fixture captures):
 *
 *    6a. ZWSP PREHEADER (stripZwspPreheader).
 *        SoFi/Chase/Discover send 127–150 U+200C ZERO WIDTH NON-JOINER chars at body start
 *        as a preheader-suppression technique. bodyPreview appears blank to humans.
 *        We strip the leading ZWSP block from body.content before safelinks decoding.
 *        The ZWSP block is pure whitespace noise — no information content.
 *
 *    6b. LEGACY TIMEZONE (normalizeTimezone).
 *        Wix/iCal sends originalStartTimeZone = "tzone://Microsoft/Utc" (non-IANA).
 *        This field is metadata-only (not used in start/end dateTime parsing by consumers).
 *        We map known legacy forms to IANA equivalents; unknown forms pass through as-is.
 *        We do NOT touch start.timeZone / end.timeZone — those come from Graph's own conversion.
 *
 *    6c. SENTINEL DATE (isSentinelDate).
 *        recurrence.range.endDate = "0001-01-01" means "no end" in Outlook's recurrence model.
 *        We preserve it verbatim — do NOT parse as a real date. Compactors that preserve the
 *        recurrence object intact automatically preserve this value.
 *
 *    6d. EMPTY BODY DISTINCTION.
 *        Calendar all-day seriesMaster: body.content === "" (truly empty string).
 *        Mail (and non-all-day calendar): body.content === "\r\n" (near-empty CRLF).
 *        The compactor preserves both as-is. DO NOT normalize "" to "\r\n" or vice versa.
 *        The distinction matters for downstream consumers testing `content === ""`.
 *
 *    6e. TEMPLATE VARIABLE PRESERVATION.
 *        Wix Automations sends unresolved template tokens like `${staff_member_n` followed by
 *        ZWSP padding (token was truncated before variable was resolved). Do NOT try to parse
 *        or evaluate `${...}` expressions. Pass through as-is in body.content.
 *
 *    6f. RECURRENCE ON SERIES MASTERS.
 *        The recurrence object is the source-of-truth for when occurrences happen.
 *        On seriesMaster events: preserve recurrence intact (pattern + range verbatim).
 *        On instance/occurrence events: recurrence is null — drop it (already the Phase 1 behavior).
 *        Sentinel endDate "0001-01-01" inside range is preserved verbatim (see 6c above).
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
// ZWSP preheader stripper (Phase 2 — structural surprise 6a)
// ---------------------------------------------------------------------------

/**
 * Regex matching a leading block of mixed whitespace + U+200C ZERO WIDTH NON-JOINER characters
 * at the start of a string. The lookahead requires the run to contain at least 20 ZWSPs to avoid
 * false-positives on legitimate leading whitespace in short bodies.
 *
 * Context: SoFi, Chase, Discover and other financial senders use 127–150 ZWSPs as a preheader
 * suppressor — flooding the email client's ~140-char bodyPreview slot so the real content
 * isn't shown as preview text. The ZWSP block carries no information.
 *
 * Phase 3 Step 2 widening: real production bodies start with a SPACE then alternate
 * SPACE+ZWSP+SPACE+ZWSP for ~290 chars (143 ZWSPs interleaved with 149 spaces in the SoFi
 * fixture). The original regex `^[‌]{20,}\s*` only matched a CONTIGUOUS leading run of ZWSPs
 * and missed the real interleaved shape. The widened pattern accepts any whitespace-or-ZWSP
 * leading run, gated by a lookahead that counts ≥20 ZWSPs in that run — keeps the
 * preheader-signature semantics, drops the contiguous-only assumption.
 */
const ZWSP_PREHEADER_RE = /^(?=(?:[\s‌]*‌){20,})[\s‌]+/;

/**
 * Strip the ZWSP preheader from a mail body string.
 * Returns the body with the leading ZWSP block removed.
 * If no ZWSP block is present, returns the original string unchanged.
 *
 * Operates on body.content only — do NOT apply to bodyPreview (bodyPreview is already
 * truncated server-side and its ZWSP content is what consumers use to detect blank preview).
 */
export function stripZwspPreheader(body: string): string {
  return body.replace(ZWSP_PREHEADER_RE, '');
}

// ---------------------------------------------------------------------------
// Legacy timezone normalizer (Phase 2 — structural surprise 6b)
// ---------------------------------------------------------------------------

/**
 * Known legacy-to-IANA timezone mappings.
 *
 * "tzone://Microsoft/Utc" is the legacy format emitted by Wix Automations and some iCal
 * generators. It appears in originalStartTimeZone / originalEndTimeZone fields (metadata only —
 * NOT in start.timeZone / end.timeZone which carry Graph's converted value).
 *
 * Choice: map to IANA equivalent when unambiguous; return as-is for unknown forms.
 * Rationale: returning as-is for unknowns is safer than silently dropping or guessing.
 * The field is metadata-only; consumers reading it for display should handle gracefully.
 */
const LEGACY_TZ_MAP: Record<string, string> = {
  'tzone://Microsoft/Utc': 'UTC',
  'tzone://Microsoft/Pacific Standard Time': 'America/Los_Angeles',
  'tzone://Microsoft/Mountain Standard Time': 'America/Denver',
  'tzone://Microsoft/Central Standard Time': 'America/Chicago',
  'tzone://Microsoft/Eastern Standard Time': 'America/New_York',
};

/**
 * Normalize a timezone string from legacy Microsoft/iCal format to IANA where known.
 * Returns the IANA equivalent if mapped, or the original string if not recognized.
 * Returns null if input is null/undefined.
 *
 * IMPORTANT: Only call this on originalStartTimeZone / originalEndTimeZone fields.
 * NEVER apply to start.timeZone / end.timeZone — those are Graph's own output.
 */
export function normalizeTimezone(tz: string | null | undefined): string | null {
  if (tz == null) return null;
  return LEGACY_TZ_MAP[tz] ?? tz;
}

// ---------------------------------------------------------------------------
// Sentinel date detector (Phase 2 — structural surprise 6c)
// ---------------------------------------------------------------------------

/**
 * Detect Outlook's "no end" sentinel date in recurrence ranges.
 *
 * recurrence.range.endDate = "0001-01-01" means the recurrence has no end date.
 * This is NOT a real date — it's a sentinel value used when range.type = "noEnd".
 * Do NOT parse it as a date, do NOT compute durations from it, do NOT use it in comparisons.
 *
 * The compactor preserves it verbatim in the recurrence object.
 * This function is provided for consumers that want to explicitly check before date math.
 */
export function isSentinelDate(d: string): boolean {
  return d === '0001-01-01';
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
 *   attendees (projected to { type, status: { response }, emailAddress: { name, address } }),
 *   recurrence (on seriesMaster events — source-of-truth for occurrence expansion)
 *
 * DROPPED:
 *   internetMessageId, iCalUId, webLink — not read by consumers
 *   recurrence when null (instances; already expanded in view)
 *   organizer (not needed for briefing routing)
 *   location.locationType, location.uniqueId, location.uniqueIdType,
 *     location.address, location.coordinates, location.locationUri — only displayName used
 *   body.contentType — always "text" in our setup; redundant
 *   attendees[].status.time — always "0001-01-01T00:00:00Z" sentinel; safe to drop
 *   originalStartTimeZone / originalEndTimeZone — legacy metadata, not used for display
 *
 * EMPTY BODY DISTINCTION (structural surprise 6d):
 *   body.content === "" (truly empty) is preserved as "" — only on all-day seriesMaster events.
 *   body.content === "\r\n" (near-empty CRLF) is preserved as "\r\n" — on most other events.
 *   DO NOT normalize one to the other.
 *
 * RECURRENCE ON SERIES MASTERS (structural surprise 6f):
 *   If event.recurrence is non-null, preserve the full recurrence object verbatim.
 *   recurrence.range.endDate = "0001-01-01" is a sentinel (isSentinelDate) — preserved as-is.
 *   On instance events, recurrence is null — we skip it (no key emitted).
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

  // Body — strip ZWSP preheader, decode safelinks, keep content only.
  // EMPTY BODY: preserve body.content as-is including "" (truly empty, all-day seriesMaster)
  // vs "\r\n" (near-empty CRLF, most other events). DO NOT normalize.
  // TEMPLATE VARIABLES: ${...} tokens pass through untouched (structural surprise 6e).
  // NOTE: Calendar events CAN have safelinks in body (e.g. Wix/iCal cancellation emails
  // sent as calendar invites — the body.content is a full email body from the booking system).
  if (event.body !== undefined) {
    const raw = event.body.content !== undefined ? event.body.content : '';
    // Apply safelinks decoding to non-empty bodies only.
    // For truly empty bodies (all-day seriesMaster, content === ""), skip to preserve the exact "" sentinel.
    // stripZwspPreheader on "" is a no-op but keeping the guard makes intent clear.
    const content = raw === '' ? '' : decodeSafelinks(stripZwspPreheader(raw));
    out.body = { content };
  }

  // Recurrence — preserve intact on seriesMaster events (non-null recurrence).
  // On instance events, recurrence is null — skip entirely (no key emitted).
  // recurrence.range.endDate "0001-01-01" sentinel is preserved verbatim inside the object.
  if (event.recurrence !== null && event.recurrence !== undefined) {
    out.recurrence = event.recurrence;
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

  // Body — strip ZWSP preheader, then decode safelinks, keep content only.
  // Order matters: strip ZWSP first so the safelinks regex doesn't have to match through
  // 150 zero-width chars at string start. Template variables (${...}) pass through untouched.
  if (msg.body !== undefined) {
    const raw = msg.body.content ?? '';
    out.body = {
      content: decodeSafelinks(stripZwspPreheader(raw)),
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
// Contact projection helpers
// ---------------------------------------------------------------------------

/**
 * Project a single Outlook contact (/me/contacts).
 *
 * PRESERVED:
 *   id, displayName, givenName, surname, middleName, nickName,
 *   emailAddresses (name + address only — drop @odata bag),
 *   businessPhones, homePhones, mobilePhone,
 *   homeAddress, businessAddress, otherAddress (street/city/state/postalCode/countryOrRegion only),
 *   companyName, department, jobTitle, officeLocation,
 *   personalNotes, birthday, fileAs, categories
 *
 * DROPPED (Graph metadata + rarely useful fields):
 *   @odata.context, @odata.etag, createdDateTime, lastModifiedDateTime, changeKey,
 *   parentFolderId, imAddresses, yomiGivenName, yomiSurname, yomiCompanyName,
 *   title (honorific — almost always null), initials, generation,
 *   profession, manager, assistantName, spouseName, children,
 *   business/home/other fax numbers
 */
function projectContact(c: JsonValue): JsonValue {
  if (!c || typeof c !== 'object') return c;

  const out: JsonValue = {};

  // Identity
  if (c.id !== undefined) out.id = c.id;

  // Names
  if (c.displayName !== undefined) out.displayName = c.displayName;
  if (c.givenName !== undefined) out.givenName = c.givenName;
  if (c.surname !== undefined) out.surname = c.surname;
  if (c.middleName !== undefined) out.middleName = c.middleName;
  if (c.nickName !== undefined) out.nickName = c.nickName;
  if (c.fileAs !== undefined) out.fileAs = c.fileAs;

  // Email — strip Graph @odata bag, keep name+address only
  if (c.emailAddresses !== undefined && Array.isArray(c.emailAddresses)) {
    out.emailAddresses = c.emailAddresses.map((e: JsonValue) => ({
      name: e?.name,
      address: e?.address,
    }));
  }

  // Phones — preserve arrays as-is (already plain strings)
  if (c.businessPhones !== undefined) out.businessPhones = c.businessPhones;
  if (c.homePhones !== undefined) out.homePhones = c.homePhones;
  if (c.mobilePhone !== undefined) out.mobilePhone = c.mobilePhone;

  // Addresses — project to a minimal shape
  const projectAddress = (a: JsonValue): JsonValue => {
    if (!a || typeof a !== 'object') return a;
    const o: JsonValue = {};
    if (a.street !== undefined) o.street = a.street;
    if (a.city !== undefined) o.city = a.city;
    if (a.state !== undefined) o.state = a.state;
    if (a.postalCode !== undefined) o.postalCode = a.postalCode;
    if (a.countryOrRegion !== undefined) o.countryOrRegion = a.countryOrRegion;
    return o;
  };
  if (c.homeAddress !== undefined) out.homeAddress = projectAddress(c.homeAddress);
  if (c.businessAddress !== undefined) out.businessAddress = projectAddress(c.businessAddress);
  if (c.otherAddress !== undefined) out.otherAddress = projectAddress(c.otherAddress);

  // Employer / role
  if (c.companyName !== undefined) out.companyName = c.companyName;
  if (c.department !== undefined) out.department = c.department;
  if (c.jobTitle !== undefined) out.jobTitle = c.jobTitle;
  if (c.officeLocation !== undefined) out.officeLocation = c.officeLocation;

  // Notes / misc
  if (c.personalNotes !== undefined) out.personalNotes = c.personalNotes;
  if (c.birthday !== undefined) out.birthday = c.birthday;
  if (c.categories !== undefined) out.categories = c.categories;

  return out;
}

/**
 * Project a single /me/people item ("relevant person").
 *
 * Different shape from /me/contacts: scoredEmailAddresses, phones with type tags,
 * personType (Person/Group), relevanceScore.
 *
 * PRESERVED:
 *   id, displayName, givenName, surname,
 *   scoredEmailAddresses (address + relevanceScore — drop selectionLikelihood),
 *   phones (number + type only),
 *   companyName, jobTitle, department, officeLocation,
 *   personType (class + subclass for disambiguating Person vs Group)
 *
 * DROPPED:
 *   @odata.* metadata, imAddress, profession, userPrincipalName, yomi* fields,
 *   birthday (almost always empty on /me/people), postalAddresses (rarely populated),
 *   websites, isFavorite (deprecated field)
 */
function projectPerson(p: JsonValue): JsonValue {
  if (!p || typeof p !== 'object') return p;

  const out: JsonValue = {};

  if (p.id !== undefined) out.id = p.id;
  if (p.displayName !== undefined) out.displayName = p.displayName;
  if (p.givenName !== undefined) out.givenName = p.givenName;
  if (p.surname !== undefined) out.surname = p.surname;

  // Email — drop selectionLikelihood (Graph's heuristic flag, low signal)
  if (p.scoredEmailAddresses !== undefined && Array.isArray(p.scoredEmailAddresses)) {
    out.scoredEmailAddresses = p.scoredEmailAddresses.map((e: JsonValue) => ({
      address: e?.address,
      relevanceScore: e?.relevanceScore,
    }));
  }

  // Phones — keep number + type only
  if (p.phones !== undefined && Array.isArray(p.phones)) {
    out.phones = p.phones.map((ph: JsonValue) => ({
      type: ph?.type,
      number: ph?.number,
    }));
  }

  // Employer / role
  if (p.companyName !== undefined) out.companyName = p.companyName;
  if (p.jobTitle !== undefined) out.jobTitle = p.jobTitle;
  if (p.department !== undefined) out.department = p.department;
  if (p.officeLocation !== undefined) out.officeLocation = p.officeLocation;

  // Person vs Group disambiguation
  if (p.personType !== undefined) {
    out.personType = {
      class: p.personType?.class,
      subclass: p.personType?.subclass,
    };
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
// Per-tool compactors (Phase 1: 3 explicit; Phase 2: 5 more; identity for remaining 262)
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
// Phase 2 compactors — same projection logic as Phase 1; now explicitly wired
// ---------------------------------------------------------------------------

/**
 * get-calendar-view — returns { value: Event[] } for /me/calendarView (default calendar).
 *
 * This is the broader calendar scan that FRIDAY uses when not targeting a specific calendar.
 * Projection identical to list-specific-calendar-events.
 */
const compactGetCalendarView: Compactor = (raw) => compactList(raw, projectCalendarEvent);

/**
 * get-specific-calendar-view — returns { value: Event[] } for /me/calendars/{id}/calendarView.
 *
 * Phase 1 confirmed: this is the dual-pull tool (Personal + Family IDs) per
 * feedback_m365_calendar_multi_pull.md. Phase 2 promotes from identity to explicit projection.
 */
const compactGetSpecificCalendarView: Compactor = (raw) => compactList(raw, projectCalendarEvent);

/**
 * list-calendar-events — returns { value: Event[] } for /me/events.
 *
 * General event listing without time-window scoping (unlike calendarView).
 * Projection identical to other calendar list tools.
 */
const compactListCalendarEvents: Compactor = (raw) => compactList(raw, projectCalendarEvent);

/**
 * get-mail-message — returns a single Message object for /me/messages/{id}.
 *
 * Full body is present (not just bodyPreview). Safelinks decoding + ZWSP stripping apply.
 * Template variables in body.content pass through untouched.
 */
const compactGetMailMessage: Compactor = (raw) => projectMailMessage(raw);

/**
 * list-mail-messages — returns { value: Message[], @odata.nextLink? } for /me/messages.
 *
 * NOTE: This tool has a stale 10-msg cache bug on Carlos's account (see feedback_m365_list_mail_messages_stale_cache.md).
 * FRIDAY uses list-mail-folder-messages instead. We compact this anyway since the projection
 * is identical — if the bug is ever fixed, compaction works correctly.
 */
const compactListMailMessages: Compactor = (raw) => compactList(raw, projectMailMessage);

// ---------------------------------------------------------------------------
// Contacts + People compactors (Phase 3 Step 2 — added 2026-05-15)
//
// Existing Graph endpoints already supported via identity; promoted here because Graph
// returns 25–40 metadata fields per contact (changeKey, parentFolderId, yomi*, fax numbers,
// imAddresses, profession, manager, assistantName, spouseName, generation, initials, etc.).
// Typical contact object: ~2.2KB raw → ~0.6KB compacted (~70% reduction).
//
// /me/people uses a different shape (scoredEmailAddresses, personType, phones with type tags)
// so it gets its own projector.
// ---------------------------------------------------------------------------

const compactListOutlookContacts: Compactor = (raw) => compactList(raw, projectContact);
const compactGetOutlookContact: Compactor = (raw) => projectContact(raw);
const compactListRelevantPeople: Compactor = (raw) => compactList(raw, projectPerson);

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
  // Phase 2 — explicit projections (promoted from identity)
  // -------------------------------------------------------------------------

  /** Broad calendar scan — default calendar time-window view */
  'get-calendar-view': compactGetCalendarView,

  /** Specific-calendar time-window view (dual-pull tool per multi-pull memory rule) */
  'get-specific-calendar-view': compactGetSpecificCalendarView,

  /** General event listing without time-window scoping */
  'list-calendar-events': compactListCalendarEvents,

  /** Single mail message with full body — safelinks decoded, ZWSP stripped */
  'get-mail-message': compactGetMailMessage,

  /** Mail message list — same projection as list-mail-folder-messages */
  'list-mail-messages': compactListMailMessages,

  /** Single calendar event fetch (default calendar path) — same projection as get-specific-calendar-event */
  'get-calendar-event': compactGetSpecificCalendarEvent,

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
  'get-outlook-contact': compactGetOutlookContact,
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
  'list-outlook-contacts': compactListOutlookContacts,
  'list-pinned-chat-messages': identity,
  'list-plan-buckets': identity,
  'list-plan-tasks': identity,
  'list-planner-tasks': identity,
  'list-relevant-people': compactListRelevantPeople,
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
