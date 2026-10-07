/*
 * Room booking: bookable rooms, booking rules, availability, attendee search and creating the event.
 * Ported from the legacy Outlook add-in (PlaceOS user-interfaces apps/outlook-addin rooms flow and
 * libs/events EventFormService.postForm). PlaceOS creates the calendar event through the Staff API.
 */

import { getMetadata, type AppSettings, type Organisation } from "./booking-settings";
import { allPages, resultList, type PlaceosApi } from "./placeos-data";

export type Space = {
  id: string;
  name: string;
  display_name?: string;
  email?: string;
  capacity?: number;
  features?: string[];
  feature_list?: string[];
  bookable?: boolean;
  room_booking_url?: string;
  zones?: string[];
  images?: string[];
  description?: string;
  tags?: string[];
  // Id of the room's element on its level's floor plan.
  map_id?: string;
};

export type Person = {
  name: string;
  email: string;
  department?: string;
  organisation?: string;
  external?: boolean;
};

export type RecurrencePattern = "none" | "daily" | "weekly" | "monthly";

export type Recurrence = {
  pattern: RecurrencePattern;
  interval: number;
  // 0 = Sunday. Weekly only.
  daysOfWeek: number[];
  // Epoch ms of the last day the series may run on.
  until: number;
};

/** The form's values. Times are epoch ms; durations are minutes. */
export type BookingRequest = {
  title: string;
  start: number;
  duration: number;
  attendees: Person[];
  notes: string;
  recurrence: Recurrence | null;
  // Set by the All day checkbox; start and duration already hold the all-day period.
  allDay: boolean;
  room: Space;
  buildingId: string;
};

/** The meeting the user is composing in Outlook. Times are epoch ms. */
export type DraftDetails = {
  title: string;
  start: number;
  end: number;
  attendees: Person[];
  // The draft repeats (only known on clients with Mailbox 1.7).
  recurring: boolean;
};

/**
 * A meeting being composed in the host (outlook-draft.ts). When set, Book a room fills in the draft
 * instead of PlaceOS creating an event; the room is booked when the user sends the invitation.
 */
export interface MeetingDraft {
  read(): Promise<DraftDetails>;
  fill(request: BookingRequest): Promise<void>;
}

export const spaceName = (space: { name: string; display_name?: string }) =>
  space.display_name || space.name;

export const spaceFeatures = (space: Space) =>
  space.feature_list?.length ? space.feature_list : space.features || [];

const unix = (ms: number) => Math.floor(ms / 1000);

/** Bookable rooms in a zone (building, or region when app.use_region is set). */
export async function getSpaces(api: PlaceosApi, zoneId: string): Promise<Space[]> {
  const result = await api<Space[] | { results?: Space[] }>(
    `/api/engine/v2/systems?zone_id=${encodeURIComponent(zoneId)}&limit=500&signage=false`,
    allPages
  );
  return resultList(result).filter(
    (space) => space.bookable && space.email && !space.room_booking_url
  );
}

// ---------------------------------------------------------------------------------------------------------
// Booking rules: metadata "room_booking_rules" (or "desk_booking_rules") on each building
// (libs/common booking-rules.ts).
// ---------------------------------------------------------------------------------------------------------

export type BookingRuleset = {
  zone: string;
  rules?: { auto_approve?: boolean; hidden?: boolean };
  conditions?: {
    groups?: string[];
    locations?: string[];
    min_length?: number;
    max_length?: number;
    is_before?: string;
    is_after?: string;
    is_between?: [number, number];
    is_period?: [number, number];
    resource_ids?: string[];
    tags?: string[];
  };
};

const durationUnits: Record<string, number> = {
  minute: 1,
  minutes: 1,
  hour: 60,
  hours: 60,
  day: 1440,
  days: 1440,
  week: 10080,
  weeks: 10080,
  month: 43200,
  months: 43200,
};

/** "2 weeks" → minutes. */
export function stringToMinutes(text: string): number {
  const [amount, unit] = (text || "").split(" ");
  return unit ? +amount * (durationUnits[unit.toLowerCase()] || 0) : 0;
}

/** What booking rules are matched against: a room, or a desk with its level and building as zones. */
export type RuleResource = Pick<Space, "id" | "name" | "zones" | "tags">;

export function getBookingRules(
  api: PlaceosApi,
  buildingId: string,
  name = "room_booking_rules"
): Promise<BookingRuleset[]> {
  return getMetadata<BookingRuleset[]>(api, buildingId, name).then((rules) =>
    Array.isArray(rules) ? rules : []
  );
}

function rulesMatch(
  ruleset: BookingRuleset,
  space: RuleResource,
  start: number,
  duration: number,
  groups: string[]
): boolean {
  const c = ruleset.conditions;
  if (!c) {
    return true;
  }
  const end = start + duration * 60000;
  const fromNow = (text: string) => Date.now() + stringToMinutes(text) * 60000;
  const date = new Date(start);
  const hour = date.getHours() + date.getMinutes() / 60;
  // Same counting as Workplace: every listed condition has to match.
  let matches = 0;
  if (Array.isArray(c.groups) && c.groups.every((g) => groups.includes(g))) matches++;
  if (c.is_before && end < fromNow(c.is_before)) matches++;
  if (c.is_after && start > fromNow(c.is_after)) matches++;
  if (c.min_length && c.min_length <= duration) matches++;
  if (c.is_between && hour >= c.is_between[0] && hour < c.is_between[1]) matches++;
  if (c.is_period && start >= c.is_period[0] && start < c.is_period[1]) matches++;
  if (c.max_length && c.max_length >= duration) matches++;
  if (c.resource_ids && c.resource_ids.includes(space.id)) matches++;
  if (c.tags && c.tags.every((tag) => (space.tags || []).includes(tag))) matches++;
  if (c.locations && c.locations.includes(space.name)) matches++;
  return matches >= Object.keys(c).length;
}

/** Whether the first matching ruleset hides the room (or desk) for this booking. */
export function hiddenByRules(
  rules: BookingRuleset[],
  space: RuleResource,
  start: number,
  duration: number,
  groups: string[]
): boolean {
  for (const ruleset of rules) {
    if (ruleset.zone === "*" || space.zones?.includes(ruleset.zone)) {
      if (rulesMatch(ruleset, space, start, duration, groups)) {
        return !!ruleset.rules?.hidden;
      }
    }
  }
  return false;
}

/** Minutes since midnight of a time in an IANA timezone, or in local time ("" or not supported). */
export function minutesOfDay(ms: number, timezone = ""): number {
  if (timezone) {
    try {
      const parts = new Intl.DateTimeFormat("en-GB", {
        timeZone: timezone,
        hour: "2-digit",
        minute: "2-digit",
        hourCycle: "h23",
      }).formatToParts(new Date(ms));
      const part = (type: string) => +(parts.find((p) => p.type === type)?.value || 0);
      return (part("hour") % 24) * 60 + part("minute");
    } catch {
      // Older webviews only know UTC: fall back to local time.
    }
  }
  return new Date(ms).getHours() * 60 + new Date(ms).getMinutes();
}

/**
 * app.events.bookable_hours (or the hours given): {start, end} in decimal hours, in `timezone` (local time
 * if not given). Returns an error message, or "".
 */
export function checkBookableHours(
  settings: AppSettings,
  start: number,
  duration: number,
  hours = settings.get<{ start?: number; end?: number } | null>("events.bookable_hours", null),
  noun = "Rooms",
  timezone = ""
): string {
  if (!hours || hours.start === undefined || hours.end === undefined) {
    return "";
  }
  const from = minutesOfDay(start, timezone);
  const to = from + duration;
  if (from < hours.start * 60 || from >= hours.end * 60 || to > hours.end * 60) {
    const label = (h: number) =>
      new Date(2000, 0, 1, Math.floor(h), Math.round((h % 1) * 60)).toLocaleTimeString([], {
        hour: "numeric",
        minute: "2-digit",
      });
    return `${noun} can only be booked between ${label(hours.start)} and ${label(hours.end)}.`;
  }
  return "";
}

// ---------------------------------------------------------------------------------------------------------
// All-day bookings (libs/common getAllDayTimeRange): app.events.allow_all_day, app.events.all_day_period
// ---------------------------------------------------------------------------------------------------------

type AllDayPeriodSetting = { start?: number | null; end?: number | null } | null;

export const allDayAllowed = (settings: AppSettings) =>
  settings.get<boolean>("events.allow_all_day") === true;

/**
 * The period an all-day booking on `day` (local midnight, ms) covers: the whole day, or the hours in
 * app.events.all_day_period, or `period` if given ({start, end} in decimal hours). Duration is in minutes.
 * As Workplace's getAllDayTimeRange, today's period doesn't start from now; callers check if it has ended.
 */
export function allDayPeriod(
  settings: AppSettings,
  day: number,
  period = settings.get<AllDayPeriodSetting>("events.all_day_period", null)
): { start: number; duration: number } {
  const date = new Date(day);
  const at = (hours: number) =>
    new Date(
      date.getFullYear(),
      date.getMonth(),
      date.getDate(),
      0,
      Math.round(hours * 60)
    ).getTime();
  let start = at(0);
  let end = at(24);
  if (typeof period?.start === "number" && typeof period?.end === "number") {
    const from = Math.max(0, Math.min(23, period.start));
    start = at(from);
    end = at(Math.max(from + 1, Math.min(24, period.end)));
  }
  return { start, duration: Math.round((end - start) / 60000) };
}

/** Whether a period runs from local midnight to the next midnight, i.e. a calendar all-day event. */
export function isFullDay(start: number, duration: number): boolean {
  const date = new Date(start);
  const midnight = new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
  const next = new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1).getTime();
  return start === midnight && start + duration * 60000 === next;
}

// ---------------------------------------------------------------------------------------------------------
// Availability
// ---------------------------------------------------------------------------------------------------------

type AvailableCalendar = { id?: string; resource?: { id?: string }; system?: { id?: string } };

/**
 * Ids of rooms free for the whole period. Staff API returns calendars for the available rooms only.
 * With app.events.use_bookings, rooms are booked as PlaceOS bookings, so clashes come from those instead.
 */
export async function getAvailableSpaceIds(
  api: PlaceosApi,
  spaces: Space[],
  start: number,
  duration: number,
  settings: AppSettings
): Promise<Set<string>> {
  const period = `period_start=${unix(start)}&period_end=${unix(start + duration * 60000)}`;
  const available = new Set<string>();
  if (settings.get<boolean>("events.use_bookings") === true) {
    const bookings = await api<{ asset_id?: string; asset_ids?: string[] }[]>(
      `/api/staff/v1/bookings?type=room&${period}`
    );
    const busy = new Set(bookings.flatMap((b) => [b.asset_id, ...(b.asset_ids || [])]));
    spaces.filter((s) => !busy.has(s.id)).forEach((s) => available.add(s.id));
    return available;
  }
  const chunkSize = 50;
  for (let i = 0; i < spaces.length; i += chunkSize) {
    const ids = spaces.slice(i, i + chunkSize).map((s) => s.id);
    const calendars = await api<AvailableCalendar[]>(
      `/api/staff/v1/calendars/availability?system_ids=${ids.map(encodeURIComponent).join(",")}&${period}`
    );
    for (const calendar of calendars || []) {
      for (const id of [calendar.id, calendar.resource?.id, calendar.system?.id]) {
        if (id && ids.includes(id)) {
          available.add(id);
        }
      }
    }
  }
  return available;
}

// ---------------------------------------------------------------------------------------------------------
// Attendees
// ---------------------------------------------------------------------------------------------------------

export async function searchPeople(
  api: PlaceosApi,
  query: string,
  settings: AppSettings
): Promise<Person[]> {
  const q = encodeURIComponent(query);
  if (settings.get<boolean>("basic_user_search")) {
    const users = await api<Person[]>(`/api/engine/v2/users?q=${q}&limit=20`);
    return users.filter((user) => user.email);
  }
  const people = await api<Person[]>(
    `/api/staff/v1/people?q=${q}&fields=id,name,email,username,organisation,department`
  );
  return (people || []).filter((person) => person.email).slice(0, 20);
}

export const isEmail = (text: string) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(text.trim());

// ---------------------------------------------------------------------------------------------------------
// Recurrence (libs/common parseRecurrence). The API wants weekday names and whole local days.
// ---------------------------------------------------------------------------------------------------------

const weekdayNames = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];

/**
 * Which occurrence of its weekday a date is in the month: 1–4, or -1 (last) for days 29–31.
 * Matches how PlaceOS/calendar `Recurrence.week_of_month` books "monthly" series.
 */
export function weekOfMonth(date: Date): number {
  const week = Math.floor((date.getDate() - 1) / 7) + 1;
  return week === 5 ? -1 : week;
}

/**
 * True when a "monthly" series can't start on this date. Staff API ignores `nth_of_month` and
 * PlaceOS/calendar works out the week as `day // 7`, which books the 7th, 14th, 21st and 28th a
 * week late (fix pending on PlaceOS/calendar branch fix/monthly-recurrence-week).
 */
export const monthlyStartUnsupported = (date: Date) => date.getDate() % 7 === 0;

export const monthlyStartUnsupportedMessage =
  "Monthly repeats can't start on the 7th, 14th, 21st or 28th yet. Choose another date or repeat weekly.";

function recurrenceBody(recurrence: Recurrence, start: number) {
  const startOfDay = new Date(start);
  startOfDay.setHours(0, 0, 0, 0);
  const endOfDay = new Date(recurrence.until);
  endOfDay.setHours(23, 59, 59, 0);
  const date = new Date(start);
  return {
    range_start: unix(startOfDay.getTime()),
    range_end: unix(endOfDay.getTime()),
    interval: recurrence.interval,
    pattern: recurrence.pattern,
    days_of_week: (recurrence.pattern === "weekly"
      ? recurrence.daysOfWeek
      : recurrence.pattern === "monthly"
        ? [date.getDay()]
        : []
    ).map((day) => weekdayNames[day]),
    nth_of_month: recurrence.pattern === "monthly" ? weekOfMonth(date) : undefined,
  };
}

// ---------------------------------------------------------------------------------------------------------
// Creating the event
// ---------------------------------------------------------------------------------------------------------

export type CreatedEvent = {
  id: string;
  ical_uid?: string;
  event_start: number;
  event_end: number;
  title?: string;
  host?: string;
  system?: { id: string; zones?: string[] };
};

/** Setup and breakdown (minutes): the largest of app.events.setup/breakdown and the room's overflow. */
function overflow(settings: AppSettings, space: Space) {
  const room = settings.get<{ setup?: number; breakdown?: number }>(
    `events.overflow.${space.id}`,
    {}
  );
  return {
    setup: Math.max(settings.get<number>("events.setup", 0) || 0, room.setup || 0),
    breakdown: Math.max(settings.get<number>("events.breakdown", 0) || 0, room.breakdown || 0),
  };
}

/** Host: app.events.force_host, else the room with app.events.room_as_host, else the signed-in user. */
function hostFor(settings: AppSettings, org: Organisation, space: Space): string {
  return (
    settings.get<string>("events.force_host") ||
    (settings.get<boolean>("events.room_as_host") ? space.email || "" : "") ||
    org.user.email
  );
}

/** The Staff API event body, built like libs/common CalendarEvent.toJSON() but only with what's set. */
export function buildEventBody(
  request: BookingRequest,
  org: Organisation,
  settings: AppSettings,
  extras: { catering: unknown[]; assets: unknown[] }
) {
  const { room, start, duration } = request;
  const host = hostFor(settings, org, room);
  const timezone = settings.get<boolean>("events.use_building_timezone")
    ? org.buildings.find((b) => b.id === request.buildingId)?.timezone ||
      Intl.DateTimeFormat().resolvedOptions().timeZone
    : Intl.DateTimeFormat().resolvedOptions().timeZone;
  const organiser: Person = { name: org.user.name || org.user.email, email: org.user.email };
  // As postForm: the organiser and attendees (without visit_expected; visitors are linked bookings), then
  // the room as a resource attendee, unique by email.
  const people: { name: string; email: string; resource?: boolean }[] = [
    organiser,
    ...request.attendees,
    { name: spaceName(room), email: room.email || "", resource: true },
  ].filter(
    (person, index, list) =>
      person.email &&
      list.findIndex((p) => p.email.toLowerCase() === person.email.toLowerCase()) === index
  );
  const { setup, breakdown } = request.allDay
    ? { setup: 0, breakdown: 0 }
    : overflow(settings, room);
  const extension: Record<string, unknown> = {
    department: org.user.department,
    catering: extras.catering,
    assets: extras.assets,
    app_name: "PlaceOS Outlook add-in",
  };
  if (host !== org.user.email) {
    extension.host_override = org.user.email;
  }
  // As CalendarEvent.toJSON(): a whole day is a calendar all-day event; a shorter all-day period
  // (app.events.all_day_period, or the rest of today) is a timed event marked custom_all_day.
  const fullDay = request.allDay && isFullDay(start, duration);
  if (request.allDay) {
    extension.all_day_date = dateValue(start);
    if (!fullDay) {
      extension.custom_all_day = true;
    }
  }
  return {
    event_start: unix(start),
    event_end: unix(start + duration * 60000),
    title: request.title || "Space Booking",
    body: request.notes ? escapeHtml(request.notes).replace(/\n/g, "<br />") : undefined,
    host,
    creator: org.user.email,
    private: false,
    all_day: fullDay,
    timezone,
    attendees: people.map((person) => ({
      name: person.name,
      email: person.email,
      resource: !!person.resource,
    })),
    system_id: room.id,
    location: spaceName(room),
    setup_time: setup,
    breakdown_time: breakdown,
    recurring: !!request.recurrence,
    recurrence: request.recurrence ? recurrenceBody(request.recurrence, start) : undefined,
    extension_data: extension,
  };
}

/** "2026-10-01" in local time. */
export function dateValue(ms: number): string {
  const date = new Date(ms);
  const pad = (n: number) => `${n}`.padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export type Clash = { booking_start: number; booking_end?: number; title?: string };

/**
 * Clashing instances of a recurring series (POST /events/clashing-assets). As Workplace's
 * findEventClashes, a failed check counts as no clashes.
 */
export async function findClashes(api: PlaceosApi, body: unknown): Promise<Clash[]> {
  const clashes = await api<Clash[]>(
    "/api/staff/v1/events/clashing-assets?include_clash_time=true&limit=10000",
    { method: "POST", body: JSON.stringify(body) }
  ).catch(() => [] as Clash[]);
  return (clashes || []).sort((a, b) => a.booking_start - b.booking_start);
}

export function createEvent(api: PlaceosApi, body: unknown): Promise<CreatedEvent> {
  return api<CreatedEvent>("/api/staff/v1/events", { method: "POST", body: JSON.stringify(body) });
}

/** Removes a new event again, e.g. when its catering or asset bookings fail. */
export function removeEvent(
  api: PlaceosApi,
  event: CreatedEvent,
  calendar: string,
  roomId: string
) {
  return api<unknown>(
    `/api/staff/v1/events/${encodeURIComponent(event.id)}?calendar=${encodeURIComponent(calendar)}` +
      `&system_id=${encodeURIComponent(roomId)}`,
    { method: "DELETE" }
  );
}

// ---------------------------------------------------------------------------------------------------------
// Linked bookings (visitors, catering orders, asset requests), created after the event.
// ---------------------------------------------------------------------------------------------------------

export type LinkedBooking = {
  booking_type: string;
  asset_id: string;
  asset_ids: string[];
  asset_name?: string;
  title?: string;
  description?: string;
  attendees?: unknown[];
  // Overrides the event host (asset requests are made for the user, as Workplace).
  user_email?: string;
  extension_data: Record<string, unknown>;
};

/** Zones for linked bookings: the org, region, building and the room's own zones. */
export function bookingZones(org: Organisation, buildingId: string, room: Space): string[] {
  const building = org.buildings.find((b) => b.id === buildingId);
  return [org.org?.id, building?.parent_id, buildingId, ...(room.zones || [])].filter(
    (id, index, list): id is string => !!id && list.indexOf(id) === index
  );
}

/** Creates bookings linked to the event one at a time. If one fails, the ones already made are removed. */
export async function createLinkedBookings(
  api: PlaceosApi,
  event: CreatedEvent,
  common: { user_email: string; zones: string[]; timezone: string },
  bookings: LinkedBooking[]
): Promise<void> {
  const created: string[] = [];
  const query = `ical_uid=${encodeURIComponent(event.ical_uid || "")}&event_id=${encodeURIComponent(event.id)}`;
  try {
    for (const booking of bookings) {
      const result = await api<{ id: string }>(`/api/staff/v1/bookings?${query}`, {
        method: "POST",
        body: JSON.stringify({
          type: booking.booking_type,
          booking_start: event.event_start,
          booking_end: event.event_end,
          title: event.title,
          description: event.title,
          attendees: [],
          ...common,
          ...booking,
          extension_data: {
            parent_id: event.id,
            app_name: "PlaceOS Outlook add-in",
            ...booking.extension_data,
          },
        }),
      });
      if (result?.id) {
        created.push(result.id);
      }
    }
  } catch (error) {
    await Promise.all(
      created.map((id) =>
        api(`/api/staff/v1/bookings/${encodeURIComponent(id)}`, { method: "DELETE" }).catch(
          () => null
        )
      )
    );
    throw error;
  }
}

/** Visitor bookings for external attendees other than the host, so reception expects them (as Workplace). */
export function visitorBookings(
  request: BookingRequest,
  org: Organisation,
  host: string
): LinkedBooking[] {
  const domain = org.user.email.split("@")[1] || "";
  return request.attendees
    .filter(
      (person) =>
        person.external &&
        person.email.toLowerCase() !== host.toLowerCase() &&
        (!domain || !person.email.toLowerCase().includes(domain))
    )
    .map((person) => ({
      booking_type: "visitor",
      asset_id: person.email,
      asset_ids: [person.email],
      asset_name: person.name,
      attendees: [{ name: person.name, email: person.email, visit_expected: true }],
      extension_data: {
        name: person.name,
        location_id: spaceName(request.room),
        details: { name: person.name, email: person.email, organisation: person.organisation },
      },
    }));
}
