/*
 * Today overview data: the signed-in user's rooms, desks, parking and visitors for today.
 * Mirrors what PlaceOS Workplace's schedule page calls (one events query plus one bookings query per type).
 * Desk and parking bookings can be checked in, checked out and cancelled, as in Workplace's booking details.
 */

import type { AppSettings } from "./booking-settings";
import type { PlaceosApi } from "./placeos-data";

/** A button on an item. With confirm, the user is asked that question first. */
export type TodayAction = {
  label: string;
  secondary?: boolean;
  confirm?: string;
  run: () => Promise<unknown>;
};

export type TodayItem = {
  start: number; // epoch seconds, for sorting
  title: string;
  details: string[];
  badge?: { text: string; kind: "ok" | "error" | "" };
  link?: { text: string; url: string };
  actions?: TodayAction[];
  // Where it is on a floor plan: the level's map and the element id of the desk or space on it.
  location?: { mapUrl: string; elementId: string; level: string };
};

/** App settings for the building among these zones, or null if they can't be loaded. */
export type SettingsForZones = (zones: string[]) => Promise<AppSettings | null>;

// map_id: a level's floor plan SVG URL.
type Zone = { id: string; name: string; display_name?: string; tags?: string[]; map_id?: string };

type Attendee = {
  name?: string;
  email: string;
  resource?: boolean;
  organizer?: boolean;
  visit_expected?: boolean;
  checked_in?: boolean;
  organisation?: string;
};

type CalendarEvent = {
  id: string;
  title?: string;
  event_start: number;
  event_end: number;
  all_day?: boolean;
  status?: string;
  private?: boolean;
  location?: string;
  online_meeting_url?: string;
  attendees?: Attendee[];
  system?: { id: string; name: string; display_name?: string } | null;
  extension_data?: { custom_all_day?: boolean };
};

type Booking = {
  id: string;
  booking_start: number;
  booking_end: number;
  all_day?: boolean;
  asset_id?: string;
  description?: string;
  title?: string;
  zones?: string[];
  checked_in?: boolean;
  checked_out_at?: number;
  rejected?: boolean;
  deleted?: boolean;
  status?: string;
  current_state?: string;
  // Start of this occurrence, on bookings in a recurring series.
  instance?: number;
  recurrence_type?: string;
  linked_event?: unknown;
  attendees?: Attendee[];
  extension_data?: {
    assigned_asset_name?: string;
    asset_name?: string;
    name?: string;
    plate_number?: string;
    visitor_name?: string;
    company?: string;
    group_members?: { name?: string; email?: string }[];
    custom_all_day?: boolean;
    map_id?: string;
  };
};

/** Local midnight to the end of today, in epoch seconds (as Workplace sends). */
export function todayPeriod(now = new Date()): { start: number; end: number } {
  const start = new Date(now);
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setHours(23, 59, 59, 0);
  return { start: Math.floor(start.getTime() / 1000), end: Math.floor(end.getTime() / 1000) };
}

const periodQuery = ({ start, end }: { start: number; end: number }) =>
  `period_start=${start}&period_end=${end}`;

/**
 * All-day events and bookings. Workplace and the Book a room form send a shorter all-day period
 * (app.events.all_day_period, or the rest of today) as a timed event with extension_data.custom_all_day.
 */
const isAllDay = (item: { all_day?: boolean; extension_data?: { custom_all_day?: boolean } }) =>
  !!item.all_day || !!item.extension_data?.custom_all_day;

function timeRange(start: number, end: number, allDay?: boolean): string {
  if (allDay) {
    return "All day";
  }
  const format = (seconds: number) =>
    new Date(seconds * 1000).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  return `${format(start)} – ${format(end)}`;
}

/** Time range, prefixed with "Tomorrow" or the weekday when it isn't today. */
function whenText(start: number, end: number, allDay?: boolean): string {
  const range = timeRange(start, end, allDay);
  const day = new Date(start * 1000);
  const today = new Date();
  const tomorrow = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
  if (day.toDateString() === today.toDateString()) {
    return range;
  }
  if (day.toDateString() === tomorrow.toDateString()) {
    return `Tomorrow, ${range}`;
  }
  return `${day.toLocaleDateString(undefined, { weekday: "long", day: "numeric", month: "short" })}, ${range}`;
}

function timingBadge(start: number, end: number): TodayItem["badge"] {
  const now = Date.now() / 1000;
  if (now >= start && now < end) {
    return { text: "Now", kind: "ok" };
  }
  return now >= end ? { text: "Ended", kind: "" } : undefined;
}

/** Status badge from the server-computed state, falling back to the flags. */
function bookingBadge(booking: Booking, checkedInText: string): TodayItem["badge"] {
  const state = booking.current_state;
  if (state === "checked_in" || booking.checked_in) {
    return { text: checkedInText, kind: "ok" };
  }
  if (state === "rejected" || booking.rejected) {
    return { text: "Declined", kind: "error" };
  }
  if (state === "no_show") {
    return { text: "No show", kind: "error" };
  }
  if (state === "checked_out" || state === "ended") {
    return { text: "Ended", kind: "" };
  }
  return timingBadge(booking.booking_start, booking.booking_end);
}

/** Desk or parking space name stored on the booking, as Workplace shows it. */
const storedAssetName = (booking: Booking) =>
  booking.description ||
  booking.extension_data?.assigned_asset_name ||
  booking.extension_data?.asset_name ||
  booking.extension_data?.name ||
  "";

type Asset = { id: string; name?: string; identifier?: string };

/** Caches zone and asset lookups so each level, building, desk or space is fetched once. */
export class PlaceosLookups {
  private cache = new Map<string, Promise<unknown>>();

  constructor(private readonly api: PlaceosApi) {}

  private fetch<T>(path: string): Promise<T | null> {
    if (!this.cache.has(path)) {
      this.cache.set(
        path,
        this.api<T>(path).catch(() => null)
      );
    }
    return this.cache.get(path) as Promise<T | null>;
  }

  private get(id: string): Promise<Zone | null> {
    return this.fetch<Zone>(`/api/engine/v2/zones/${encodeURIComponent(id)}`);
  }

  /**
   * Desk or parking space name: the name stored on the booking, else the PlaceOS Asset's name for
   * asset-based desks/spaces (asset_id "asset-..."), else a legacy (metadata) desk id.
   */
  async assetName(booking: Booking): Promise<string> {
    const stored = storedAssetName(booking);
    if (stored || !booking.asset_id) {
      return stored;
    }
    if (!booking.asset_id.startsWith("asset-")) {
      return booking.asset_id;
    }
    const asset = await this.fetch<Asset>(
      `/api/engine/v2/assets/${encodeURIComponent(booking.asset_id)}`
    );
    return asset?.name || asset?.identifier || "";
  }

  /** "Level 3, Sydney Office" from a booking's zones (usually org, region, building, level). */
  /** The level among a booking's zones. */
  async level(zoneIds: string[] = []): Promise<Zone | null> {
    const zones = await Promise.all(zoneIds.map((id) => this.get(id)));
    return zones.find((zone) => zone?.tags?.includes("level")) || null;
  }

  /**
   * A desk or space on its level's floor plan: the element id is extension_data.map_id, else the asset id,
   * as Workplace's viewLocation. Undefined if the level has no floor plan.
   */
  async mapLocation(booking: Booking): Promise<TodayItem["location"]> {
    const level = await this.level(booking.zones);
    const elementId = booking.extension_data?.map_id || booking.asset_id;
    if (!level?.map_id || !elementId) {
      return undefined;
    }
    return {
      mapUrl: level.map_id,
      elementId,
      level: level.display_name || level.name,
    };
  }

  async location(zoneIds: string[] = []): Promise<string> {
    const zones = (await Promise.all(zoneIds.map((id) => this.get(id)))).filter(
      (zone): zone is Zone => !!zone
    );
    const named = (tag: string) => {
      const zone = zones.find((z) => z.tags?.includes(tag));
      return zone ? zone.display_name || zone.name : "";
    };
    return [named("level"), named("building")].filter(Boolean).join(", ");
  }
}

export function getEvents(api: PlaceosApi, period = todayPeriod()) {
  // With no calendars/zones/systems, staff-api returns only the caller's own calendar.
  return api<CalendarEvent[]>(`/api/staff/v1/events?${periodQuery(period)}`);
}

function getBookings(api: PlaceosApi, type: string, period = todayPeriod()) {
  // No zones/email: staff-api scopes this to the caller. Don't add zones without user=current.
  return api<Booking[]>(
    `/api/staff/v1/bookings?type=${type}&${periodQuery(period)}` +
      "&include_booked_by=true&include_checked_out=true"
  );
}

const byStart = (a: TodayItem, b: TodayItem) => a.start - b.start;

const upcoming = (events: CalendarEvent[], now: number) =>
  events
    .filter((event) => event.status !== "cancelled" && !isAllDay(event) && event.event_end > now)
    .sort((a, b) => a.event_start - b.event_start)[0];

/**
 * The user's next calendar event (any event, room or not) that hasn't ended: from today's events, else
 * the next 7 days (one extra calendar query, only when nothing is left today). Null if there's none.
 */
export async function nextMeetingItem(
  api: PlaceosApi,
  todaysEvents: CalendarEvent[]
): Promise<TodayItem | null> {
  const now = Math.floor(Date.now() / 1000);
  let event = upcoming(todaysEvents, now);
  if (!event) {
    event = upcoming(await getEvents(api, { start: now, end: now + 7 * 24 * 3600 }), now);
  }
  if (!event) {
    return null;
  }
  const minutesAway = Math.round((event.event_start - now) / 60);
  const badge: TodayItem["badge"] =
    event.event_start <= now
      ? { text: "Now", kind: "ok" }
      : minutesAway <= 60
        ? { text: `In ${minutesAway} min`, kind: "ok" }
        : undefined;
  const url = event.online_meeting_url;
  return {
    start: event.event_start,
    title: event.private ? "Private meeting" : event.title || "(No title)",
    details: [
      whenText(event.event_start, event.event_end),
      event.system?.display_name || event.system?.name || event.location || "",
    ].filter(Boolean),
    badge,
    link: url && /^https:\/\//.test(url) ? { text: "Join", url } : undefined,
  };
}

/** Meetings in a PlaceOS room (events whose room matched a PlaceOS system). */
export function roomItems(events: CalendarEvent[]): TodayItem[] {
  return events
    .filter((event) => event.system?.id && event.status !== "cancelled")
    .map((event) => ({
      start: event.event_start,
      title: event.system?.display_name || event.system?.name || "Room",
      details: [
        timeRange(event.event_start, event.event_end, isAllDay(event)),
        event.private ? "Private meeting" : event.title || "",
      ].filter(Boolean),
      badge: timingBadge(event.event_start, event.event_end),
    }))
    .sort(byStart);
}

export async function deskItems(
  api: PlaceosApi,
  lookups: PlaceosLookups,
  settingsFor: SettingsForZones = () => Promise.resolve(null)
): Promise<TodayItem[]> {
  const bookings = await getBookings(api, "desk");
  const items = await Promise.all(
    bookings
      .filter((booking) => !booking.deleted)
      .map(async (booking) => ({
        start: booking.booking_start,
        title: (await lookups.assetName(booking)) || "Desk",
        details: [
          timeRange(booking.booking_start, booking.booking_end, isAllDay(booking)),
          await lookups.location(booking.zones),
        ].filter(Boolean),
        badge: bookingBadge(booking, "Checked in"),
        actions: bookingActions(api, booking, await settingsFor(booking.zones || []), "desk"),
        location: await lookups.mapLocation(booking),
      }))
  );
  return items.sort(byStart);
}

const isRecurring = (booking: Booking) =>
  !!booking.instance || (!!booking.recurrence_type && booking.recurrence_type !== "none");

/**
 * Check in or out, as Workplace's setBookingCheckedIn: only this occurrence of a recurring booking (its first
 * occurrence has no instance, so its start is used).
 */
function setCheckedIn(api: PlaceosApi, booking: Booking, state: boolean) {
  const id = encodeURIComponent(booking.id);
  const path = isRecurring(booking)
    ? `${id}/check_in/${booking.instance || booking.booking_start}`
    : `${id}/check_in`;
  return api(`/api/staff/v1/bookings/${path}?state=${state}`, { method: "POST" });
}

/** Cancels the booking, or only this occurrence of a recurring one. */
function cancelBooking(api: PlaceosApi, booking: Booking) {
  const id = encodeURIComponent(booking.id);
  const path = isRecurring(booking)
    ? `${id}/instance/${booking.instance || booking.booking_start}`
    : id;
  return api(`/api/staff/v1/bookings/${path}`, { method: "DELETE" });
}

/** A setting that's on in any of app.<type>, app.<type>s or app.bookings (e.g. app.desk, app.desks). */
const anySetting = (settings: AppSettings | null, type: string, key: string) =>
  !!settings &&
  [type, `${type}s`, "bookings"].some(
    (prefix) => settings.get<boolean>(`${prefix}.${key}`) === true
  );

/**
 * The buttons Workplace's booking details offers (can_checkin, can_cancel). Check in from 15 minutes before
 * the start until the end, unless app.desks.hide_checkin or auto_checkin (app.parking.* for parking) is
 * set, or it's a parking request with no space allocated yet; check out once checked in. Cancel until the
 * booking ends or is checked in.
 */
function bookingActions(
  api: PlaceosApi,
  booking: Booking,
  settings: AppSettings | null,
  type: "desk" | "parking"
): TodayAction[] {
  const now = Date.now() / 1000;
  const state = booking.current_state || "";
  const cancelled =
    booking.rejected ||
    ["cancelled", "declined"].includes(booking.status || "") ||
    ["cancelled", "rejected"].includes(state);
  const checkedOut = !!booking.checked_out_at || state === "checked_out";
  const done = checkedOut || now >= booking.booking_end || state === "ended";
  if (cancelled || done) {
    return [];
  }
  const checkedIn = !!booking.checked_in || state === "checked_in";
  const actions: TodayAction[] = [];
  const noun = type === "desk" ? "desk" : "parking space";
  const unallocated = type === "parking" && (booking.asset_id || "").startsWith("unallocated");
  const checkinAllowed =
    !unallocated &&
    !anySetting(settings, type, "hide_checkin") &&
    !anySetting(settings, type, "auto_checkin");
  if (checkinAllowed && checkedIn) {
    actions.push({
      label: "Check out",
      secondary: true,
      confirm: `Check out now? The ${noun} becomes free for others to book.`,
      run: () => setCheckedIn(api, booking, false),
    });
  } else if (checkinAllowed && now >= booking.booking_start - 15 * 60) {
    actions.push({ label: "Check in", run: () => setCheckedIn(api, booking, true) });
  }
  if (!checkedIn) {
    actions.push({
      label: "Cancel",
      secondary: true,
      confirm: isRecurring(booking)
        ? `Cancel this day's ${type} booking? The rest of the series stays booked.`
        : `Cancel this ${type} booking?`,
      run: () => cancelBooking(api, booking),
    });
  }
  return actions;
}

export async function parkingItems(
  api: PlaceosApi,
  lookups: PlaceosLookups,
  settingsFor: SettingsForZones = () => Promise.resolve(null)
): Promise<TodayItem[]> {
  const bookings = await getBookings(api, "parking");
  const items = await Promise.all(
    bookings
      .filter((booking) => !booking.deleted)
      .map(async (booking) => ({
        start: booking.booking_start,
        title: (await lookups.assetName(booking)) || "Parking",
        details: [
          timeRange(booking.booking_start, booking.booking_end, isAllDay(booking)),
          await lookups.location(booking.zones),
          booking.extension_data?.plate_number
            ? `Plate ${booking.extension_data.plate_number}`
            : "",
        ].filter(Boolean),
        badge: bookingBadge(booking, "Checked in"),
        actions: bookingActions(api, booking, await settingsFor(booking.zones || []), "parking"),
        location: await lookups.mapLocation(booking),
      }))
  );
  return items.sort(byStart);
}

/**
 * Visitors the user is hosting: visitor bookings (minus those linked to a meeting, which Workplace shows on
 * the meeting) plus guests expected at the user's meetings.
 */
export async function visitorItems(api: PlaceosApi, events: CalendarEvent[]): Promise<TodayItem[]> {
  const bookings = await getBookings(api, "visitor");
  const fromBookings: TodayItem[] = bookings
    .filter((booking) => !booking.deleted && !booking.linked_event)
    .map((booking) => {
      const attendee = booking.attendees?.[0];
      const name =
        booking.extension_data?.group_members?.[0]?.name ||
        attendee?.name ||
        booking.extension_data?.visitor_name ||
        booking.asset_id ||
        "Visitor";
      return {
        start: booking.booking_start,
        title: name,
        details: [
          timeRange(booking.booking_start, booking.booking_end, isAllDay(booking)),
          booking.extension_data?.company || attendee?.organisation || "",
          booking.title || booking.description || "",
        ].filter(Boolean),
        badge: booking.checked_in ? { text: "Arrived", kind: "ok" } : undefined,
      };
    });

  const fromMeetings: TodayItem[] = [];
  for (const event of events) {
    if (event.status === "cancelled") {
      continue;
    }
    for (const attendee of event.attendees ?? []) {
      if (attendee.visit_expected && !attendee.resource && !attendee.organizer) {
        fromMeetings.push({
          start: event.event_start,
          title: attendee.name || attendee.email,
          details: [
            timeRange(event.event_start, event.event_end, isAllDay(event)),
            attendee.organisation || "",
            event.private ? "Private meeting" : event.title || "",
          ].filter(Boolean),
          badge: attendee.checked_in ? { text: "Arrived", kind: "ok" } : undefined,
        });
      }
    }
  }
  return [...fromBookings, ...fromMeetings].sort(byStart);
}
