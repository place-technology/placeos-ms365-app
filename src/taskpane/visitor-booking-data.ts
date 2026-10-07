/*
 * Visitor invites: settings, checks and creating the visitor booking, as PlaceOS Workplace's invite visitor
 * form does for one visitor (user-interfaces libs/bookings invite-visitor-form.component.ts and
 * BookingFormService.postForm).
 */

import type { AppSettings, Organisation } from "./booking-settings";
import { resultList, type PlaceosApi } from "./placeos-data";
import { isFullDay } from "./room-booking-data";

/** The form's values. Times are epoch ms; durations are minutes. */
export type VisitorRequest = {
  start: number;
  duration: number;
  allDay: boolean;
  buildingId: string;
  name: string;
  email: string;
  company: string;
  reason: string;
  passNumber: string;
  international: boolean;
};

/** A past visitor, from the user's "visitor-invitees" setting. */
export type PastVisitor = { email: string; name: string; company: string; international: boolean };

const unix = (ms: number) => Math.floor(ms / 1000);

// User setting with the visitors the user has invited before, as "email|name|company|international".
const inviteesKey = "visitor-invitees";

/**
 * A visitor setting, looked up as Workplace's BookingFormService.setting does for visitors:
 * app.visitor.<key>, then app.visitors.<key>, then app.bookings.<key>.
 */
export function visitorSetting<T>(settings: AppSettings, key: string, fallback?: T): T {
  for (const prefix of ["visitor", "visitors", "bookings"]) {
    const value = settings.get<T>(`${prefix}.${key}`);
    if (value !== undefined && value !== null) {
      return value;
    }
  }
  return fallback as T;
}

/** All day is off unless app.visitors.allow_all_day (else app.bookings.allow_all_day) is true. */
export const visitorAllDayAllowed = (settings: AppSettings): boolean =>
  settings.get<boolean>("visitors.allow_all_day") ??
  settings.get<boolean>("bookings.allow_all_day") ??
  false;

/** Shortest and longest visits, in minutes (Workplace's visitor form: 30 and 240 by default). */
export function visitorDurationLimits(settings: AppSettings) {
  const pick = (key: string, fallback: number) =>
    settings.get<number>(`visitors.${key}`) || settings.get<number>(`bookings.${key}`) || fallback;
  const min = pick("min_duration", 30);
  return { min, max: Math.max(min, pick("max_duration", 4 * 60)) };
}

/** Visitors the user has invited before, most recent last. */
export function pastVisitors(org: Organisation): PastVisitor[] {
  return org
    .userList(inviteesKey)
    .filter((item): item is string => typeof item === "string")
    .map((item) => {
      const [email = "", name = "", company = "", international] = item.split("|");
      return { email, name: name || email, company, international: international === "1" };
    })
    .filter((visitor) => !!visitor.email);
}

/** Remembers the visitor for next time, replacing any earlier entry for the same email. */
export function rememberVisitor(org: Organisation, request: VisitorRequest): Promise<void> {
  const email = request.email.toLowerCase();
  const others = org
    .userList(inviteesKey)
    .filter((item) => `${item}`.split("|")[0].toLowerCase() !== email);
  const entry = [request.email, request.name, request.company, request.international ? "1" : "0"]
    .map((part) => part.replace(/\|/g, " "))
    .join("|");
  return org.saveUserSetting(inviteesKey, [...others, entry]);
}

type ExistingBooking = {
  id: string;
  asset_id?: string;
  user_email?: string;
  status?: string;
  rejected?: boolean;
  deleted?: boolean;
  booking_end?: number;
  checked_out_at?: number;
};

// Not cancelled, declined or ended, as Workplace's Booking.status works it out: an ended booking was
// checked out or its end has passed.
const isActive = (booking: ExistingBooking) =>
  !booking.rejected &&
  !booking.deleted &&
  !["declined", "cancelled"].includes(booking.status || "") &&
  !booking.checked_out_at &&
  !(booking.booking_end && booking.booking_end * 1000 < Date.now());

/**
 * Workplace's check just before booking (_checkResourceAvailable). Returns an error message, or "".
 * The visitor may already be invited by the user at this time, or the user may have reached
 * app.bookings.allowed_daily_visitor_count (default 1; 0 or less turns the limit off) for overlapping visits.
 * If the bookings can't be read, the check passes, as Workplace's queryBookings returns [] on errors.
 */
export async function checkBeforeInvite(
  api: PlaceosApi,
  settings: AppSettings,
  org: Organisation,
  request: VisitorRequest
): Promise<string> {
  const bookings = await api<ExistingBooking[]>(
    `/api/staff/v1/bookings?period_start=${unix(request.start)}` +
      `&period_end=${unix(request.start + request.duration * 60000)}&type=visitor` +
      `&email=${encodeURIComponent(org.user.email)}&limit=1000`
  ).catch((): ExistingBooking[] => []);
  const active = resultList(bookings).filter(isActive);
  const email = request.email.toLowerCase();
  if (active.some((booking) => (booking.asset_id || "").toLowerCase() === email)) {
    return `${request.email} is already invited at this time.`;
  }
  const limit = settings.get<number>("bookings.allowed_daily_visitor_count") ?? 1;
  const mine = org.user.email.toLowerCase();
  const count = active.filter(
    (booking) => (booking.user_email || "").toLowerCase() === mine
  ).length;
  if (limit > 0 && count >= limit) {
    return limit === 1
      ? "You already have a visitor booked at this time."
      : `You already have ${count} visitors booked at this time, the most you can invite.`;
  }
  return "";
}

/** The booking's zones: org, region and building, as Workplace's visitor form sets them. */
function visitorZones(org: Organisation, buildingId: string): string[] {
  const building = org.buildings.find((b) => b.id === buildingId);
  return [org.org?.id, building?.parent_id, buildingId].filter(
    (id, index, list): id is string => !!id && list.indexOf(id) === index
  );
}

/**
 * The Staff API booking body, as Workplace's postForm and Booking.toJSON() build it for one visitor, with
 * only what's set. A whole-day visit ends at 23:59, as Workplace's Booking does.
 */
export function buildVisitorBooking(
  request: VisitorRequest,
  org: Organisation,
  settings: AppSettings
) {
  // Workplace's BookingFormService.timezone: use_building_timezone resolved as visitorSetting does.
  const browserTimezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const timezone = visitorSetting<boolean>(settings, "use_building_timezone", false)
    ? org.buildings.find((b) => b.id === request.buildingId)?.timezone || browserTimezone
    : browserTimezone;
  const fullDay = request.allDay && isFullDay(request.start, request.duration);
  const duration = fullDay ? request.duration - 1 : request.duration;
  const user = org.user;
  const name = request.name || request.email;
  const reason = request.reason || "Visit";
  const visitor: Record<string, string> = { name, email: request.email };
  if (request.company) {
    visitor.organisation = request.company;
  }
  const extension: Record<string, unknown> = {
    visitor_name: name,
    international: request.international,
    department: user.department,
    app_name: "PlaceOS Outlook add-in",
  };
  if (request.company) {
    extension.company = request.company;
  }
  if (request.passNumber) {
    extension.pass_number = request.passNumber;
  }
  return {
    type: "visitor",
    booking_type: "visitor",
    asset_id: request.email,
    asset_ids: [request.email],
    asset_name: name,
    title: reason,
    description: reason,
    booking_start: unix(request.start),
    booking_end: unix(request.start + duration * 60000),
    all_day: request.allDay,
    timezone,
    user_id: user.id,
    user_email: user.email,
    user_name: user.name || user.email,
    booked_by_id: user.id,
    booked_by_email: user.email,
    zones: visitorZones(org, request.buildingId),
    approved: settings.get<boolean>("bookings.no_approval") === true,
    permission: "PRIVATE",
    attendees: [visitor],
    extension_data: extension,
  };
}

export function createVisitorBooking(api: PlaceosApi, body: unknown): Promise<{ id: string }> {
  return api<{ id: string }>("/api/staff/v1/bookings", {
    method: "POST",
    body: JSON.stringify(body),
  });
}
