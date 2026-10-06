/*
 * Desk and parking booking: desks or parking spaces, availability, checks and creating the booking, as
 * PlaceOS Workplace does with app.desks.use_assets and for parking (user-interfaces libs/assets
 * desk-assets.fn.ts and parking-assets.fn.ts, libs/bookings BookingFormService and ParkingService).
 *
 * Desks are PlaceOS Assets of the asset type "_DESKS_" in the hidden category "_DESKS_", parking spaces of
 * the type "_PARKING_SPACES_" in the hidden category "_PARKING_", one asset per desk or space, with zone_id
 * set to its level. The legacy "desks" and "parking-spaces" zone metadata isn't read.
 */

import type { AppSettings, Organisation, Zone } from "./booking-settings";
import { allPages, resultList, type PlaceosApi } from "./placeos-data";
import {
  getBookingRules,
  hiddenByRules,
  isFullDay,
  type BookingRuleset,
} from "./room-booking-data";

/** What's being booked. Both work the same way, with their own asset type, rules and settings. */
export type ResourceKind = "desk" | "parking";

type KindInfo = {
  category: string;
  type: string;
  // Where Workplace reads this kind's own form settings: app.desks.* or app.parking.*.
  settings: string;
  // Shortest booking and booking window defaults (Workplace's desk and parking forms).
  minDuration: number;
  availablePeriod: number;
  noun: string;
};

export const kinds: Record<ResourceKind, KindInfo> = {
  desk: {
    category: "_DESKS_",
    type: "_DESKS_",
    settings: "desks",
    minDuration: 60,
    availablePeriod: 90,
    noun: "desk",
  },
  parking: {
    category: "_PARKING_",
    type: "_PARKING_SPACES_",
    settings: "parking",
    minDuration: 30,
    availablePeriod: 7,
    noun: "parking space",
  },
};

type Named = { id: string; name: string; created_at?: number };

type ResourceAsset = {
  id: string;
  name?: string;
  identifier?: string;
  map_id?: string;
  zone_id?: string;
  zones?: string[];
  bookable?: boolean;
  assigned_to?: string;
  assigned_name?: string;
  place_groups?: string[];
  features?: string[];
  images?: string[];
  notes?: string;
};

/** A desk or parking space. */
export type Desk = {
  kind: ResourceKind;
  // The asset id ("asset-..."), sent as the booking's asset_id.
  id: string;
  name: string;
  // Id of its element on its level's floor plan.
  map_id: string;
  level: Zone;
  // The asset's own zones, which booking rules match and the booking also gets.
  zones: string[];
  bookable: boolean;
  groups: string[];
  features: string[];
  assigned_to: string;
  assigned_name: string;
};

/** The form's values. Times are epoch ms; durations are minutes. */
export type DeskRequest = {
  start: number;
  duration: number;
  allDay: boolean;
  buildingId: string;
  // Parking only.
  plateNumber?: string;
};

const unix = (ms: number) => Math.floor(ms / 1000);

/**
 * A booking setting for this kind, looked up as Workplace's settingForType does: app.<kind>.<key>, then
 * app.<kind>s.<key>, then app.bookings.<key>.
 */
export function kindSetting<T>(
  settings: AppSettings,
  kind: ResourceKind,
  key: string,
  fallback?: T
): T {
  for (const prefix of [kind, `${kind}s`, "bookings"]) {
    const value = settings.get<T>(`${prefix}.${key}`);
    if (value !== undefined && value !== null) {
      return value;
    }
  }
  return fallback as T;
}

/** A desk booking setting: app.desk.<key>, then app.desks.<key>, then app.bookings.<key>. */
export const deskSetting = <T>(settings: AppSettings, key: string, fallback?: T): T =>
  kindSetting(settings, "desk", key, fallback);

/** A setting of this kind's form (app.desks.<key> or app.parking.<key>), else app.bookings.<key>. */
export function formSetting<T>(
  settings: AppSettings,
  kind: ResourceKind,
  key: string
): T | undefined {
  return (
    settings.get<T>(`${kinds[kind].settings}.${key}`) ??
    settings.get<T>(`bookings.${key}`) ??
    undefined
  );
}

/** All day is offered unless app.desks/app.parking.allow_all_day (else app.bookings.allow_all_day) is false. */
export function deskAllDayAllowed(settings: AppSettings, kind: ResourceKind = "desk"): boolean {
  return formSetting<boolean>(settings, kind, "allow_all_day") ?? true;
}

/**
 * Shortest and longest bookings, in minutes (Workplace's desk form: 60 and 480 by default; parking: 30 and
 * 480).
 */
export function deskDurationLimits(settings: AppSettings, kind: ResourceKind = "desk") {
  const pick = (key: string, fallback: number) =>
    settings.get<number>(`${kinds[kind].settings}.${key}`) ||
    settings.get<number>(`bookings.${key}`) ||
    fallback;
  const min = pick("min_duration", kinds[kind].minDuration);
  return { min, max: Math.max(min, pick("max_duration", 480)) };
}

const oldest = <T extends Named>(items: T[], name: string): T | undefined =>
  items
    .filter((item) => item.name === name)
    .sort((a, b) => (a.created_at ?? 0) - (b.created_at ?? 0))[0];

const typeIds = new Map<string, Promise<string | null>>();

/**
 * Id of the asset type with this name in the hidden category with this name, or null if the domain has none.
 * Workplace creates the category and type when they're missing; the add-in only reads them.
 */
function getAssetTypeId(api: PlaceosApi, categoryName: string, typeName: string) {
  const key = `${categoryName}/${typeName}`;
  let id = typeIds.get(key);
  if (!id) {
    id = (async () => {
      const categories = resultList(
        await api<Named[]>("/api/engine/v2/asset_categories?hidden=true&limit=500", allPages)
      );
      const category = oldest(categories, categoryName);
      if (!category) {
        return null;
      }
      const types = resultList(
        await api<Named[]>(
          `/api/engine/v2/asset_types?category_id=${encodeURIComponent(category.id)}&limit=500`,
          allPages
        )
      );
      return oldest(types, typeName)?.id ?? null;
    })();
    typeIds.set(key, id);
    id.catch(() => typeIds.delete(key));
  }
  return id;
}

/** Id of the "_DESKS_" or "_PARKING_SPACES_" asset type, or null if the domain has none. */
export const getResourceTypeId = (api: PlaceosApi, kind: ResourceKind) =>
  getAssetTypeId(api, kinds[kind].category, kinds[kind].type);

/**
 * As desk-assets.fn deskFromAsset: the name is the identifier, and bookable must be set. Workplace's parking
 * list doesn't check bookable, but its Concierge sets it on every space.
 */
function deskFromAsset(asset: ResourceAsset, level: Zone, kind: ResourceKind): Desk {
  return {
    kind,
    id: asset.id,
    name: asset.identifier || asset.name || asset.id,
    map_id: asset.map_id || asset.id,
    level,
    zones: asset.zones || [],
    bookable: asset.bookable === true,
    groups: asset.place_groups || [],
    features: asset.features || [],
    assigned_to: asset.assigned_to || "",
    assigned_name: asset.assigned_name || "",
  };
}

/**
 * Desks or parking spaces on these levels: one assets query per level, as Workplace's
 * queryDeskAssetsForZones and queryParkingSpacesForZones.
 */
export async function getDesks(
  api: PlaceosApi,
  levels: Zone[],
  kind: ResourceKind = "desk"
): Promise<Desk[]> {
  const typeId = await getResourceTypeId(api, kind);
  if (!typeId || !levels.length) {
    return [];
  }
  const perLevel = await Promise.all(
    levels.map(async (level) => {
      const assets = await api<ResourceAsset[]>(
        `/api/engine/v2/assets?zone_id=${encodeURIComponent(level.id)}` +
          `&type_id=${encodeURIComponent(typeId)}&limit=500`,
        allPages
      );
      return resultList(assets).map((asset) => deskFromAsset(asset, level, kind));
    })
  );
  return perLevel.flat();
}

/** The user's entry in the building's parking users (Concierge), as Workplace's ParkingService.user_details. */
export type ParkingUser = { plate_number: string; deny: boolean };

type ParkingUserAsset = { id: string; other_data?: Record<string, unknown> };

/**
 * The signed-in user's parking user entry in the building: an asset of the type "_PARKING_USERS_" (category
 * "_PARKING_") with zone_id = the building and other_data.email = the user. Null if there's none.
 */
export async function getParkingUser(
  api: PlaceosApi,
  buildingId: string,
  email: string
): Promise<ParkingUser | null> {
  const typeId = await getAssetTypeId(api, kinds.parking.category, "_PARKING_USERS_");
  if (!typeId) {
    return null;
  }
  const assets = resultList(
    await api<ParkingUserAsset[]>(
      `/api/engine/v2/assets?zone_id=${encodeURIComponent(buildingId)}` +
        `&type_id=${encodeURIComponent(typeId)}&limit=500`,
      allPages
    )
  );
  const lower = email.toLowerCase();
  const data = assets
    .map((asset) => asset.other_data || {})
    .find((other) => `${other.email || ""}`.toLowerCase() === lower);
  return data
    ? { plate_number: `${data.plate_number || ""}`, deny: `${data.deny}` === "true" }
    : null;
}

/**
 * Ids of desks or spaces with a booking overlapping the period, in the zone (building, or region with
 * use_region).
 */
export async function getBookedDeskIds(
  api: PlaceosApi,
  zoneId: string,
  request: Pick<DeskRequest, "start" | "duration">,
  deskCount: number,
  kind: ResourceKind = "desk"
): Promise<Set<string>> {
  const ids = await api<string[]>(
    `/api/staff/v1/bookings/booked?period_start=${unix(request.start)}` +
      `&period_end=${unix(request.start + request.duration * 60000)}&type=${kind}` +
      `&zones=${encodeURIComponent(zoneId)}&limit=${Math.max(200, deskCount)}`,
    allPages
  );
  return new Set(resultList(ids));
}

/** "desk_booking_rules" or "parking_booking_rules" metadata for each building, by building id. */
export async function getDeskRules(
  api: PlaceosApi,
  buildingIds: string[],
  kind: ResourceKind = "desk"
): Promise<Map<string, BookingRuleset[]>> {
  const rules = await Promise.all(
    buildingIds.map((id) => getBookingRules(api, id, `${kind}_booking_rules`))
  );
  return new Map(buildingIds.map((id, index) => [id, rules[index]]));
}

/**
 * Desks or spaces the user may book at this time: bookable, in one of its groups (if it has any) and not
 * hidden by the building's booking rules. Booked ones are still included.
 */
export function bookableDesks(
  desks: Desk[],
  rules: Map<string, BookingRuleset[]>,
  request: DeskRequest,
  groups: string[]
): Desk[] {
  return desks.filter((desk) => {
    if (!desk.bookable) {
      return false;
    }
    if (desk.groups.length && !desk.groups.some((group) => groups.includes(group))) {
      return false;
    }
    const rulesets = rules.get(desk.level.parent_id || "") || [];
    const resource = { id: desk.id, name: desk.name, zones: [desk.level.id, ...desk.zones] };
    return !hiddenByRules(rulesets, resource, request.start, request.duration, groups);
  });
}

/**
 * The desk or space assigned to the user, if they may not book another one.
 * app.bookings.assigned_resource_booking (looked up for the kind first) is "allow", "deny" or "other_only"
 * (the default); only "allow" lets someone with an assigned desk or space book another for themselves.
 */
export function blockingAssignedDesk(
  settings: AppSettings,
  desks: Desk[],
  email: string,
  kind: ResourceKind = "desk"
): Desk | undefined {
  if (kindSetting<string>(settings, kind, "assigned_resource_booking", "other_only") === "allow") {
    return undefined;
  }
  const lower = email.toLowerCase();
  return desks.find((desk) => desk.assigned_to.toLowerCase() === lower);
}

type ExistingBooking = {
  id: string;
  asset_id?: string;
  user_email?: string;
  status?: string;
  rejected?: boolean;
  deleted?: boolean;
  checked_out_at?: number;
  current_state?: string;
  extension_data?: { is_assigned?: boolean };
};

const isActive = (booking: ExistingBooking) =>
  !booking.rejected &&
  !booking.deleted &&
  !booking.checked_out_at &&
  !["declined", "cancelled", "ended"].includes(booking.status || "") &&
  !["checked_out", "ended", "rejected"].includes(booking.current_state || "");

/**
 * Workplace's checks just before booking (_checkResourceAvailable). Returns an error message, or "".
 * Someone else may have taken the desk or space since the list loaded, the user may already have one at this
 * time (app.bookings.allowed_daily_desk_count / allowed_daily_parking_count, default 1; 0 or less turns the
 * limit off), or have an assigned one booked.
 */
export async function checkBeforeBooking(
  api: PlaceosApi,
  settings: AppSettings,
  org: Organisation,
  desk: Desk,
  request: DeskRequest
): Promise<string> {
  const period =
    `period_start=${unix(request.start)}` +
    `&period_end=${unix(request.start + request.duration * 60000)}`;
  const { kind } = desk;
  const noun = kinds[kind].noun;
  const email = org.user.email.toLowerCase();
  const [booked, mine] = await Promise.all([
    getBookedDeskIds(api, desk.level.id, request, 1, kind),
    api<ExistingBooking[]>(
      `/api/staff/v1/bookings?${period}&type=${kind}` +
        `&email=${encodeURIComponent(org.user.email)}&limit=1000`
    ),
  ]);
  const active = resultList(mine).filter(isActive);
  if (active.some((booking) => booking.asset_id === desk.id)) {
    return `You've already booked this ${noun} at this time.`;
  }
  if (booked.has(desk.id)) {
    return `This ${noun} has just been booked by someone else. Choose another ${noun}.`;
  }
  const assignedAllowed =
    kindSetting<string>(settings, kind, "assigned_resource_booking", "other_only") === "allow";
  if (!assignedAllowed && active.some((booking) => booking.extension_data?.is_assigned)) {
    return `You have an assigned ${noun} and cannot book another ${noun}.`;
  }
  const limit = settings.get<number>(`bookings.allowed_daily_${kind}_count`) ?? 1;
  const count = active.filter(
    (booking) => (booking.user_email || "").toLowerCase() === email
  ).length;
  if (limit > 0 && count >= limit) {
    return limit === 1
      ? `You already have a ${noun} booked at this time.`
      : `You already have ${count} ${noun}s booked at this time, the most you can book.`;
  }
  return "";
}

/** The booking's zones: org, region, building and level, plus the asset's own zones. */
function deskZones(org: Organisation, desk: Desk): string[] {
  const building = org.buildings.find((b) => b.id === desk.level.parent_id);
  return [
    org.org?.id,
    building?.parent_id,
    desk.level.parent_id,
    desk.level.id,
    ...desk.zones,
  ].filter((id, index, list): id is string => !!id && list.indexOf(id) === index);
}

/**
 * The Staff API booking body, as Workplace's postForm and Booking.toJSON() build it for a desk or parking
 * space, with only what's set. A whole-day booking ends at 23:59, as Workplace's Booking does.
 */
export function buildDeskBooking(
  request: DeskRequest,
  desk: Desk,
  org: Organisation,
  settings: AppSettings
) {
  const { kind } = desk;
  const browserTimezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const timezone = kindSetting<boolean>(settings, kind, "use_building_timezone")
    ? org.buildings.find((b) => b.id === desk.level.parent_id)?.timezone || browserTimezone
    : browserTimezone;
  const fullDay = request.allDay && isFullDay(request.start, request.duration);
  const duration = fullDay ? request.duration - 1 : request.duration;
  const user = org.user;
  // Workplace stores the desk as the assigned asset, and for parking the plate and the user's groups.
  const kindData =
    kind === "desk"
      ? { assigned_asset_id: desk.id, assigned_asset_name: desk.name }
      : {
          plate_number: request.plateNumber || undefined,
          requires_manual_approval: false,
          user_groups: user.groups || [],
        };
  return {
    type: kind,
    booking_type: kind,
    asset_id: desk.id,
    asset_ids: [desk.id],
    asset_name: desk.name,
    description: desk.name,
    title: kind === "desk" ? "Desk Booking" : "Parking Booking",
    booking_start: unix(request.start),
    booking_end: unix(request.start + duration * 60000),
    all_day: request.allDay,
    timezone,
    user_id: user.id,
    user_email: user.email,
    user_name: user.name || user.email,
    booked_by_id: user.id,
    booked_by_email: user.email,
    zones: deskZones(org, desk),
    approved: settings.get<boolean>("bookings.no_approval") === true,
    permission: "PRIVATE",
    attendees: [],
    extension_data: {
      ...kindData,
      map_id: desk.map_id,
      name: desk.name,
      department: user.department,
      app_name: "PlaceOS Outlook add-in",
    },
  };
}

export function createDeskBooking(api: PlaceosApi, body: unknown): Promise<{ id: string }> {
  return api<{ id: string }>("/api/staff/v1/bookings", {
    method: "POST",
    body: JSON.stringify(body),
  });
}
