/* This file reads PlaceOS data for Phase 5. All calls are read-only GETs. */

import { getResourceTypeId } from "./desk-booking-data";
import type { PlaceosRequestInit } from "./placeos-helper";

/**
 * Authenticated PlaceOS call (placeosApi in app-ui.ts). GET unless init says otherwise.
 * Pass `allPages` for engine list routes, which page with Link headers.
 */
export type PlaceosApi = <T>(path: string, init?: PlaceosRequestInit) => Promise<T>;

/** Every page of a paginated list route. */
export const allPages = { allPages: true } as const;

export type Authority = {
  id: string;
  name: string;
  domain: string;
  config?: { org_zone?: string };
};

export type Zone = {
  id: string;
  name: string;
  display_name?: string;
  tags?: string[];
};

export type Room = {
  id: string;
  name: string;
  display_name?: string;
  email?: string;
  capacity?: number;
};

export type DeskSummary = {
  levels: number;
  desks: number;
  bookedToday: number;
};

/** A list response, either a plain array or a paginated { results } object. */
export const resultList = <T>(result: T[] | { results?: T[] } | null): T[] =>
  Array.isArray(result) ? result : result?.results || [];

export const zoneName = (zone: { name: string; display_name?: string }) =>
  zone.display_name || zone.name;

/**
 * Staff API has no current-user endpoint; this one runs its full auth path (token, scope, domain, tenant)
 * without needing Graph, so success means Staff API accepts the token.
 */
export function checkStaffApi(api: PlaceosApi): Promise<unknown> {
  return api<unknown>("/api/staff/v1/tenants/current_limits");
}

export function getAuthority(api: PlaceosApi): Promise<Authority> {
  return api<Authority>("/auth/authority");
}

export function getBuildings(api: PlaceosApi, orgZoneId?: string): Promise<Zone[]> {
  const parent = orgZoneId ? `&parent_id=${encodeURIComponent(orgZoneId)}` : "";
  return api<Zone[]>(`/api/engine/v2/zones?tags=building${parent}&limit=500`, allPages);
}

export function getBookableRooms(api: PlaceosApi, buildingId: string): Promise<Room[]> {
  return api<Room[]>(
    `/api/engine/v2/systems?zone_id=${encodeURIComponent(buildingId)}&bookable=true&limit=500`,
    allPages
  );
}

/** Counts the building's desk assets and the levels they're on, and desks booked today in the building. */
export async function getDeskSummary(api: PlaceosApi, buildingId: string): Promise<DeskSummary> {
  const building = encodeURIComponent(buildingId);
  // Assets list their building in `zones`.
  const typeId = await getResourceTypeId(api, "desk");
  const desks = typeId
    ? resultList(
        await api<{ zone_id?: string }[]>(
          `/api/engine/v2/assets?zones=${building}&type_id=${encodeURIComponent(typeId)}&limit=500`,
          allPages
        )
      )
    : [];

  const start = new Date();
  start.setHours(0, 0, 0, 0);
  const end = new Date(start.getTime() + 24 * 60 * 60 * 1000);
  // Paged like Workplace's bookedResourceList (libs/bookings bookings.fn.ts): at least 200 per page.
  const booked = await api<string[] | { results?: string[] }>(
    `/api/staff/v1/bookings/booked?type=desk&period_start=${Math.floor(start.getTime() / 1000)}` +
      `&period_end=${Math.floor(end.getTime() / 1000)}&zones=${building}&limit=200`,
    allPages
  );

  return {
    levels: new Set(desks.map((desk) => desk.zone_id).filter(Boolean)).size,
    desks: desks.length,
    // Workplace removes duplicates too.
    bookedToday: new Set(resultList(booked)).size,
  };
}
