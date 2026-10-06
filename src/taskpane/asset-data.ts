/*
 * Assets for room bookings (equipment such as laptops or projectors), as in the legacy Outlook add-in
 * (libs/assets AssetStateService, asset-list-field and validateAssetRequestsForResource).
 * Availability is per asset type: its assets in the building minus those in overlapping asset requests.
 * Requests are saved on the event (extension_data.assets) and as "asset-request" bookings linked to it.
 */

import { getMetadata, type AppSettings } from "./booking-settings";
import { attachedRulesAllow, type AttachedRule } from "./catering-data";
import { allPages, resultList as list, type PlaceosApi } from "./placeos-data";
import type { LinkedBooking, Space } from "./room-booking-data";

type AssetCategory = { id: string; name: string; hidden?: boolean };
type AssetTypeRecord = {
  id: string;
  name: string;
  brand?: string;
  description?: string;
  category_id?: string;
  images?: string[];
  hidden?: boolean;
};
type Asset = { id: string; asset_type_id?: string; hidden?: boolean };
type AssetBooking = { asset_id?: string; asset_ids?: string[]; status?: string };

/** An asset type with the ids of its assets that are free for the booking. */
export type AssetType = AssetTypeRecord & { category: string; available: string[] };

/** A requested asset type and quantity. */
export type AssetSelection = { type: AssetType; quantity: number };

/** Assets are offered when app.events.allow_assets is set (and app.has_assets isn't false). */
export function assetsEnabled(settings: AppSettings): boolean {
  return (
    settings.get<boolean>("events.allow_assets") === true && settings.get("has_assets") !== false
  );
}

/** Rooms listed in the building's "assets-settings" disabled_rooms can't have assets. */
export async function assetsDisabledForRoom(
  api: PlaceosApi,
  buildingId: string,
  room: Space
): Promise<boolean> {
  const details = await getMetadata<{ disabled_rooms?: string[] }>(
    api,
    buildingId,
    "assets-settings"
  );
  return !!details?.disabled_rooms?.includes(room.id);
}

/** Asset types with at least one free asset in the zone for the booking period. */
export async function getAvailableAssetTypes(
  api: PlaceosApi,
  zoneId: string,
  room: Space,
  start: number,
  duration: number
): Promise<AssetType[]> {
  const zone = encodeURIComponent(zoneId);
  const period =
    `period_start=${Math.floor(start / 1000)}` +
    `&period_end=${Math.floor((start + duration * 60000) / 1000)}`;
  const [types, assets, bookings, categories, rules] = await Promise.all([
    api<AssetTypeRecord[]>(`/api/engine/v2/asset_types?zone_id=${zone}&limit=500`, allPages),
    api<Asset[]>(`/api/engine/v2/assets?zone_id=${zone}&limit=500`, allPages),
    api<AssetBooking[]>(
      `/api/staff/v1/bookings?zones=${zone}&${period}&type=asset-request&rejected=false`
    ),
    api<AssetCategory[]>("/api/engine/v2/asset_categories?limit=500", allPages),
    getMetadata<AttachedRule[]>(api, zoneId, "assets_config"),
  ]);
  const busy = new Set<string>();
  for (const booking of list(bookings)) {
    if (booking.status !== "declined" && booking.status !== "cancelled") {
      [booking.asset_id, ...(booking.asset_ids || [])].forEach((id) => id && busy.add(id));
    }
  }
  const visibleCategories = new Map(
    list(categories)
      .filter((category) => !category.hidden)
      .map((category) => [category.id, category.name])
  );
  return list(types)
    .filter((type) => !type.hidden && type.category_id && visibleCategories.has(type.category_id))
    .map((type) => ({
      ...type,
      category: visibleCategories.get(type.category_id as string) || "",
      available: list(assets)
        .filter((asset) => !asset.hidden && asset.asset_type_id === type.id && !busy.has(asset.id))
        .map((asset) => asset.id),
    }))
    .filter(
      (type) =>
        type.available.length > 0 &&
        attachedRulesAllow([type.name, type.category], rules, room, start, duration)
    )
    .sort((a, b) => a.name.localeCompare(b.name));
}

/** Asset filters, as in Workplace's asset select modal (AssetStateService). Categories are names. */
export type AssetFilters = { search: string; categories: string[]; favouritesOnly: boolean };

export function filterAssetTypes(
  types: AssetType[],
  filters: AssetFilters,
  favourites: string[]
): AssetType[] {
  const search = filters.search.trim().toLowerCase();
  return types.filter(
    (type) =>
      (!search ||
        type.name.toLowerCase().includes(search) ||
        (type.description || "").toLowerCase().includes(search) ||
        (type.brand || "").toLowerCase().includes(search)) &&
      (!filters.categories.length || filters.categories.includes(type.category)) &&
      (!filters.favouritesOnly || favourites.includes(type.id))
  );
}

/** Delivery offset limits in minutes after the start (app.assets.min_offset / end_offset). */
export function assetOffsetLimits(settings: AppSettings, duration: number) {
  const min = settings.get<number>("assets.min_offset", 0) || 0;
  return {
    min,
    max: Math.max(min, duration - (settings.get<number>("assets.end_offset", 0) || 0)),
  };
}

const requestId = () => `order-${Math.floor(1000000 + Math.random() * 8999999)}`;

/**
 * AssetRequest JSON (libs/common asset-request.class toJSON), picking the specific free assets.
 * Call with fresh availability right before saving. Throws if an asset type no longer has enough.
 */
export function buildAssetRequest(
  selections: AssetSelection[],
  available: AssetType[],
  start: number,
  deliverOffset: number
) {
  const items = selections
    .filter((selection) => selection.quantity > 0)
    .map((selection) => {
      const free = available.find((type) => type.id === selection.type.id)?.available || [];
      if (free.length < selection.quantity) {
        throw new Error(`Not enough ${selection.type.name} available any more.`);
      }
      return {
        id: selection.type.id,
        category_id: selection.type.category_id,
        quantity: selection.quantity,
        name: selection.type.name,
        item_ids: free.slice(0, selection.quantity),
      };
    });
  if (!items.length) {
    return null;
  }
  const deliverAt = start + deliverOffset * 60000;
  return {
    id: requestId(),
    items,
    item_count: items.reduce((total, item) => total + item.quantity, 0),
    deliver_day_offset: 0,
    deliver_offset: deliverOffset,
    notes: "",
    ref_id: `${deliverAt}|${items.map((item) => `${item.id}:${item.quantity}`).join("|")}`,
    conflict: false,
  };
}

export type AssetRequest = NonNullable<ReturnType<typeof buildAssetRequest>>;

export function assetBooking(request: AssetRequest, room: Space, roomName: string): LinkedBooking {
  const ids = request.items.flatMap((item) => item.item_ids);
  const names = request.items.map((item) => item.name).join(", ");
  return {
    booking_type: "asset-request",
    asset_id: ids[0],
    asset_ids: ids,
    asset_name: names,
    title: names,
    description: roomName,
    extension_data: { request_id: request.id, location_id: room.id, request },
  };
}
