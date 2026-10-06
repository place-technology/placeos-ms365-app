/*
 * Catering for room bookings, as in the legacy Outlook add-in (libs/catering, libs/assets catering-assets.fn).
 * The menu is PlaceOS Assets: category "_CATERING_", asset types "CATERING:<caterer>", one asset per item.
 * Orders are saved on the event (extension_data.catering) and as "catering-order" bookings linked to it.
 */

import { getMetadata, type AppSettings } from "./booking-settings";
import { allPages, resultList as list, type PlaceosApi } from "./placeos-data";
import { stringToMinutes, type LinkedBooking, type Space } from "./room-booking-data";

export type CateringOption = {
  id: string;
  name: string;
  group?: string;
  multiple?: boolean;
  unit_price?: number;
};

export type CateringItem = {
  id: string;
  name: string;
  caterer: string;
  category: string;
  description?: string;
  // Cents.
  unit_price: number;
  tags: string[];
  options: CateringOption[];
  images: string[];
};

/** A menu item the user picked: quantity and chosen option ids. */
export type CateringSelection = { item: CateringItem; quantity: number; optionIds: string[] };

export type CateringSettings = { requireNotes: boolean; chargeCodes: string[]; currency: string };

type AssetCategory = { id: string; name: string; created_at?: number };
type AssetType = { id: string; name: string };
type Asset = {
  id: string;
  name?: string;
  identifier?: string;
  images?: string[];
  other_data?: {
    category?: string;
    description?: string;
    unit_price?: number;
    options?: CateringOption[];
    tags?: string[];
    images?: string[];
  };
};

/** Catering is offered when app.events.catering_enabled or app.events.has_catering is set. */
export function cateringEnabled(settings: AppSettings): boolean {
  return !!settings.get("events.catering_enabled") || !!settings.get("events.has_catering");
}

export async function getCateringSettings(
  api: PlaceosApi,
  buildingId: string,
  settings: AppSettings,
  buildingCurrency?: string
): Promise<CateringSettings> {
  const details = await getMetadata<{ require_notes?: boolean; charge_codes?: string[] }>(
    api,
    buildingId,
    "catering-settings"
  );
  return {
    requireNotes: !!settings.get("events.catering_notes_required") || !!details?.require_notes,
    chargeCodes: Array.isArray(details?.charge_codes) ? details.charge_codes : [],
    currency: settings.get<string>("currency") || buildingCurrency || "USD",
  };
}

export type AttachedRule = { name: string; rules: [string, string | number][] };

/**
 * Lead-time rules for catering items (metadata "catering_config") and asset types ("assets_config").
 * A ruleset applies when its name is one of the item's names (category, tags, type) or the room's zones, or
 * "*". Every condition of an applying ruleset must pass. Ported from libs/catering utilities.
 */
export function attachedRulesAllow(
  names: string[],
  rules: AttachedRule[] | null,
  room: Space,
  start: number,
  duration: number
): boolean {
  const now = Date.now();
  for (const rule of Array.isArray(rules) ? rules : []) {
    const applies =
      names.includes(rule.name) || (room.zones || []).includes(rule.name) || rule.name === "*";
    if (!applies || !Array.isArray(rule.rules)) {
      continue;
    }
    let matches = 0;
    for (const [type, value] of rule.rules) {
      const minutes = typeof value === "string" ? stringToMinutes(value) : +value * 60;
      const atHour = new Date(start);
      atHour.setHours(minutes / 60);
      switch (type) {
        case "is_before":
          matches += now < start - minutes * 60000 ? 1 : 0;
          break;
        case "within_hours":
          matches += now > start - minutes * 60000 ? 1 : 0;
          break;
        case "after_hour":
          matches += start > atHour.getTime() ? 1 : 0;
          break;
        case "before_hour":
          matches += start < atHour.getTime() ? 1 : 0;
          break;
        case "min_length":
          matches += duration >= minutes ? 1 : 0;
          break;
        case "max_length":
          matches += duration <= minutes ? 1 : 0;
          break;
        case "visitor_type":
          // Room bookings from the add-in have no visitor type.
          break;
        default:
          matches += 1;
      }
    }
    if (matches < rule.rules.length) {
      return false;
    }
  }
  return true;
}

/** Menu items available in the zone (the room's building) for this booking. */
export async function getCateringMenu(
  api: PlaceosApi,
  zoneId: string,
  room: Space,
  start: number,
  duration: number
): Promise<CateringItem[]> {
  const categories = list(
    await api<AssetCategory[]>("/api/engine/v2/asset_categories?hidden=true&limit=500", allPages)
  )
    .filter((category) => category.name === "_CATERING_")
    .sort((a, b) => (a.created_at || 0) - (b.created_at || 0));
  if (!categories.length) {
    return [];
  }
  const types = list(
    await api<AssetType[]>(
      `/api/engine/v2/asset_types?category_id=${encodeURIComponent(categories[0].id)}&limit=500`,
      allPages
    )
  ).filter((type) => type.name.startsWith("CATERING:"));
  const [assetLists, rules] = await Promise.all([
    Promise.all(
      types.map((type) =>
        api<Asset[]>(
          `/api/engine/v2/assets?zone_id=${encodeURIComponent(zoneId)}` +
            `&type_id=${encodeURIComponent(type.id)}&limit=500`,
          allPages
        ).then((assets) => ({ type, assets: list(assets) }))
      )
    ),
    getMetadata<AttachedRule[]>(api, zoneId, "catering_config"),
  ]);
  const items: CateringItem[] = [];
  for (const { type, assets } of assetLists) {
    const caterer = type.name.replace("CATERING:", "");
    for (const asset of assets) {
      const data = asset.other_data || {};
      items.push({
        id: asset.id,
        name: asset.name || asset.identifier || asset.id,
        caterer: caterer === "_STANDALONE_" ? "standalone" : caterer,
        category: data.category || "Other",
        description: data.description,
        unit_price: data.unit_price || 0,
        tags: data.tags || [],
        options: data.options || [],
        images: asset.images?.length ? asset.images : data.images || [],
      });
    }
  }
  return items
    .filter((item) =>
      attachedRulesAllow([item.category, ...item.tags], rules, room, start, duration)
    )
    .sort((a, b) => a.name.localeCompare(b.name));
}

/** The caterer every menu item must come from (app.catering_provider), if the building has one. */
export function cateringProvider(settings: AppSettings): string {
  return settings.get<string>("catering_provider") || "";
}

/** Menu filters, as in Workplace's catering order modal (CateringOrderStateService). */
export type CateringFilters = {
  search: string;
  // "" = all caterers.
  caterer: string;
  categories: string[];
  // Items must have every selected tag.
  tags: string[];
  favouritesOnly: boolean;
};

export function filterCateringMenu(
  items: CateringItem[],
  filters: CateringFilters,
  favourites: string[]
): CateringItem[] {
  const search = filters.search.trim().toLowerCase();
  return items.filter(
    (item) =>
      (!search ||
        item.name.toLowerCase().includes(search) ||
        (item.description || "").toLowerCase().includes(search)) &&
      (!filters.caterer || item.caterer === filters.caterer) &&
      (!filters.categories.length || filters.categories.includes(item.category)) &&
      filters.tags.every((tag) => item.tags.includes(tag)) &&
      (!filters.favouritesOnly || favourites.includes(item.id))
  );
}

/** Options grouped by group name. A group is multiple choice if any option in it is. */
export function optionGroups(item: CateringItem) {
  const groups = new Map<string, { name: string; multiple: boolean; options: CateringOption[] }>();
  for (const option of item.options) {
    const name = option.group || "Options";
    const group = groups.get(name) || { name, multiple: false, options: [] };
    group.multiple = group.multiple || !!option.multiple;
    group.options.push(option);
    groups.set(name, group);
  }
  return Array.from(groups.values());
}

export function unitPriceWithOptions(selection: CateringSelection): number {
  return (
    selection.item.unit_price +
    selection.item.options
      .filter((option) => selection.optionIds.includes(option.id))
      .reduce((total, option) => total + (option.unit_price || 0), 0)
  );
}

export function formatPrice(cents: number, currency: string): string {
  try {
    return (cents / 100).toLocaleString(undefined, { style: "currency", currency });
  } catch {
    return `${(cents / 100).toFixed(2)} ${currency}`;
  }
}

const orderId = () => `order-${Math.floor(1000000 + Math.random() * 8999999)}`;

/** CateringOrder JSON (libs/common catering.class toJSON): one order per caterer, delivered to the room. */
export function buildCateringOrders(
  selections: CateringSelection[],
  room: Space,
  start: number,
  deliverOffset: number,
  notes: string,
  chargeCode: string
) {
  const byCaterer = new Map<string, CateringSelection[]>();
  for (const selection of selections.filter((s) => s.quantity > 0)) {
    byCaterer.set(selection.item.caterer, [
      ...(byCaterer.get(selection.item.caterer) || []),
      selection,
    ]);
  }
  return Array.from(byCaterer.entries()).map(([caterer, chosen]) => {
    const items = chosen.map((selection) => {
      const unitWithOptions = unitPriceWithOptions(selection);
      return {
        id: selection.item.id,
        name: selection.item.name,
        category: selection.item.category,
        caterer,
        description: selection.item.description || "",
        unit_price: selection.item.unit_price,
        quantity: selection.quantity,
        options: selection.item.options.map((option) => ({
          ...option,
          active: selection.optionIds.includes(option.id),
        })),
        tags: selection.item.tags,
        images: selection.item.images,
        unit_price_with_options: unitWithOptions,
        total_cost: unitWithOptions * selection.quantity,
      };
    });
    return {
      id: orderId(),
      system_id: room.id,
      items,
      item_count: items.reduce((total, item) => total + item.quantity, 0),
      total_cost: items.reduce((total, item) => total + item.total_cost, 0),
      charge_code: chargeCode || undefined,
      deliver_offset: deliverOffset,
      deliver_day_offset: 0,
      notes: notes || undefined,
      caterer,
      deliver_at_time: start + deliverOffset * 60000,
      status: "accepted",
    };
  });
}

export type CateringOrder = ReturnType<typeof buildCateringOrders>[number];

export function cateringBookings(
  orders: CateringOrder[],
  room: Space,
  roomName: string
): LinkedBooking[] {
  return orders.map((order) => ({
    booking_type: "catering-order",
    asset_id: room.id,
    asset_ids: [room.id],
    asset_name: roomName,
    extension_data: { name: roomName, location_id: room.id, details: order },
  }));
}
