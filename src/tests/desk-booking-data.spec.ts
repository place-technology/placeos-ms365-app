import type { Zone } from "../taskpane/booking-settings";
import { deskTypeRoutes, fakeApi } from "./fake-api";

async function loadModule() {
  // the desk asset type id is cached per module
  vi.resetModules();
  return import("../taskpane/desk-booking-data");
}

const level1: Zone = { id: "lvl-1", name: "Level 1", parent_id: "bld-1" } as Zone;
const level2: Zone = { id: "lvl-2", name: "Level 2", parent_id: "bld-1" } as Zone;

const assets = [
  {
    id: "asset-1",
    identifier: "6.06.01",
    map_id: "desk-6.06.1",
    zone_id: "lvl-1",
    zones: ["org-1", "bld-1", "lvl-1"],
    bookable: true,
    place_groups: ["staff"],
    client_ids: { floorsense_id: "fs-1" },
    other_data: { tags: ["quiet", 3], homebase: "engineering" },
  },
  {
    // migrated desk whose original id differs from its map id
    id: "asset-2",
    identifier: "6.06.02",
    map_id: "table-2",
    zone_id: "lvl-1",
    zones: ["org-1", "bld-1", "lvl-1"],
    bookable: true,
    client_ids: { place_id: "desk-6.06.2" },
  },
  {
    // on a level that wasn't asked for
    id: "asset-3",
    identifier: "7.01",
    zone_id: "lvl-7",
    zones: ["org-1", "bld-1", "lvl-7"],
    bookable: true,
  },
];

describe("getDesks", () => {
  it("loads the desks of each building with one zones query", async () => {
    const { getDesks } = await loadModule();
    const { api, calls } = fakeApi({
      ...deskTypeRoutes,
      "/api/engine/v2/assets?zones=bld-1": assets,
    });

    const desks = await getDesks(api, [level1, level2]);

    expect(calls.filter((path) => path.startsWith("/api/engine/v2/assets?"))).toEqual([
      "/api/engine/v2/assets?zones=bld-1&type_id=type-desks&limit=500",
    ]);
    expect(desks.map((desk) => desk.id)).toEqual(["asset-1", "asset-2"]);
    expect(desks[0]).toMatchObject({
      name: "6.06.01",
      map_id: "desk-6.06.1",
      level: level1,
      groups: ["staff"],
      tags: ["quiet"],
    });
    expect(desks[0].place_id).toBeUndefined();
    expect(desks[1]).toMatchObject({ map_id: "table-2", place_id: "desk-6.06.2", tags: [] });
  });

  it("uses the asset id as the map id when there is none", async () => {
    const { getDesks } = await loadModule();
    const { api } = fakeApi({
      ...deskTypeRoutes,
      "/api/engine/v2/assets?zones=bld-1": [{ id: "asset-9", zone_id: "lvl-1", bookable: true }],
    });

    const [desk] = await getDesks(api, [level1]);
    expect(desk).toMatchObject({ id: "asset-9", name: "asset-9", map_id: "asset-9" });
  });

  it("has no desks when the domain has no desk asset type", async () => {
    const { getDesks } = await loadModule();
    const { api, calls } = fakeApi({ "/api/engine/v2/asset_categories?hidden=true": [] });

    expect(await getDesks(api, [level1])).toEqual([]);
    expect(calls.some((path) => path.startsWith("/api/engine/v2/assets?"))).toBe(false);
  });
});

describe("bookableDesks", () => {
  const request = {
    start: Date.now() + 60 * 60 * 1000,
    duration: 60,
    allDay: false,
    buildingId: "bld-1",
  };

  it("hides desks named by booking rules with their original desk id", async () => {
    const { bookableDesks, getDesks } = await loadModule();
    const { api } = fakeApi({ ...deskTypeRoutes, "/api/engine/v2/assets?zones=bld-1": assets });
    const desks = await getDesks(api, [level1]);
    const rule = (resource_ids: string[]) =>
      new Map([["bld-1", [{ zone: "*", rules: { hidden: true }, conditions: { resource_ids } }]]]);

    // the rule lists the place_id of asset-2
    expect(
      bookableDesks(desks, rule(["desk-6.06.2"]), request, ["staff"]).map((desk) => desk.id)
    ).toEqual(["asset-1"]);
    // the map id of asset-1
    expect(
      bookableDesks(desks, rule(["desk-6.06.1"]), request, ["staff"]).map((desk) => desk.id)
    ).toEqual(["asset-2"]);
  });

  it("matches booking rule tags from other_data", async () => {
    const { bookableDesks, getDesks } = await loadModule();
    const { api } = fakeApi({ ...deskTypeRoutes, "/api/engine/v2/assets?zones=bld-1": assets });
    const desks = await getDesks(api, [level1]);
    const rules = new Map([
      ["bld-1", [{ zone: "*", rules: { hidden: true }, conditions: { tags: ["quiet"] } }]],
    ]);

    expect(bookableDesks(desks, rules, request, ["staff"]).map((desk) => desk.id)).toEqual([
      "asset-2",
    ]);
  });

  it("only offers desks with groups to members of those groups", async () => {
    const { bookableDesks, getDesks } = await loadModule();
    const { api } = fakeApi({ ...deskTypeRoutes, "/api/engine/v2/assets?zones=bld-1": assets });
    const desks = await getDesks(api, [level1]);

    expect(bookableDesks(desks, new Map(), request, []).map((desk) => desk.id)).toEqual([
      "asset-2",
    ]);
  });
});
