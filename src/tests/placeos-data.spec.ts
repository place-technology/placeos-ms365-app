import { deskTypeRoutes, fakeApi } from "./fake-api";

async function loadModule() {
  // the desk asset type id is cached per module
  vi.resetModules();
  return import("../taskpane/placeos-data");
}

describe("getDeskSummary", () => {
  it("counts the building's desk assets, their levels and today's bookings", async () => {
    const { getDeskSummary } = await loadModule();
    const { api, calls } = fakeApi({
      ...deskTypeRoutes,
      "/api/engine/v2/assets?zones=bld-1": [
        { id: "asset-1", zone_id: "lvl-1" },
        { id: "asset-2", zone_id: "lvl-1" },
        { id: "asset-3", zone_id: "lvl-2" },
      ],
      "/api/staff/v1/bookings/booked": ["asset-1", "asset-1"],
    });

    expect(await getDeskSummary(api, "bld-1")).toEqual({ levels: 2, desks: 3, bookedToday: 1 });
    expect(calls.some((path) => path.includes("/metadata/"))).toBe(false);
  });
});
