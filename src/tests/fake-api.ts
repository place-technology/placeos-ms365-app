import type { PlaceosApi } from "../taskpane/placeos-data";

/** A PlaceosApi that answers each call from the first route whose path prefix matches, and records the calls. */
export function fakeApi(routes: Record<string, unknown>) {
  const calls: string[] = [];
  const api = (async (path: string) => {
    calls.push(path);
    const route = Object.keys(routes)
      .sort((a, b) => b.length - a.length)
      .find((prefix) => path.startsWith(prefix));
    if (!route) {
      throw new Error(`unexpected request ${path}`);
    }
    return routes[route];
  }) as PlaceosApi;
  return { api, calls };
}

/** The `_DESKS_` category and type routes. */
export const deskTypeRoutes = {
  "/api/engine/v2/asset_categories?hidden=true": [{ id: "cat-desks", name: "_DESKS_" }],
  "/api/engine/v2/asset_types?category_id=cat-desks": [{ id: "type-desks", name: "_DESKS_" }],
};
