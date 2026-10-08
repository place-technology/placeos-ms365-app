/*
 * PlaceOS serves the add-in without Cache-Control headers, so Outlook's webview can keep serving an old
 * taskpane.html / app.html (and the bundles it names) from cache after a deploy. This fetches the page again,
 * bypassing the cache, and if it comes from a newer build, reloads at a URL with ?v=<build> that the cache can't have.
 */

/* global fetch, window, URL */

// Set by webpack DefinePlugin: true for development builds only.
declare const __DEV_TOOLS__: boolean;

const BUILD_META = /<meta[^>]+name="?placeos-build"?[^>]+content="([^"]+)"/i;

/** Reloads the page if the server has a newer build than the one running. Never throws. */
export async function reloadIfOutdated(currentBuild: string): Promise<void> {
  if (__DEV_TOOLS__) return; // The dev server always serves the latest build.
  try {
    const response = await fetch(window.location.pathname, {
      cache: "no-store",
      credentials: "same-origin",
    });
    if (!response.ok) return;
    const latest = BUILD_META.exec(await response.text())?.[1];
    if (!latest || latest === currentBuild) return;
    const url = new URL(window.location.href);
    // Already tried this build: don't loop if the reload still came from cache.
    if (url.searchParams.get("v") === latest) return;
    url.searchParams.set("v", latest);
    window.location.replace(url.toString());
  } catch {
    // Offline or blocked: keep running the cached build.
  }
}
