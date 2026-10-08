/*
 * Shared PlaceOS UI and auth flow for both hosts: the Outlook add-in task pane (taskpane.ts) and the
 * Teams / Outlook / Microsoft 365 app bar personal tab (app.ts). Host-specific startup lives in those files.
 */

/* global document, console, window, HTMLElement, HTMLButtonElement, HTMLSelectElement */

import { AccountManager, type TokenResult } from "./authConfig";
import { loadAddinConfig, type AddinConfig } from "./addin-config";
import {
  exchangeForPlaceosToken,
  placeosFetch,
  PlaceosRequestError,
  type PlaceosRequestInit,
  refreshPlaceosToken,
  type PlaceosToken,
} from "./placeos-helper";
import { checkAccessTokenClaims, decodeAccessTokenClaims } from "./token-inspector";
import { renderRawTokenDevOnly } from "./dev-raw-token";
import { holdTokenForExpiryTest, renderFailureTestsDevOnly } from "./dev-failure-tests";
import { loadTodayOverview } from "./today-view";
import { favouriteRoomsView } from "./favourite-rooms";
import { loadOrganisation } from "./booking-settings";
import { initDeskBooking, openDeskBooking } from "./desk-booking-view";
import { initVisitorBooking, openVisitorBooking } from "./visitor-booking-view";
import { initRoomBooking, openRoomBooking } from "./room-booking-view";
import type { MeetingDraft } from "./room-booking-data";
import {
  checkStaffApi,
  getAuthority,
  getBookableRooms,
  getBuildings,
  getDeskSummary,
  zoneName,
} from "./placeos-data";

// Set by webpack DefinePlugin: true for development builds only.
declare const __DEV_TOOLS__: boolean;

let accountManager: AccountManager;
// Whether the host is in a mail context, where the Today view also shows the next meeting.
let showNextMeeting: () => boolean = () => false;
let favourites: ReturnType<typeof favouriteRoomsView> | undefined;
// Per-domain auth config, discovered from the PlaceOS domain on load.
let addinConfig: AddinConfig;

const el = (id: string) => document.getElementById(id) as HTMLElement;
const sideloadMsg = el("sideload-msg");
const appBody = el("app-body");
const todayView = el("today-view");
const bookView = el("book-view");
const deskView = el("desk-view");
const visitorView = el("visitor-view");
const viewToggle = el("viewToggle") as HTMLButtonElement;
const todayGreeting = el("todayGreeting");
const todayDate = el("todayDate");
const todayStatus = el("todayStatus");
const todaySignInButton = el("todaySignInButton") as HTMLButtonElement;
const todayCards = el("todayCards");
const todayRefreshButton = el("todayRefreshButton") as HTMLButtonElement;
const msAuthStatus = el("msAuthStatus");
const msAuthPath = el("msAuthPath");
const configSource = el("configSource");
const msAccount = el("msAccount");
const tokenSection = el("tokenSection");
const tokenSummary = el("tokenSummary");
const tokenClaims = el("tokenClaims");
const rawToken = el("rawToken");
const placeosSection = el("placeosSection");
const placeosUser = el("placeosUser");
const placeosUserDetail = el("placeosUserDetail");
const dataSection = el("dataSection");
const staffApiStatus = el("staffApiStatus");
const orgStatus = el("orgStatus");
const buildingSelect = el("buildingSelect") as HTMLSelectElement;
const roomsStatus = el("roomsStatus");
const roomsList = el("roomsList");
const desksStatus = el("desksStatus");
const loadDataButton = el("loadDataButton") as HTMLButtonElement;
const status = el("status");
const signInButton = el("signInButton") as HTMLButtonElement;
const getPlaceosTokenButton = el("getPlaceosTokenButton") as HTMLButtonElement;
const refreshPlaceosTokenButton = el("refreshPlaceosTokenButton") as HTMLButtonElement;

/**
 * Starts the UI once the host (Office or TeamsJS) is ready.
 * @param manager Auth for this host.
 * @param hostLabel Shown in the UI, e.g. "Outlook add-in" or "Teams app".
 */
export async function startApp(
  manager: AccountManager,
  hostLabel: string,
  options: {
    showNextMeeting?: () => boolean;
    // The meeting being composed in the host, which Book a room fills in instead of booking through PlaceOS.
    meetingDraft?: () => MeetingDraft | null;
    onMeetingChanged?: (listener: () => void) => void;
  } = {}
) {
  accountManager = manager;
  showNextMeeting = options.showNextMeeting ?? (() => false);
  sideloadMsg.style.display = "none";
  viewToggle.style.display = "";
  viewToggle.addEventListener("click", () =>
    showView(todayView.style.display === "none" ? "today" : "diagnostics")
  );
  showView(window.location.hash === "#diagnostics" ? "diagnostics" : "today");
  updateGreeting();
  todaySignInButton.addEventListener("click", signInAndExchange);
  todayRefreshButton.addEventListener("click", loadToday);
  favourites = favouriteRoomsView(placeosApi, options.meetingDraft ?? (() => null), (room) => {
    showView("book");
    openRoomBooking(room);
  });
  options.onMeetingChanged?.(() => {
    el("bookRoomActionLabel").textContent = options.meetingDraft?.()
      ? "Add a room to this meeting"
      : "Book a room";
    if (placeosToken && todayView.style.display !== "none") {
      favourites?.refresh();
    }
  });
  initRoomBooking(
    placeosApi,
    (booked) => {
      showView("today");
      if (booked) {
        loadToday();
      }
    },
    options.meetingDraft
  );
  if (options.meetingDraft?.()) {
    el("bookRoomActionLabel").textContent = "Add a room to this meeting";
  }
  el("bookRoomButton").addEventListener("click", () => {
    showView("book");
    openRoomBooking();
  });
  initDeskBooking(placeosApi, (booked) => {
    showView("today");
    if (booked) {
      loadToday();
    }
  });
  el("bookDeskButton").addEventListener("click", () => {
    showView("desk");
    openDeskBooking("desk");
  });
  el("bookParkingButton").addEventListener("click", () => {
    showView("desk");
    openDeskBooking("parking");
  });
  initVisitorBooking(placeosApi, (booked) => {
    showView("today");
    if (booked) {
      loadToday();
    }
  });
  el("inviteVisitorButton").addEventListener("click", () => {
    showView("visitor");
    openVisitorBooking();
  });
  signInButton.addEventListener("click", signInAndExchange);
  getPlaceosTokenButton.addEventListener("click", getPlaceosToken);
  refreshPlaceosTokenButton.addEventListener("click", refreshPlaceosTokenAction);
  loadDataButton.addEventListener("click", loadPlaceosData);
  buildingSelect.addEventListener("change", () => loadBuilding(buildingSelect.value));

  try {
    addinConfig = await loadAddinConfig();
  } catch (error) {
    setText(msAuthStatus, describeError(error), "error");
    setText(todayStatus, describeError(error), "error");
    getPlaceosTokenButton.disabled = true;
    return;
  }
  if (addinConfig.source === "domain") {
    setText(
      configSource,
      `${hostLabel}. Config from the PlaceOS domain (/auth/authority)`,
      "detail ok"
    );
  } else {
    setText(
      configSource,
      `${hostLabel}. Config: hardcoded dev fallback. Add "outlook_addin" to the domain's authority config.`,
      "detail error"
    );
  }
  await accountManager.initialize(addinConfig);
  await signInSilent();
}

function setText(element: HTMLElement, text: string, className = "") {
  element.innerText = text;
  element.className = className;
}

function showMicrosoftAuthResult(result: TokenResult) {
  setText(msAuthStatus, "✓ Authenticated", "ok");
  setText(msAuthPath, `Auth path: ${result.authPath}`, "detail");
  setText(msAccount, result.username || "-");
  microsoftUsername = result.username;
  signInButton.style.display = "none";
  showTokenClaims(result.accessToken);
}

let failureTestsShown = false;

/**
 * DEV ONLY: adds the Phase 4 failure-test panel once, after the first successful sign-in.
 */
function showFailureTestsDevOnly() {
  if (failureTestsShown) {
    return;
  }
  failureTestsShown = true;
  const section = document.createElement("section");
  section.className = "field";
  appBody.insertBefore(section, el("signOutButton"));
  renderFailureTestsDevOnly(section, {
    placeosClientId: addinConfig.placeosClientId,
    getEntraToken: async () =>
      (await accountManager.ssoGetAccessToken([addinConfig.scope])).accessToken,
    getGraphToken: async () => {
      const result = await accountManager.acquireTokenSilentOnly(["User.Read"]);
      // The dialog fallback always returns a PlaceOS-scoped token, so it can't supply a Graph token.
      return result && result.authPath !== "dialog" ? result.accessToken : null;
    },
  });
}

/**
 * Shows decoded claims. The raw token is never logged, and only development builds can reveal it.
 */
function showTokenClaims(accessToken: string) {
  tokenSection.style.display = "";
  tokenClaims.textContent = "";
  // typeof guard: a dev server started before the flag existed must not break sign-in.
  if (typeof __DEV_TOOLS__ !== "undefined" && __DEV_TOOLS__) {
    renderRawTokenDevOnly(rawToken, accessToken);
    holdTokenForExpiryTest(accessToken);
    showFailureTestsDevOnly();
  }
  let checks;
  try {
    checks = checkAccessTokenClaims(decodeAccessTokenClaims(accessToken), addinConfig);
  } catch (error) {
    setText(tokenSummary, `Could not decode token: ${error}`, "error");
    return;
  }

  const failed = checks.filter((check) => check.passed === false);
  if (failed.length === 0) {
    setText(tokenSummary, "✓ iss, aud, tid, scp and exp match the PlaceOS registration", "ok");
  } else {
    setText(tokenSummary, `✗ Mismatch: ${failed.map((check) => check.claim).join(", ")}`, "error");
  }

  for (const check of checks) {
    const dt = document.createElement("dt");
    const mark = check.passed === undefined ? "" : check.passed ? " ✓" : " ✗";
    dt.textContent = check.claim + mark;
    if (check.passed !== undefined) {
      dt.className = check.passed ? "ok" : "error";
    }
    const dd = document.createElement("dd");
    dd.textContent = check.value;
    tokenClaims.appendChild(dt);
    tokenClaims.appendChild(dd);
    if (check.passed === false && check.expected) {
      const expected = document.createElement("dd");
      expected.className = "expected";
      expected.textContent = `expected: ${check.expected}`;
      tokenClaims.appendChild(expected);
    }
  }
}

/**
 * On load, try to authenticate without any UI. If that fails, the user can sign in interactively.
 */
async function signInSilent() {
  setText(msAuthStatus, "Authenticating...");
  const naaSupported = accountManager.isNestedAppAuthSupported();
  try {
    const result = await accountManager.acquireTokenSilentOnly([addinConfig.scope]);
    if (result) {
      showMicrosoftAuthResult(result);
      // Go straight on to PlaceOS: exchange the token we already have, then load the data.
      await exchangeEntraToken(result);
      return;
    }
    setText(msAuthStatus, "Silent sign-in not available", "error");
  } catch (error) {
    console.error(error);
    setText(msAuthStatus, `Silent sign-in failed: ${error}`, "error");
  }
  setText(
    msAuthPath,
    `Nested app authentication supported: ${naaSupported ? "yes" : "no"}`,
    "detail"
  );
  signInButton.style.display = "";
  setText(todayStatus, "Sign in with your Microsoft account to see your day.");
  todaySignInButton.style.display = "";
}

/**
 * Runs the full token flow: silent, then popup, then Office dialog fallback.
 */
async function signInInteractive(): Promise<TokenResult | null> {
  signInButton.disabled = true;
  setText(msAuthStatus, "Signing in...");
  try {
    const result = await accountManager.ssoGetAccessToken([addinConfig.scope]);
    showMicrosoftAuthResult(result);
    return result;
  } catch (error) {
    console.error(error);
    setText(msAuthStatus, `Sign-in failed: ${error}`, "error");
    return null;
  } finally {
    signInButton.disabled = false;
  }
}

// PlaceOS token (and refresh token), in memory only.
let placeosToken: PlaceosToken | null = null;
// Microsoft account of the last Entra token, to compare with the PlaceOS user.
let microsoftUsername = "";

// Treat tokens this close to expiry as expired.
const expiryMarginMs = 60 * 1000;

type PlaceosUser = { id: string; name?: string; email?: string; login_name?: string };

function isUsable(token: PlaceosToken | null): token is PlaceosToken {
  return (
    !!token && (token.expiresAt === undefined || token.expiresAt - Date.now() > expiryMarginMs)
  );
}

/**
 * Gets a fresh Entra token (silently if possible) and exchanges it via auth.cr.
 */
async function exchangeWithEntra(): Promise<PlaceosToken> {
  const entraResult = await signInInteractive();
  if (!entraResult) {
    throw new Error("Could not get a Microsoft token.");
  }
  return exchangeForPlaceosToken(entraResult.accessToken, addinConfig.placeosClientId);
}

async function refreshWithPlaceos(): Promise<PlaceosToken> {
  if (!placeosToken?.refreshToken) {
    throw new Error("No PlaceOS refresh token yet. Get a PlaceOS token first.");
  }
  return refreshPlaceosToken(placeosToken.refreshToken, addinConfig.placeosClientId);
}

// The token renewal in progress, shared by every call that needs a token meanwhile (as ts-client shares one
// token request), so parallel calls don't each refresh or open a Microsoft sign-in.
let placeosTokenRenewal: Promise<PlaceosToken> | null = null;

/**
 * Returns a usable PlaceOS token: the cached one if not near expiry, else a refresh, else a new Entra exchange.
 */
function getValidPlaceosToken(): Promise<PlaceosToken> {
  if (isUsable(placeosToken)) {
    return Promise.resolve(placeosToken);
  }
  if (!placeosTokenRenewal) {
    placeosTokenRenewal = renewPlaceosToken().then(
      (token) => {
        placeosTokenRenewal = null;
        return token;
      },
      (error) => {
        placeosTokenRenewal = null;
        throw error;
      }
    );
  }
  return placeosTokenRenewal;
}

async function renewPlaceosToken(): Promise<PlaceosToken> {
  if (placeosToken?.refreshToken) {
    try {
      placeosToken = await refreshWithPlaceos();
      return placeosToken;
    } catch (error) {
      console.warn(`PlaceOS refresh failed, exchanging a new Entra token: ${error}`);
      // As ts-client: forget the refresh token only if PlaceOS rejected it, not on network or server errors.
      const rejected =
        error instanceof PlaceosRequestError && error.status >= 400 && error.status < 500;
      if (rejected && placeosToken) {
        placeosToken = { ...placeosToken, refreshToken: undefined };
      }
    }
  }
  placeosToken = await exchangeWithEntra();
  return placeosToken;
}

/**
 * Marks the access token as expired after PlaceOS rejected it, keeping its refresh token (ts-client
 * `invalidateToken`). Does nothing if another call has already replaced it.
 */
function invalidatePlaceosToken(accessToken: string) {
  if (placeosToken?.accessToken === accessToken) {
    placeosToken = { ...placeosToken, expiresAt: 0 };
  }
}

/**
 * Calls PlaceOS with a valid token, retrying once with a new token if PlaceOS returns 401.
 */
async function placeosApi<T>(path: string, init?: PlaceosRequestInit): Promise<T> {
  const { accessToken } = await getValidPlaceosToken();
  try {
    return await placeosFetch<T>(path, accessToken, init);
  } catch (error) {
    if (error instanceof PlaceosRequestError && error.status === 401) {
      invalidatePlaceosToken(accessToken);
      return placeosFetch<T>(path, (await getValidPlaceosToken()).accessToken, init);
    }
    throw error;
  }
}

/**
 * Makes an authenticated PlaceOS request to prove PlaceOS accepts the token.
 */
async function showPlaceosUser() {
  const user = await placeosApi<PlaceosUser>("/api/engine/v2/users/current");
  placeosUserName = user.name ?? "";
  updateGreeting();
  placeosSection.style.display = "";
  setText(placeosUser, `${user.name ?? "-"} (${user.email ?? "-"})`);
  const emailMatches = !!user.email && user.email.toLowerCase() === microsoftUsername.toLowerCase();
  setText(
    placeosUserDetail,
    `PlaceOS user ID: ${user.id}. Email ${emailMatches ? "matches" : "does not match"} the Microsoft account.`,
    emailMatches ? "detail ok" : "detail error"
  );
}

/**
 * Runs a token button action, timing it so exchange and refresh can be compared.
 */
async function runTokenAction(label: string, action: () => Promise<PlaceosToken>) {
  getPlaceosTokenButton.disabled = true;
  refreshPlaceosTokenButton.disabled = true;
  setText(status, `${label}...`);
  try {
    const start = Date.now();
    placeosToken = await action();
    const summary = `${label} took ${Date.now() - start} ms${
      placeosToken.expiresAt
        ? `, expires ${new Date(placeosToken.expiresAt).toLocaleTimeString()}`
        : ""
    }`;
    setText(status, `✓ ${summary}. Checking it with PlaceOS...`, "ok");
    await showPlaceosUser();
    setText(status, `✓ PlaceOS token accepted. ${summary}.`, "ok");
    todaySignInButton.style.display = "none";
    if (todayCards.style.display === "none") {
      loadToday();
    }
    if (appBody.style.display !== "none" && dataSection.style.display === "none") {
      loadPlaceosData();
    }
  } catch (error) {
    console.error(error);
    setText(status, `${error instanceof Error ? error.message : error}`, "error");
    if (!placeosToken) {
      setText(todayStatus, "Couldn't connect to PlaceOS. Open Diagnostics for details.", "error");
    }
  } finally {
    getPlaceosTokenButton.disabled = false;
    refreshPlaceosTokenButton.disabled = false;
    refreshPlaceosTokenButton.style.display = placeosToken?.refreshToken ? "" : "none";
  }
}

function getPlaceosToken() {
  return runTokenAction("Exchange (Microsoft token + auth.cr)", exchangeWithEntra);
}

/**
 * Exchanges an Entra token that was just acquired, without acquiring another one.
 */
function exchangeEntraToken(entraResult: TokenResult) {
  return runTokenAction("Exchange (automatic after sign-in)", () =>
    exchangeForPlaceosToken(entraResult.accessToken, addinConfig.placeosClientId)
  );
}

/**
 * The Sign in button (shown when silent sign-in fails): sign in interactively, then exchange automatically.
 */
async function signInAndExchange() {
  setText(todayStatus, "Signing you in...");
  const entraResult = await signInInteractive();
  if (entraResult) {
    await exchangeEntraToken(entraResult);
  } else {
    setText(
      todayStatus,
      "Sign-in didn't complete. Try again, or open Diagnostics for details.",
      "error"
    );
  }
}

/**
 * Loads the user's day (rooms, desks, parking, visitors) with the PlaceOS token.
 */
async function loadToday() {
  todayCards.style.display = "";
  todayRefreshButton.disabled = true;
  setText(todayStatus, "Loading your day...");
  updateGreeting();
  favourites?.refresh();
  try {
    await loadTodayOverview(
      placeosApi,
      {
        rooms: el("todayRooms"),
        desks: el("todayDesks"),
        parking: el("todayParking"),
        visitors: el("todayVisitors"),
        nextMeeting: showNextMeeting()
          ? { section: el("nextMeeting"), list: el("nextMeetingList") }
          : undefined,
      },
      settingsForZones
    );
    setText(
      todayStatus,
      `Updated ${new Date().toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}`
    );
  } finally {
    todayRefreshButton.disabled = false;
  }
}

/**
 * Settings of the building in a booking's zones (for check-in options), else of the default building, as
 * Workplace reads the active building's settings; null if there are none.
 */
async function settingsForZones(zones: string[]) {
  try {
    const org = await loadOrganisation(placeosApi);
    const building =
      org.buildings.find((b) => zones.includes(b.id)) || (await org.defaultBuilding());
    return building ? await org.settings(building.id) : null;
  } catch {
    return null;
  }
}

type View = "today" | "diagnostics" | "book" | "desk" | "visitor";

function showView(view: View) {
  todayView.style.display = view === "today" ? "flex" : "none";
  bookView.style.display = view === "book" ? "flex" : "none";
  deskView.style.display = view === "desk" ? "flex" : "none";
  visitorView.style.display = view === "visitor" ? "flex" : "none";
  appBody.style.display = view === "diagnostics" ? "flex" : "none";
  viewToggle.textContent = view === "today" ? "Diagnostics" : "Back to today";
  if (view === "today" && placeosToken) {
    favourites?.refresh();
  }
  // The diagnostics data checks only run when someone looks at them.
  if (view === "diagnostics" && placeosToken && dataSection.style.display === "none") {
    loadPlaceosData();
  }
}

// PlaceOS display name, for the greeting.
let placeosUserName = "";

function updateGreeting() {
  const hour = new Date().getHours();
  const timeOfDay = hour < 12 ? "Good morning" : hour < 18 ? "Good afternoon" : "Good evening";
  const firstName = placeosUserName.split(" ")[0];
  todayGreeting.textContent = firstName ? `${timeOfDay}, ${firstName}` : timeOfDay;
  todayDate.textContent = new Date().toLocaleDateString(undefined, {
    weekday: "long",
    day: "numeric",
    month: "long",
  });
}

function refreshPlaceosTokenAction() {
  return runTokenAction("Refresh (PlaceOS refresh token)", refreshWithPlaceos);
}

function describeError(error: unknown): string {
  if (error instanceof PlaceosRequestError) {
    return `${error.message}${error.body ? `: ${error.body.slice(0, 200)}` : ""}`;
  }
  return error instanceof Error ? error.message : `${error}`;
}

/**
 * Phase 5: read-only PlaceOS and Staff API calls with the PlaceOS token. Each row reports its own errors.
 */
async function loadPlaceosData() {
  dataSection.style.display = "";
  loadDataButton.disabled = true;
  try {
    await Promise.all([loadStaffApiCheck(), loadOrgAndBuildings()]);
  } finally {
    loadDataButton.disabled = false;
  }
}

async function loadStaffApiCheck() {
  setText(staffApiStatus, "Staff API: checking...");
  try {
    await checkStaffApi(placeosApi);
    setText(staffApiStatus, "✓ Staff API accepts the PlaceOS token", "ok");
  } catch (error) {
    setText(staffApiStatus, `✗ Staff API: ${describeError(error)}`, "error");
  }
}

async function loadOrgAndBuildings() {
  let orgZoneId: string | undefined;
  setText(orgStatus, "Organisation: loading...");
  try {
    const authority = await getAuthority(placeosApi);
    orgZoneId = authority.config?.org_zone;
    const domain = authority.name.includes(authority.domain) ? "" : ` (${authority.domain})`;
    setText(
      orgStatus,
      `Organisation: ${authority.name}${domain}${orgZoneId ? `, org zone ${orgZoneId}` : ", no org_zone configured"}`
    );
  } catch (error) {
    setText(orgStatus, `✗ Organisation: ${describeError(error)}`, "error");
  }

  buildingSelect.disabled = true;
  buildingSelect.textContent = "";
  try {
    let buildings = await getBuildings(placeosApi, orgZoneId);
    let note = "";
    if (buildings.length === 0 && orgZoneId) {
      // Buildings may not be direct children of the org zone; fall back to any building zone visible to the user.
      buildings = await getBuildings(placeosApi);
      note = buildings.length
        ? ` (none are children of ${orgZoneId}; showing all building zones)`
        : "";
    }
    for (const building of buildings) {
      const option = document.createElement("option");
      option.value = building.id;
      option.textContent = zoneName(building);
      buildingSelect.appendChild(option);
    }
    if (buildings.length === 0) {
      setText(
        roomsStatus,
        `No zones tagged "building" visible to this user${orgZoneId ? `, under ${orgZoneId} or anywhere else` : ""}`,
        "error"
      );
      setText(desksStatus, "");
      return;
    }
    buildingSelect.disabled = false;
    if (note) {
      setText(orgStatus, orgStatus.innerText + note);
    }
    await loadBuilding(buildings[0].id);
  } catch (error) {
    setText(roomsStatus, `✗ Buildings: ${describeError(error)}`, "error");
    setText(desksStatus, "");
  }
}

async function loadBuilding(buildingId: string) {
  await Promise.all([loadRooms(buildingId), loadDesks(buildingId)]);
}

async function loadRooms(buildingId: string) {
  const maxShown = 10;
  setText(roomsStatus, "Rooms: loading...");
  roomsList.textContent = "";
  try {
    const rooms = await getBookableRooms(placeosApi, buildingId);
    setText(roomsStatus, `Rooms: ${rooms.length} bookable`, rooms.length ? "ok" : "");
    for (const room of rooms.slice(0, maxShown)) {
      const item = document.createElement("li");
      item.textContent = `${zoneName(room)}${room.capacity ? ` (capacity ${room.capacity})` : ""}`;
      roomsList.appendChild(item);
    }
    if (rooms.length > maxShown) {
      const more = document.createElement("li");
      more.textContent = `+${rooms.length - maxShown} more`;
      roomsList.appendChild(more);
    }
  } catch (error) {
    setText(roomsStatus, `✗ Rooms: ${describeError(error)}`, "error");
  }
}

async function loadDesks(buildingId: string) {
  setText(desksStatus, "Desks: loading...");
  try {
    const summary = await getDeskSummary(placeosApi, buildingId);
    setText(
      desksStatus,
      `Desks: ${summary.desks} on ${summary.levels} level(s), ${summary.bookedToday} booked today`,
      summary.desks ? "ok" : ""
    );
  } catch (error) {
    setText(desksStatus, `✗ Desks: ${describeError(error)}`, "error");
  }
}
