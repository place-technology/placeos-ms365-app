/*
 * The Book a desk and Book parking view: date and time → free desks or parking spaces (list or floor plan map,
 * with filters) → confirm → done. Mirrors PlaceOS Workplace's desk flow with app.desks.use_assets and its
 * parking flow; the data side is in desk-booking-data.ts. Both use the same page elements (the desk* ids),
 * with the text set for the kind being booked.
 */

/* global document, console, window, HTMLElement, HTMLInputElement, HTMLSelectElement, HTMLButtonElement */

import {
  isParkingLevel,
  loadOrganisation,
  zoneLabel,
  type AppSettings,
  type Organisation,
  type Zone,
} from "./booking-settings";
import {
  blockingAssignedDesk,
  bookableDesks,
  buildDeskBooking,
  checkBeforeBooking,
  createDeskBooking,
  deskAllDayAllowed,
  deskDurationLimits,
  deskDurationStep,
  getBookedDeskIds,
  getDeskRules,
  getDesks,
  getParkingUser,
  kinds,
  kindSetting,
  type Desk,
  type DeskRequest,
  type ParkingUser,
  type ResourceKind,
} from "./desk-booking-data";
import {
  defaultMapColours,
  FloorMap,
  loadMapSvg,
  type MapRoom,
  type MapRoomStatus,
} from "./floor-map";
import type { PlaceosApi } from "./placeos-data";
import { resultCard, selectedResult } from "./result-card";
import { PlaceosRequestError } from "./placeos-helper";
import { allDayPeriod, checkBookableHours, dateValue, isFullDay } from "./room-booking-data";

const el = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
const input = (id: string) => el<HTMLInputElement>(id);
const select = (id: string) => el<HTMLSelectElement>(id);
const show = (element: HTMLElement, visible: boolean) =>
  (element.style.display = visible ? "" : "none");

const status = el("deskStatus");
const steps = {
  assigned: el("deskAssigned"),
  form: el("deskForm"),
  desks: el("deskDesks"),
  confirm: el("deskConfirm"),
  done: el("deskDone"),
};
type Step = keyof typeof steps;

const minute = 60000;
// Start times are offered every 15 minutes.
const startStep = 15;

let api: PlaceosApi;
let onClose: (booked: boolean) => void = () => undefined;
let org: Organisation | null = null;
let settings: AppSettings;
let buildingId = "";
// What's being booked, set when the view opens.
let kind: ResourceKind = "desk";
// Parking: the user's parking user entry in the building, if any.
let parkingUser: ParkingUser | null = null;

let request: DeskRequest | null = null;
// Desks step. candidates are the free desks the user may book; searchedDesks also has the booked ones and
// those the user can't book (for the map). bookedIds are the desks booked at this time, by anyone.
let searchedDesks: Desk[] = [];
let bookedIds = new Set<string>();
let candidates: Desk[] = [];
let deskView: "list" | "map" = "list";
// Set once the user picks List, so later searches don't switch back to the map.
let listChosen = false;
let floorMap: FloorMap | null = null;
let mapLevels: Zone[] = [];
let shownMapUrl = "";
let mapLoadSeq = 0;
let mapSelection: Desk | null = null;
let chosen: Desk | null = null;
let submitting = false;
// Desks per searched set of buildings, loaded once each time the view opens.
let deskCache = new Map<string, Promise<Desk[]>>();
// The user's assigned desk or space, when it stops them booking another, and its own map.
let assignedDesk: Desk | null = null;
let assignedMap: FloorMap | null = null;

const noun = () => kinds[kind].noun;
const nouns = () => `${noun()}s`;
const parking = () => kind === "parking";

function showStep(step: Step) {
  (Object.keys(steps) as Step[]).forEach((key) => show(steps[key], key === step));
  // The building can be changed on the form, or when the building doesn't let the user book.
  show(
    el("deskBuildingField"),
    (step === "form" || step === "assigned") && (org?.buildings.length || 0) > 1
  );
  window.scrollTo(0, 0);
}

/** Sets the page's text for the kind being booked. */
function applyLabels() {
  const parkingKind = parking();
  setText(el("deskTitle"), parkingKind ? "Book parking" : "Book a desk");
  setText(el("deskFindButton"), parkingKind ? "Find parking" : "Find desk");
  setText(el("deskSubmitButton"), parkingKind ? "Book parking space" : "Book desk");
  setText(el("deskOtherButton"), `Choose another ${noun()}`);
  setText(el("deskAnotherButton"), `Book another ${noun()}`);
  el("deskViewToggle").setAttribute("aria-label", `Show ${nouns()} as`);
  show(el("deskPlateField"), parkingKind);
}

function setText(element: HTMLElement, text: string, className?: string) {
  element.textContent = text;
  if (className !== undefined) {
    element.className = className;
  }
}

function describeError(error: unknown): string {
  if (error instanceof PlaceosRequestError) {
    if (error.status === 409) {
      return `That ${noun()} was booked by someone else in the meantime. Choose another ${noun()}.`;
    }
    if (error.status === 401 || error.status === 403) {
      return "PlaceOS didn't allow this booking.";
    }
    return `PlaceOS couldn't complete this (${error.status}). Try again.`;
  }
  return error instanceof Error ? error.message : `${error}`;
}

const formatTime = (ms: number) =>
  new Date(ms).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
const formatDate = (ms: number) =>
  new Date(ms).toLocaleDateString(undefined, { weekday: "long", day: "numeric", month: "long" });
const formatDuration = (minutes: number) => {
  const hours = Math.floor(minutes / 60);
  const rest = minutes % 60;
  return [hours ? `${hours} h` : "", rest ? `${rest} min` : ""].filter(Boolean).join(" ");
};

/** Local midnight of a "YYYY-MM-DD" value, or NaN. */
function parseDate(value: string): number {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value.trim());
  return match ? new Date(+match[1], +match[2] - 1, +match[3]).getTime() : NaN;
}

/** "Thursday 1 October, 9:00 am – 5:00 pm", or "Thursday 1 October, all day". */
function whenText({ start, duration, allDay }: DeskRequest): string {
  const times =
    allDay && isFullDay(start, duration)
      ? "all day"
      : `${formatTime(start)} – ${formatTime(start + duration * minute)}${allDay ? " (all day)" : ""}`;
  return `${formatDate(start)}, ${times}`;
}

function option(parent: HTMLSelectElement, value: string, text: string) {
  const item = document.createElement("option");
  item.value = value;
  item.textContent = text;
  parent.appendChild(item);
}

function line(parent: HTMLElement, text: string, className = "item-detail") {
  if (!text) {
    return;
  }
  const div = document.createElement("div");
  div.className = className;
  div.textContent = text;
  parent.appendChild(div);
}

// ---------------------------------------------------------------------------------------------------------
// Setup
// ---------------------------------------------------------------------------------------------------------

/** Wires up the view once. onDone is called when the user leaves it (booked = a booking was made). */
export function initDeskBooking(placeosApi: PlaceosApi, onDone: (booked: boolean) => void) {
  api = placeosApi;
  onClose = onDone;
  el("deskForm").addEventListener("submit", (event) => {
    event.preventDefault();
    findDesks();
  });
  el("deskCancelButton").addEventListener("click", () => onClose(false));
  select("deskBuilding").addEventListener("change", () =>
    changeBuilding(select("deskBuilding").value)
  );
  input("deskDate").addEventListener("change", fillTimes);
  input("deskAllDay").addEventListener("change", updateAllDay);
  select("deskStart").addEventListener("change", () => fillEndTimes());

  el("deskFiltersToggle").addEventListener("click", () =>
    show(el("deskFilters"), el("deskFilters").style.display === "none")
  );
  el("deskBackToForm").addEventListener("click", () => showStep("form"));
  select("deskLevel").addEventListener("change", renderDesks);
  el("deskListViewButton").addEventListener("click", () => {
    listChosen = true;
    setDeskView("list");
  });
  el("deskMapViewButton").addEventListener("click", () => {
    listChosen = false;
    setDeskView("map");
  });
  select("deskMapLevel").addEventListener("change", () => {
    mapSelection = null;
    renderMap();
  });
  floorMap = new FloorMap(el("deskMapCanvas"));
  floorMap.onSelect = selectOnMap;
  el("deskMapZoomIn").addEventListener("click", () => floorMap?.zoomBy(1.5));
  el("deskMapZoomOut").addEventListener("click", () => floorMap?.zoomBy(1 / 1.5));
  el("deskMapReset").addEventListener("click", () => floorMap?.reset());

  assignedMap = new FloorMap(el("deskAssignedMapCanvas"));
  el("deskAssignedMapButton").addEventListener("click", () => showAssignedMap());
  el("deskAssignedBackButton").addEventListener("click", () => onClose(false));
  el("deskAssignedMapZoomIn").addEventListener("click", () => assignedMap?.zoomBy(1.5));
  el("deskAssignedMapZoomOut").addEventListener("click", () => assignedMap?.zoomBy(1 / 1.5));
  el("deskAssignedMapReset").addEventListener("click", () => assignedMap?.reset());

  el("deskOtherButton").addEventListener("click", () => showStep("desks"));
  el("deskSubmitButton").addEventListener("click", () => submit());
  el("deskDoneButton").addEventListener("click", () => onClose(true));
  el("deskAnotherButton").addEventListener("click", () => {
    resetForm();
    showStep("form");
  });
}

/** Opens the view at the form for desks or parking. Loads the org and settings the first time. */
export async function openDeskBooking(newKind: ResourceKind = "desk") {
  if (newKind !== kind) {
    kind = newKind;
    request = null;
    input("deskPlate").value = "";
  }
  applyLabels();
  showStep("form");
  show(steps.form, false);
  setText(status, "Loading...", "detail loading");
  deskCache = new Map();
  try {
    org = await loadOrganisation(api);
  } catch (error) {
    console.error(error);
    setText(status, `Couldn't load your organisation. ${describeError(error)}`, "detail error");
    return;
  }
  if (!org.buildings.length) {
    setText(status, "No buildings are set up in PlaceOS for you to book in.", "detail error");
    return;
  }
  const buildings = select("deskBuilding");
  if (!buildings.options.length) {
    org.buildings.forEach((building) => option(buildings, building.id, zoneLabel(building)));
  }
  const blocked = await changeBuilding(
    buildingId || (await org.defaultBuilding())?.id || org.buildings[0].id
  );
  if (!request) {
    resetForm();
  }
  setText(status, "");
  if (!blocked) {
    showStep("form");
  }
}

/**
 * Switches building. If the user has an assigned desk or space there that stops them booking another, or
 * may not book parking there, shows that instead of the form and returns true.
 */
async function changeBuilding(id: string): Promise<boolean> {
  if (!org) {
    return false;
  }
  buildingId = id;
  select("deskBuilding").value = id;
  org.rememberBuilding(id);
  settings = await org.settings(id);
  const building = org.buildings.find((b) => b.id === id);
  setText(el("deskSubtitle"), building ? zoneLabel(building) : "");
  show(el("deskAllDayField"), deskAllDayAllowed(settings, kind));
  if (!deskAllDayAllowed(settings, kind)) {
    input("deskAllDay").checked = false;
  }
  const date = input("deskDate");
  date.min = dateValue(Date.now());
  date.max = dateValue(Date.now() + maxFutureDays() * 24 * 60 * minute);
  updateAllDay();
  // Duration limits are per building.
  if (select("deskStart").value) {
    fillEndTimes();
  }
  const parkingBlock = parking() ? await parkingRestriction(id) : "";
  if (buildingId !== id) {
    return false;
  }
  if (parkingBlock) {
    showBlocked(parkingBlock);
    return true;
  }
  let assigned: Desk | undefined;
  try {
    assigned = blockingAssignedDesk(
      settings,
      await desksIn(searchScope().buildingIds),
      org.user.email,
      kind
    );
  } catch (error) {
    // Finding desks reports it.
    console.error(error);
  }
  if (assigned) {
    showAssigned(assigned);
    return true;
  }
  if (steps.assigned.style.display !== "none") {
    showStep("form");
  }
  return false;
}

/** The zone to search (the building, or its region with app.use_region) and its buildings. */
function searchScope(): { zoneId: string; buildingIds: string[] } {
  const regionId = settings.get<boolean>("use_region")
    ? org?.buildings.find((b) => b.id === buildingId)?.parent_id
    : undefined;
  return {
    zoneId: regionId || buildingId,
    buildingIds: regionId
      ? (org?.buildings || []).filter((b) => b.parent_id === regionId).map((b) => b.id)
      : [buildingId],
  };
}

/**
 * Parking only: why the user may not book parking in this building, or "". Also loads their parking user
 * entry and fills in their plate number. As Workplace's parking flow: a parking user with deny set may not
 * book, and with app.parking.restrict_home_location nor may someone with an assigned desk in the building.
 */
async function parkingRestriction(id: string): Promise<string> {
  if (!org) {
    return "";
  }
  parkingUser = await getParkingUser(api, id, org.user.email).catch((error: unknown) => {
    console.error(error);
    return null;
  });
  // As Workplace's parking form: only a parking user gets the plate filled in, from their plate_number user
  // setting, else their entry.
  const plate = input("deskPlate");
  if (!plate.value && parkingUser) {
    plate.value = `${org.userSettings.plate_number || ""}` || parkingUser.plate_number;
  }
  if (parkingUser?.deny) {
    return "Your user account is not allowed to book parking in this building.";
  }
  if (settings.get<boolean>("parking.restrict_home_location")) {
    const email = org.user.email.toLowerCase();
    const desks = await getDesks(api, org.levelsIn([id]), "desk").catch(() => []);
    if (desks.some((desk) => desk.assigned_to.toLowerCase() === email)) {
      return "You have an assigned desk in this building, so you can't book parking here.";
    }
  }
  return "";
}

/** Levels to search: for parking only the levels tagged "parking". */
function searchLevels(buildingIds: string[]): Zone[] {
  const levels = org?.levelsIn(buildingIds) || [];
  return parking() ? levels.filter(isParkingLevel) : levels;
}

function desksIn(buildingIds: string[]): Promise<Desk[]> {
  const key = `${kind}:${buildingIds.join(",")}`;
  let desks = deskCache.get(key);
  if (!desks) {
    desks = getDesks(api, searchLevels(buildingIds), kind);
    deskCache.set(key, desks);
    desks.catch(() => deskCache.delete(key));
  }
  return desks;
}

function resetForm() {
  request = null;
  chosen = null;
  input("deskDate").value = dateValue(Date.now());
  input("deskAllDay").checked =
    deskAllDayAllowed(settings, kind) &&
    kindSetting<boolean>(settings, kind, "all_day_default", false) === true;
  select("deskStart").textContent = "";
  fillTimes();
  setText(el("deskFormError"), "");
}

/** A setting of this kind's form: app.desks.<key> or app.parking.<key>. */
const ownSetting = <T>(key: string) => settings?.get<T>(`${kinds[kind].settings}.${key}`);

// Workplace's booking window (app.desks.available_period, 90 days; app.parking.available_period, 7).
const maxFutureDays = () => ownSetting<number>("available_period") || kinds[kind].availablePeriod;

// ---------------------------------------------------------------------------------------------------------
// Date and time
// ---------------------------------------------------------------------------------------------------------

/** Start times for the chosen day; today starts from the next slot, other days at 9:00. */
function fillTimes() {
  const day = parseDate(input("deskDate").value);
  const start = select("deskStart");
  const previous = start.value ? new Date(+start.value) : null;
  start.textContent = "";
  if (isNaN(day)) {
    fillEndTimes();
    updateAllDay();
    return;
  }
  const now = Date.now();
  for (let minutes = 0; minutes < 24 * 60; minutes += startStep) {
    const time = day + minutes * minute;
    if (time + startStep * minute > now) {
      option(start, `${time}`, formatTime(time));
    }
  }
  const match = previous
    ? day + (previous.getHours() * 60 + previous.getMinutes()) * minute
    : dateValue(day) === dateValue(now)
      ? Math.ceil(now / (startStep * minute)) * startStep * minute
      : day + 9 * 60 * minute;
  start.value = `${match}`;
  if (!start.value && start.options.length) {
    start.selectedIndex = 0;
  }
  fillEndTimes(
    previous ? undefined : kindSetting<number>(settings, kind, "default_duration") || 60
  );
  updateAllDay();
}

const allDayChecked = () => deskAllDayAllowed(settings, kind) && input("deskAllDay").checked;

/** The all-day period: app.desk(s)/parking.all_day_period, else app.bookings.all_day_period. */
const deskAllDayPeriod = (day: number) =>
  allDayPeriod(settings, day, kindSetting(settings, kind, "all_day_period", null));

/** The chosen start (ms) and duration (minutes): the all-day period, or the start and end pickers. */
function formTimes(): { start: number; duration: number } {
  if (allDayChecked()) {
    const day = parseDate(input("deskDate").value);
    return isNaN(day) ? { start: 0, duration: 0 } : deskAllDayPeriod(day);
  }
  return { start: +select("deskStart").value, duration: +select("deskEnd").value };
}

function updateAllDay() {
  const allDay = allDayChecked();
  show(el("deskTimeFields"), !allDay);
  const { start, duration } = formTimes();
  setText(
    el("deskAllDayTimes"),
    !allDay || !start
      ? ""
      : start + duration * minute <= Date.now()
        ? "The all-day period today has ended."
        : isFullDay(start, duration)
          ? ""
          : `${formatTime(start)} – ${formatTime(start + duration * minute)}`
  );
}

/** End times from the shortest to the longest booking, in deskDurationStep steps. */
function fillEndTimes(keepDuration?: number) {
  const start = +select("deskStart").value;
  const end = select("deskEnd");
  const previous = keepDuration ?? (end.value ? +end.value : 60);
  end.textContent = "";
  if (!start) {
    return;
  }
  const { min, max } = deskDurationLimits(settings, kind);
  const step = deskDurationStep(settings, kind);
  for (let minutes = min; minutes <= max; minutes += step) {
    option(
      end,
      `${minutes}`,
      `${formatTime(start + minutes * minute)} (${formatDuration(minutes)})`
    );
  }
  const options = Array.from(end.options);
  const closest = options.find((o) => +o.value >= previous) || options[options.length - 1];
  if (closest) {
    end.value = closest.value;
  }
}

function readForm(): DeskRequest | string {
  const day = parseDate(input("deskDate").value);
  if (isNaN(day)) {
    return "Choose a date.";
  }
  if (day > Date.now() + maxFutureDays() * 24 * 60 * minute) {
    return `${parking() ? "Parking" : "Desks"} can only be booked up to ${maxFutureDays()} days ahead.`;
  }
  const allDay = allDayChecked();
  const { start, duration } = formTimes();
  if (!start || duration <= 0) {
    return "Choose a time.";
  }
  if (allDay && start + duration * minute <= Date.now()) {
    return "The all-day period today has ended. Choose another day.";
  }
  if (!allDay) {
    if (start + duration * minute <= Date.now()) {
      return "Choose a time that hasn't passed.";
    }
    const hours = checkBookableHours(
      settings,
      start,
      duration,
      kindSetting(settings, kind, "bookable_hours", null),
      parking() ? "Parking" : "Desks"
    );
    if (hours) {
      return hours;
    }
  }
  if (!parking()) {
    return { start, duration, allDay, buildingId };
  }
  const plateNumber = input("deskPlate").value.trim();
  if (!plateNumber && settings.get<boolean>("parking.require_plate_number")) {
    return "Enter your vehicle's plate number.";
  }
  return { start, duration, allDay, buildingId, plateNumber };
}

// ---------------------------------------------------------------------------------------------------------
// Free desks
// ---------------------------------------------------------------------------------------------------------

async function findDesks() {
  const form = readForm();
  if (typeof form === "string") {
    setText(el("deskFormError"), form, "error");
    return;
  }
  if (!org) {
    return;
  }
  request = form;
  const button = el<HTMLButtonElement>("deskFindButton");
  button.disabled = true;
  setText(el("deskFormError"), `Finding free ${nouns()}...`, "detail loading");
  try {
    const { zoneId, buildingIds } = searchScope();
    const [desks, rules] = await Promise.all([
      desksIn(buildingIds),
      getDeskRules(api, buildingIds, kind),
    ]);
    if (!desks.length) {
      setText(
        el("deskFormError"),
        `No ${nouns()} are set up in PlaceOS for this building.`,
        "error"
      );
      return;
    }
    const assigned = blockingAssignedDesk(settings, desks, org.user.email, kind);
    if (assigned) {
      setText(el("deskFormError"), "");
      showAssigned(assigned);
      return;
    }
    const booked = await getBookedDeskIds(api, zoneId, form, desks.length, kind);
    const allowed = bookableDesks(desks, rules, form, org.user.groups || []);
    searchedDesks = desks;
    bookedIds = booked;
    candidates = allowed.filter((desk) => !booked.has(desk.id)).sort(byLevelAndName);
    mapSelection = null;
    setText(el("deskFormError"), "");
    prepareFilters(buildingIds);
    renderDesks();
    showStep("desks");
  } catch (error) {
    console.error(error);
    setText(el("deskFormError"), describeError(error), "error");
  } finally {
    button.disabled = false;
  }
}

const byLevelAndName = (a: Desk, b: Desk) =>
  zoneLabel(a.level).localeCompare(zoneLabel(b.level), undefined, { numeric: true }) ||
  a.name.localeCompare(b.name, undefined, { numeric: true });

function prepareFilters(buildingIds: string[]) {
  if (!org || !request) {
    return;
  }
  const levels = select("deskLevel");
  const previousLevel = levels.value;
  levels.textContent = "";
  option(levels, "", "All levels");
  const seen = new Map<string, Zone>();
  candidates.forEach((desk) => seen.set(desk.level.id, desk.level));
  Array.from(seen.values())
    .sort((a, b) => zoneLabel(a).localeCompare(zoneLabel(b), undefined, { numeric: true }))
    .forEach((level) => option(levels, level.id, levelLabel(level, buildingIds)));
  levels.value = seen.has(previousLevel) ? previousLevel : "";

  const features = el("deskFeatures");
  const checked = new Set(
    Array.from(features.querySelectorAll("input:checked")).map(
      (box) => (box as HTMLInputElement).value
    )
  );
  features.textContent = "";
  const all = Array.from(new Set(candidates.flatMap((desk) => desk.features))).sort();
  for (const feature of all) {
    const label = document.createElement("label");
    label.className = "checkbox";
    const box = document.createElement("input");
    box.type = "checkbox";
    box.value = feature;
    box.checked = checked.has(feature);
    box.addEventListener("change", renderDesks);
    label.appendChild(box);
    label.appendChild(document.createTextNode(` ${feature}`));
    features.appendChild(label);
  }
  setText(el("deskDesksSummary"), whenText(request));
  prepareMapLevels(buildingIds);
}

/** "Level 3", with the building when desks or spaces come from a region. */
function levelLabel(level: Zone, buildingIds: string[]): string {
  const building =
    buildingIds.length > 1 ? org?.buildings.find((b) => b.id === level.parent_id) : undefined;
  return `${building ? `${zoneLabel(building)}, ` : ""}${zoneLabel(level)}`;
}

function deskLocation(desk: Desk): string {
  const building = org?.buildings.find((b) => b.id === desk.level.parent_id);
  return [zoneLabel(desk.level), building ? zoneLabel(building) : ""].filter(Boolean).join(", ");
}

/**
 * Free desks or spaces that match the filters. The map shows one level itself, so it ignores the Level
 * filter.
 */
function filteredDesks(ignoreLevel = deskView === "map"): Desk[] {
  const level = ignoreLevel ? "" : select("deskLevel").value;
  const features = Array.from(el("deskFeatures").querySelectorAll("input:checked")).map(
    (box) => (box as HTMLInputElement).value
  );
  return candidates.filter(
    (desk) =>
      (!level || desk.level.id === level) &&
      features.every((feature) => desk.features.includes(feature))
  );
}

function renderDesks() {
  const list = el("deskList");
  list.textContent = "";
  const desks = filteredDesks();
  setText(
    el("deskFiltersToggle"),
    desks.length !== candidates.length
      ? `Filters (${candidates.length - desks.length} hidden)`
      : "Filters"
  );
  if (deskView === "map") {
    renderMap();
    return;
  }
  if (!desks.length) {
    const item = document.createElement("li");
    item.className = "empty";
    item.textContent = candidates.length
      ? `No ${nouns()} match these filters.`
      : `No ${nouns()} are free at this time. Try another time or building.`;
    list.appendChild(item);
    return;
  }
  for (const desk of desks) {
    const item = document.createElement("li");
    const button = resultCard({
      name: desk.name,
      details: [deskLocation(desk), desk.features.slice(0, 4).join(", ")],
      action: `Choose ${noun()} →`,
      onChoose: () => chooseDesk(desk),
    });
    item.appendChild(button);
    list.appendChild(item);
  }
}

// ---------------------------------------------------------------------------------------------------------
// Map: one level's floor plan, free desks or spaces coloured and tappable
// ---------------------------------------------------------------------------------------------------------

function setDeskView(view: "list" | "map", render = true) {
  deskView = view;
  const list = view === "list";
  for (const [id, pressed] of [
    ["deskListViewButton", list],
    ["deskMapViewButton", !list],
  ] as [string, boolean][]) {
    el(id).setAttribute("aria-pressed", `${pressed}`);
    el(id).className = pressed ? "" : "secondary";
  }
  show(el("deskList"), list);
  show(el("deskMapView"), !list);
  // The map has its own level picker.
  show(el("deskLevelField"), list);
  if (!list) {
    chooseDefaultMapLevel();
  }
  if (render) {
    renderDesks();
  }
}

/**
 * Fills the map's level picker with levels that have a floor plan and desks or spaces.
 * Opens the map view when there is one, unless the user picked List.
 */
function prepareMapLevels(buildingIds: string[]) {
  if (!org) {
    return;
  }
  const withDesks = new Set(searchedDesks.map((desk) => desk.level.id));
  mapLevels = org.mapLevels(buildingIds, parking()).filter((level) => withDesks.has(level.id));
  const picker = select("deskMapLevel");
  const previous = picker.value;
  picker.textContent = "";
  mapLevels.forEach((level) => option(picker, level.id, levelLabel(level, buildingIds)));
  picker.value = mapLevels.some((level) => level.id === previous) ? previous : "";
  updateMapLevelLabels(buildingIds);
  show(el("deskViewToggle"), mapLevels.length > 0);
  if (!mapLevels.length && deskView === "map") {
    setDeskView("list", false);
  } else if (mapLevels.length && deskView === "list" && !listChosen) {
    setDeskView("map", false);
  } else if (deskView === "map" && !picker.value) {
    chooseDefaultMapLevel();
  }
}

/** "Level 3 (12 free)". */
function updateMapLevelLabels(buildingIds?: string[]) {
  const free = filteredDesks(true);
  const options = select("deskMapLevel").options;
  const ids = buildingIds ?? Array.from(new Set(mapLevels.map((level) => level.parent_id || "")));
  mapLevels.forEach((level, index) => {
    const count = free.filter((desk) => desk.level.id === level.id).length;
    options[index].textContent = `${levelLabel(level, ids)} (${count} free)`;
  });
}

/** The Level filter's level if it has a map, else the current pick, else the level with most free. */
function chooseDefaultMapLevel() {
  const picker = select("deskMapLevel");
  const filterLevel = select("deskLevel").value;
  if (filterLevel && mapLevels.some((level) => level.id === filterLevel)) {
    picker.value = filterLevel;
    return;
  }
  if (picker.value) {
    return;
  }
  const free = filteredDesks(true);
  let best: Zone | undefined;
  let bestCount = -1;
  for (const level of mapLevels) {
    const count = free.filter((desk) => desk.level.id === level.id).length;
    if (count > bestCount) {
      best = level;
      bestCount = count;
    }
  }
  picker.value = best?.id || "";
}

/**
 * Colours from app.explore.colors, as Workplace's desk and parking maps (desk-free, parking-free etc.), else
 * its defaults.
 */
function mapColours(): Record<MapRoomStatus, string> {
  const colours = settings.get<Record<string, string>>("explore.colors", {}) || {};
  const pick = (status: MapRoomStatus, key: string) =>
    colours[`${kind}-${key}`] || colours[key] || defaultMapColours[status];
  return {
    free: pick("free", "free"),
    busy: pick("busy", "busy"),
    filtered: pick("filtered", "not-bookable"),
  };
}

async function renderMap() {
  if (!floorMap) {
    return;
  }
  updateMapLevelLabels();
  const mapStatus = el("deskMapStatus");
  const level = mapLevels.find((l) => l.id === select("deskMapLevel").value);
  if (!level?.map_id) {
    floorMap.clear();
    shownMapUrl = "";
    setText(mapStatus, "No floor plans are set up for this building.", "detail");
    renderMapSelection();
    return;
  }
  const seq = ++mapLoadSeq;
  if (shownMapUrl !== level.map_id) {
    setText(mapStatus, "Loading the floor plan...", "detail loading");
    try {
      const svg = await loadMapSvg(level.map_id, api);
      if (seq !== mapLoadSeq) {
        return;
      }
      floorMap.show(svg);
      shownMapUrl = level.map_id;
    } catch (error) {
      if (seq !== mapLoadSeq) {
        return;
      }
      console.error(error);
      floorMap.clear();
      shownMapUrl = "";
      setText(mapStatus, `Couldn't load the floor plan. ${describeError(error)}`, "detail error");
      renderMapSelection();
      return;
    }
  }

  const free = new Set(filteredDesks(true).map((desk) => desk.id));
  const freeAnyFilter = new Set(candidates.map((desk) => desk.id));
  const onLevel = searchedDesks.filter((desk) => desk.level.id === level.id);
  // As Workplace's desk and parking maps: free desks that match the filters are green, booked desks red
  // (whoever may book them), and the rest (filtered out, or not bookable by this user) grey.
  const desks: MapRoom[] = onLevel.map((desk) => {
    const status: MapRoomStatus = free.has(desk.id)
      ? "free"
      : bookedIds.has(desk.id)
        ? "busy"
        : "filtered";
    const label = {
      free: "free",
      busy: "booked",
      filtered: freeAnyFilter.has(desk.id) ? "doesn't match the filters" : "not available to you",
    }[status];
    return {
      elementId: desk.map_id,
      status,
      title: `${desk.name} (${label})`,
      selectable: status === "free",
    };
  });
  const colours = mapColours();
  const accent =
    window.getComputedStyle(document.documentElement).getPropertyValue("--accent").trim() ||
    "#0078d4";
  const missing = new Set(floorMap.setRooms(desks, colours, accent));
  if (mapSelection && !(free.has(mapSelection.id) && mapSelection.level.id === level.id)) {
    mapSelection = null;
  }
  floorMap.select(mapSelection ? mapSelection.map_id : null);

  const freeHere = onLevel.filter((desk) => free.has(desk.id));
  const unplaced = freeHere.filter((desk) => missing.has(desk.map_id)).length;
  setText(
    mapStatus,
    [
      freeHere.length
        ? `Tap a green ${noun()} to choose it.`
        : candidates.length
          ? `No free ${nouns()} on this level match your filters. Try another level.`
          : `No ${nouns()} are free at this time. Try another time or building.`,
      unplaced
        ? `${unplaced} free ${noun()}${unplaced === 1 ? " isn't" : "s aren't"} on this floor plan; use the list to see ${unplaced === 1 ? "it" : "them"}.`
        : "",
    ]
      .filter(Boolean)
      .join(" "),
    "detail"
  );
  renderMapLegend(colours);
  renderMapSelection();
}

function renderMapLegend(colours: Record<MapRoomStatus, string>) {
  const legend = el("deskMapLegend");
  legend.textContent = "";
  const entries: [MapRoomStatus, string][] = [
    ["free", "Free"],
    ["busy", "Booked"],
    ["filtered", "Not available"],
  ];
  for (const [state, text] of entries) {
    const item = document.createElement("li");
    const swatch = document.createElement("span");
    swatch.className = "swatch";
    swatch.style.backgroundColor = colours[state];
    item.appendChild(swatch);
    item.appendChild(document.createTextNode(text));
    legend.appendChild(item);
  }
}

function selectOnMap(elementId: string) {
  const level = select("deskMapLevel").value;
  mapSelection =
    filteredDesks(true).find((desk) => desk.map_id === elementId && desk.level.id === level) ||
    null;
  floorMap?.select(mapSelection ? elementId : null);
  renderMapSelection();
  if (mapSelection) {
    el("deskMapSelection").scrollIntoView({ block: "nearest" });
  }
}

function renderMapSelection() {
  const card = el("deskMapSelection");
  card.textContent = "";
  const desk = mapSelection;
  show(card, !!desk);
  if (!desk) {
    return;
  }
  selectedResult(card, `Selected ${noun()}`);
  line(card, desk.name, "item-title");
  line(card, deskLocation(desk));
  line(card, desk.features.slice(0, 6).join(", "));
  const button = document.createElement("button");
  button.type = "button";
  button.className = "card-action";
  button.textContent = `Choose this ${noun()}`;
  button.addEventListener("click", () => chooseDesk(desk));
  card.appendChild(button);
}

// ---------------------------------------------------------------------------------------------------------
// Assigned desk or space: shown instead of booking when app.bookings.assigned_resource_booking doesn't allow
// someone with one assigned to book another. The same section explains other reasons the user can't book.
// ---------------------------------------------------------------------------------------------------------

function showAssigned(desk: Desk) {
  assignedDesk = desk;
  const details = el("deskAssignedDetails");
  details.textContent = "";
  line(details, `Your assigned ${noun()}`, "item-detail");
  line(details, desk.name, "item-title");
  line(details, deskLocation(desk));
  line(details, desk.features.join(", "));
  line(details, `You can use it any time, so you can't book another ${noun()}.`);
  show(el("deskAssignedMapButton"), !!levelMapUrl(desk));
  show(el("deskAssignedMap"), false);
  assignedMap?.clear();
  showStep("assigned");
}

/** Shows why the user can't book in this building instead of the form. */
function showBlocked(message: string) {
  assignedDesk = null;
  const details = el("deskAssignedDetails");
  details.textContent = "";
  line(details, message, "item-title");
  show(el("deskAssignedMapButton"), false);
  show(el("deskAssignedMap"), false);
  assignedMap?.clear();
  showStep("assigned");
}

/** The floor plan of a desk's level, if it has one. */
const levelMapUrl = (desk: Desk) => desk.level.map_id || "";

/** Shows the assigned desk's level with the desk highlighted and zoomed in on. */
async function showAssignedMap() {
  const desk = assignedDesk;
  const url = desk ? levelMapUrl(desk) : "";
  if (!desk || !url || !assignedMap) {
    return;
  }
  const mapStatus = el("deskAssignedMapStatus");
  show(el("deskAssignedMap"), true);
  show(el("deskAssignedMapButton"), false);
  setText(mapStatus, "Loading the floor plan...", "detail loading");
  try {
    assignedMap.show(await loadMapSvg(url, api));
  } catch (error) {
    console.error(error);
    assignedMap.clear();
    setText(mapStatus, `Couldn't load the floor plan. ${describeError(error)}`, "detail error");
    return;
  }
  if (assignedDesk !== desk) {
    return;
  }
  const accent =
    window.getComputedStyle(document.documentElement).getPropertyValue("--accent").trim() ||
    "#0078d4";
  const missing = assignedMap.setRooms(
    [
      {
        elementId: desk.map_id,
        status: "free",
        title: `${desk.name} (your ${noun()})`,
        selectable: false,
      },
    ],
    { ...mapColours(), free: accent },
    accent
  );
  if (missing.length) {
    setText(mapStatus, `${desk.name} isn't on the ${zoneLabel(desk.level)} floor plan.`, "detail");
    return;
  }
  assignedMap.select(desk.map_id);
  assignedMap.focus(desk.map_id);
  setText(mapStatus, `${desk.name}, ${zoneLabel(desk.level)}`, "detail");
  el("deskAssignedMap").scrollIntoView({ block: "nearest" });
}

// ---------------------------------------------------------------------------------------------------------
// Confirm and book
// ---------------------------------------------------------------------------------------------------------

function chooseDesk(desk: Desk) {
  if (!request) {
    return;
  }
  chosen = desk;
  const details = el("deskDetails");
  details.textContent = "";
  selectedResult(details, `Selected ${noun()}`);
  line(details, desk.name, "item-title");
  line(details, deskLocation(desk));
  line(details, desk.features.join(", "));
  line(details, whenText(request));
  line(details, request.plateNumber ? `Plate ${request.plateNumber}` : "");
  setText(el("deskConfirmError"), "");
  showStep("confirm");
}

async function submit() {
  if (!request || !chosen || !org || submitting) {
    return;
  }
  const errorText = el("deskConfirmError");
  submitting = true;
  const buttons = ["deskSubmitButton", "deskOtherButton"].map((id) => el<HTMLButtonElement>(id));
  buttons.forEach((button) => (button.disabled = true));
  setText(errorText, "Booking...", "detail");
  try {
    const problem = await checkBeforeBooking(api, settings, org, chosen, request);
    if (problem) {
      throw new Error(problem);
    }
    await createDeskBooking(api, buildDeskBooking(request, chosen, org, settings));
    showDone();
  } catch (error) {
    console.error(error);
    setText(errorText, describeError(error), "error");
  } finally {
    submitting = false;
    buttons.forEach((button) => (button.disabled = false));
  }
}

function showDone() {
  if (!request || !chosen) {
    return;
  }
  const summary = el("deskDoneSummary");
  summary.textContent = "";
  line(summary, parking() ? "✓ Parking booked" : "✓ Desk booked", "item-title ok");
  line(summary, `${chosen.name}, ${deskLocation(chosen)}`);
  line(summary, whenText(request));
  line(summary, request.plateNumber ? `Plate ${request.plateNumber}` : "");
  request = null;
  chosen = null;
  showStep("done");
}
