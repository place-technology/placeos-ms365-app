/*
 * The Book a room view: details form → available rooms (list or floor plan map, with filters) → confirm with
 * catering and assets → done. Mirrors the legacy Outlook add-in's room flow and meeting form.
 */

/* global document, console, window, HTMLElement, HTMLInputElement, HTMLSelectElement, HTMLTextAreaElement, HTMLButtonElement */

import {
  loadOrganisation,
  zoneLabel,
  type AppSettings,
  type Organisation,
  type Zone,
} from "./booking-settings";
import {
  assetBooking,
  assetsDisabledForRoom,
  assetsEnabled,
  buildAssetRequest,
  filterAssetTypes,
  getAvailableAssetTypes,
  type AssetSelection,
} from "./asset-data";
import {
  buildCateringOrders,
  cateringBookings,
  cateringEnabled,
  cateringProvider,
  deliveryOffsets,
  filterCateringMenu,
  formatPrice,
  getCateringMenu,
  getCateringSettings,
  optionGroups,
  unitPriceWithOptions,
  type CateringSelection,
  type CateringSettings,
} from "./catering-data";
import {
  defaultMapColours,
  FloorMap,
  loadMapSvg,
  type MapRoom,
  type MapRoomStatus,
} from "./floor-map";
import type { PlaceosApi } from "./placeos-data";
import { PlaceosRequestError } from "./placeos-helper";
import {
  allDayAllowed,
  allDayPeriod,
  bookingZones,
  buildEventBody,
  checkBookableHours,
  createEvent,
  createLinkedBookings,
  dateValue,
  findClashes,
  getAvailableSpaceIds,
  getBookingRules,
  getSpaces,
  hiddenByRules,
  isEmail,
  isFullDay,
  removeEvent,
  searchPeople,
  spaceFeatures,
  spaceName,
  visitorBookings,
  minutesOfDay,
  monthlyStartUnsupported,
  monthlyStartUnsupportedMessage,
  weekOfMonth,
  type BookingRequest,
  type DraftDetails,
  type MeetingDraft,
  type Person,
  type RecurrencePattern,
  type Space,
} from "./room-booking-data";

const el = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
const input = (id: string) => el<HTMLInputElement>(id);
const select = (id: string) => el<HTMLSelectElement>(id);
const show = (element: HTMLElement, visible: boolean) =>
  (element.style.display = visible ? "" : "none");

const status = el("bookStatus");
const subtitle = el("bookSubtitle");
const steps = {
  form: el("bookForm"),
  rooms: el("bookRooms"),
  confirm: el("bookConfirm"),
  done: el("bookDone"),
};
type Step = keyof typeof steps;

const weekdays = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const minute = 60000;
// User settings lists (Workplace's keys).
const favouriteMenuItems = "favourite_menu_items";
const favouriteAssets = "favourite_assets";

let api: PlaceosApi;
let onClose: (booked: boolean) => void = () => undefined;
let org: Organisation | null = null;
let settings: AppSettings;
let buildingId = "";
// The Outlook meeting being composed, if any: the room is added to it instead of PlaceOS booking it.
let getDraft: () => MeetingDraft | null = () => null;
let draft: MeetingDraft | null = null;

// Form state.
let attendees: Person[] = [];
let request: BookingRequest | null = null;
// Rooms step. candidates are the free rooms; searchedRooms also has the busy ones (for the map).
let candidates: Space[] = [];
let searchedRooms: Space[] = [];
let searchedBuildingIds: string[] = [];
let roomView: "list" | "map" = "list";
let floorMap: FloorMap | null = null;
let mapLevels: Zone[] = [];
// map_id of the floor plan on screen, and the request for the one loading.
let shownMapUrl = "";
let mapLoadSeq = 0;
let mapSelection: Space | null = null;
// Confirm step.
let cateringSelections: CateringSelection[] = [];
let cateringSettings: CateringSettings | null = null;
let assetSelections: AssetSelection[] = [];
let assetsAllowed = false;
let acceptClashes = false;
let submitting = false;

function showStep(step: Step) {
  (Object.keys(steps) as Step[]).forEach((key) => show(steps[key], key === step));
  window.scrollTo(0, 0);
}

function setText(element: HTMLElement, text: string, className?: string) {
  element.textContent = text;
  if (className !== undefined) {
    element.className = className;
  }
}

function describeError(error: unknown): string {
  if (error instanceof PlaceosRequestError) {
    if (error.status === 511) {
      return "PlaceOS can't access your calendar yet. Sign in to PlaceOS on the web once, then try again.";
    }
    if (error.status === 409) {
      return "That room or one of the extras was booked by someone else in the meantime. Choose again.";
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

/** "Thursday 1 October, 9:00 am – 10:00 am", or "Thursday 1 October, all day". */
function whenText(booking: { start: number; duration: number; allDay: boolean }): string {
  const { start, duration, allDay } = booking;
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

// ---------------------------------------------------------------------------------------------------------
// Setup
// ---------------------------------------------------------------------------------------------------------

/**
 * Wires up the view once. onDone is called when the user leaves it (booked = a booking was made).
 * meetingDraft returns the Outlook meeting being composed, if any, each time the view opens.
 */
export function initRoomBooking(
  placeosApi: PlaceosApi,
  onDone: (booked: boolean) => void,
  meetingDraft?: () => MeetingDraft | null
) {
  api = placeosApi;
  onClose = onDone;
  getDraft = meetingDraft ?? (() => null);
  el("bookForm").addEventListener("submit", (event) => {
    event.preventDefault();
    findRooms();
  });
  el("bookClearButton").addEventListener("click", resetForm);
  el("bookCancelButton").addEventListener("click", () => onClose(false));
  select("bookBuilding").addEventListener("change", () =>
    changeBuilding(select("bookBuilding").value)
  );
  input("bookDate").addEventListener("change", fillTimes);
  input("bookAllDay").addEventListener("change", updateAllDay);
  select("bookStart").addEventListener("change", () => fillEndTimes());
  select("bookRepeat").addEventListener("change", updateRepeatFields);
  setupAttendeeSearch();

  el("bookFiltersToggle").addEventListener("click", () =>
    show(el("bookFilters"), el("bookFilters").style.display === "none")
  );
  el("bookBackToForm").addEventListener("click", () => showStep("form"));
  for (const id of ["bookLevel", "bookCapacity", "bookFavourites"]) {
    el(id).addEventListener("change", renderRooms);
  }
  el("bookListViewButton").addEventListener("click", () => setRoomView("list"));
  el("bookMapViewButton").addEventListener("click", () => setRoomView("map"));
  select("bookMapLevel").addEventListener("change", () => {
    mapSelection = null;
    renderMap();
  });
  floorMap = new FloorMap(el("bookMapCanvas"));
  floorMap.onSelect = selectOnMap;
  el("bookMapZoomIn").addEventListener("click", () => floorMap?.zoomBy(1.5));
  el("bookMapZoomOut").addEventListener("click", () => floorMap?.zoomBy(1 / 1.5));
  el("bookMapReset").addEventListener("click", () => floorMap?.reset());

  el("bookOtherRoom").addEventListener("click", () => showStep("rooms"));
  el("bookSubmitButton").addEventListener("click", () => submit());
  el("bookAnywayButton").addEventListener("click", () => {
    acceptClashes = true;
    submit();
  });
  select("bookCateringDeliver").addEventListener("change", renderCateringTotal);
  setupMenuFilters("bookCatering", renderCateringMenu);
  select("bookCaterer").addEventListener("change", renderCateringMenu);
  setupMenuFilters("bookAssets", renderAssetList);
  el("bookDoneButton").addEventListener("click", () => onClose(true));
  el("bookAnotherButton").addEventListener("click", () => {
    resetForm();
    showStep("form");
  });
}

/** Opens the view at the details form. Loads the org and settings the first time. */
export async function openRoomBooking() {
  showStep("form");
  show(steps.form, false);
  setText(status, "Loading...", "detail");
  draft = getDraft();
  showDraftMode();
  let details: DraftDetails | null = null;
  try {
    details = draft ? await draft.read() : null;
  } catch (error) {
    console.error(error);
    setText(
      status,
      `Couldn't read the meeting from Outlook. ${describeError(error)}`,
      "detail error"
    );
    return;
  }
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
  const buildings = select("bookBuilding");
  if (!buildings.options.length) {
    org.buildings.forEach((building) => option(buildings, building.id, zoneLabel(building)));
  }
  show(el("bookBuildingField"), org.buildings.length > 1);
  const building = buildingId || (await org.defaultBuilding())?.id || org.buildings[0].id;
  await changeBuilding(building);
  if (details) {
    // Start from the draft each time, as it may have changed in Outlook.
    prefillFromDraft(details);
  } else if (!request) {
    resetForm();
  }
  setText(status, "");
  show(steps.form, true);
  input("bookTitle").focus();
}

async function changeBuilding(id: string) {
  if (!org) {
    return;
  }
  buildingId = id;
  select("bookBuilding").value = id;
  org.rememberBuilding(id);
  settings = await org.settings(id);
  const building = org.buildings.find((b) => b.id === id);
  setText(subtitle, building ? zoneLabel(building) : "");
  // A draft keeps its own body and repeat; all day isn't set on it, so only times are offered.
  show(el("bookNotesField"), !draft && !settings.get<boolean>("events.hide_notes"));
  // Workplace's meeting form only offers repeats with app.events.allow_recurrence.
  show(el("bookRepeatFields"), !draft && settings.get<boolean>("events.allow_recurrence") === true);
  show(el("bookAllDayField"), !draft && allDayAllowed(settings));
  if (draft || !allDayAllowed(settings)) {
    input("bookAllDay").checked = false;
  }
  updateAllDay();
  // Duration limits are per building.
  if (select("bookStart").value) {
    fillEndTimes();
  }
}

function resetForm() {
  request = null;
  attendees = [];
  input("bookTitle").value = "";
  el<HTMLTextAreaElement>("bookNotes").value = "";
  input("bookDate").value = dateValue(Date.now());
  input("bookDate").min = dateValue(Date.now());
  input("bookDate").max = dateValue(Date.now() + maxFutureDays() * 24 * 60 * minute);
  input("bookAllDay").checked = false;
  select("bookRepeat").value = "none";
  input("bookRepeatInterval").value = "1";
  setText(el("bookFormError"), "");
  fillTimes();
  renderAttendees();
}

/** Labels for adding a room to the Outlook draft, or for booking it through PlaceOS. */
function showDraftMode() {
  setText(el("bookHeading"), draft ? "Add a room" : "Book a room");
  setText(el("bookSubmitButton"), draft ? "Add to meeting" : "Book room");
  show(el("bookAnotherButton"), !draft);
  show(el("bookDraftNote"), !!draft);
  setText(el("bookDraftNote"), "");
}

/** Fills the form from the Outlook draft: title, attendees, date and times. */
function prefillFromDraft(details: DraftDetails) {
  resetForm();
  input("bookTitle").value = details.title;
  const self = org?.user.email.toLowerCase();
  attendees = details.attendees
    .filter((person) => person.email.toLowerCase() !== self)
    .map((person) => ({ ...person, external: org?.isExternal(person.email, settings) }));
  renderAttendees();
  const duration = Math.round((details.end - details.start) / minute);
  if (details.start > Date.now() && duration > 0) {
    input("bookDate").value = dateValue(details.start);
    fillTimes();
    const start = select("bookStart");
    // Drafts can start off the 15-minute steps.
    ensureOption(start, details.start, formatTime(details.start));
    start.value = `${details.start}`;
    fillEndTimes(duration);
    const { min, max } = durationLimits();
    if (duration >= min && duration <= max) {
      const end = select("bookEnd");
      ensureOption(
        end,
        duration,
        `${formatTime(details.start + duration * minute)} (${formatDuration(duration)})`
      );
      end.value = `${duration}`;
    }
  }
  const note = details.recurring
    ? "This meeting repeats. Rooms are only checked for the first meeting, and the room may decline the series if it's busy for any of them."
    : "";
  setText(
    el("bookDraftNote"),
    `The room is added to the meeting you're writing. Send the invitation to book it. ${note}`.trim(),
    "detail"
  );
}

/** Adds an option for a numeric value if it's missing, keeping the list in order. */
function ensureOption(target: HTMLSelectElement, value: number, text: string) {
  const options = Array.from(target.options);
  if (options.some((item) => +item.value === value)) {
    return;
  }
  const item = document.createElement("option");
  item.value = `${value}`;
  item.textContent = text;
  target.insertBefore(item, options.find((o) => +o.value > value) || null);
}

const maxFutureDays = () => settings?.get<number>("events.allowed_future_days", 180) || 180;

// ---------------------------------------------------------------------------------------------------------
// Date and time
// ---------------------------------------------------------------------------------------------------------

const step = 15;

/** Start times for the chosen day in 15-minute steps; today starts from the next step. */
function fillTimes() {
  const day = parseDate(input("bookDate").value);
  const start = select("bookStart");
  const previous = start.value ? new Date(+start.value) : null;
  start.textContent = "";
  if (isNaN(day)) {
    fillEndTimes();
    return;
  }
  const now = Date.now();
  for (let minutes = 0; minutes < 24 * 60; minutes += step) {
    const time = day + minutes * minute;
    if (time + step * minute > now) {
      option(start, `${time}`, formatTime(time));
    }
  }
  // Keep the same time of day if possible, else the next slot (or 9:00 on another day).
  const match = previous
    ? day + (previous.getHours() * 60 + previous.getMinutes()) * minute
    : dateValue(day) === dateValue(now)
      ? Math.ceil(now / (step * minute)) * step * minute
      : day + 9 * 60 * minute;
  start.value = `${match}`;
  if (!start.value && start.options.length) {
    start.selectedIndex = 0;
  }
  fillEndTimes();
  updateAllDay();
}

const allDayChecked = () => allDayAllowed(settings) && input("bookAllDay").checked;

/** The chosen start (ms) and duration (minutes): the all-day period, or the start and end pickers. */
function formTimes(): { start: number; duration: number } {
  if (allDayChecked()) {
    const day = parseDate(input("bookDate").value);
    return isNaN(day) ? { start: 0, duration: 0 } : allDayPeriod(settings, day);
  }
  return { start: +select("bookStart").value, duration: +select("bookEnd").value };
}

/** Swaps the start and end pickers for the all-day period's times. */
function updateAllDay() {
  const allDay = allDayChecked();
  show(el("bookTimeFields"), !allDay);
  const { start, duration } = formTimes();
  setText(
    el("bookAllDayTimes"),
    !allDay || !start
      ? ""
      : start + duration * minute <= Date.now()
        ? "The all-day period today has ended."
        : isFullDay(start, duration)
          ? ""
          : `${formatTime(start)} – ${formatTime(start + duration * minute)}`
  );
  updateRepeatFields();
}

/** app.events.min_duration / max_duration, as Workplace's meeting form (30 and 480 minutes). */
function durationLimits() {
  const min = settings?.get<number>("events.min_duration", 30) || 30;
  const max = settings?.get<number>("events.max_duration", 480) || 480;
  return { min, max: Math.max(min, max) };
}

/** The building's timezone for bookable hours (as postForm), else local time. */
const buildingTimezone = () => org?.buildings.find((b) => b.id === buildingId)?.timezone || "";

/**
 * End times as Workplace's duration field: from the minimum in app.events.duration_step steps (default 30)
 * up to the maximum or the end of app.events.bookable_hours, plus app.events.custom_duration_options.
 * The default is app.events.default_duration (60).
 */
function fillEndTimes(keepDuration?: number) {
  const start = +select("bookStart").value;
  const end = select("bookEnd");
  const previous =
    keepDuration ??
    (end.value ? +end.value : settings?.get<number>("events.default_duration", 60) || 60);
  end.textContent = "";
  if (!start) {
    return;
  }
  const { min, max } = durationLimits();
  const durationStep = Math.max(1, settings?.get<number>("events.duration_step", 30) || 30);
  const hours = settings?.get<{ start?: number; end?: number } | null>(
    "events.bookable_hours",
    null
  );
  const latest =
    typeof hours?.end === "number"
      ? Math.max(0, hours.end * 60 - minutesOfDay(start, buildingTimezone()))
      : Number.POSITIVE_INFINITY;
  const custom = (settings?.get<number[]>("events.custom_duration_options", []) || [])
    .map((value) => Math.round(+value || 0))
    .filter((value) => value > 0 && value <= latest);
  const durations = new Set(custom);
  for (let minutes = min; minutes <= Math.min(max, latest); minutes += durationStep) {
    durations.add(minutes);
  }
  Array.from(durations)
    .sort((a, b) => a - b)
    .forEach((minutes) =>
      option(
        end,
        `${minutes}`,
        `${formatTime(start + minutes * minute)} (${formatDuration(minutes)})`
      )
    );
  end.value = `${previous}`;
  if (!end.value && end.options.length) {
    // The nearest choice that isn't shorter, else the longest.
    const values = Array.from(end.options).map((item) => +item.value);
    end.value = `${values.find((value) => value >= previous) ?? values[values.length - 1]}`;
  }
}

function updateRepeatFields() {
  const pattern = select("bookRepeat").value as RecurrencePattern;
  show(el("bookRepeatOptions"), pattern !== "none");
  const start = formTimes().start || Date.now();
  setText(
    el("bookRepeatUnit"),
    pattern === "daily" ? "day(s)" : pattern === "weekly" ? "week(s)" : "month(s)"
  );
  const days = el("bookRepeatDays");
  show(days, pattern === "weekly");
  if (pattern === "weekly") {
    const selected = new Set(
      Array.from(days.querySelectorAll("input:checked")).map(
        (box) => +(box as HTMLInputElement).value
      )
    );
    if (!selected.size) {
      selected.add(new Date(start).getDay());
    }
    days.textContent = "";
    weekdays.forEach((name, index) => {
      const label = document.createElement("label");
      label.className = "checkbox";
      const box = document.createElement("input");
      box.type = "checkbox";
      box.value = `${index}`;
      box.checked = selected.has(index);
      label.appendChild(box);
      label.appendChild(document.createTextNode(` ${name}`));
      days.appendChild(label);
    });
  }
  const week = weekOfMonth(new Date(start));
  const nth = week === -1 ? "last" : ["first", "second", "third", "fourth"][week - 1];
  const unsupported = pattern === "monthly" && monthlyStartUnsupported(new Date(start));
  setText(
    el("bookRepeatMonthly"),
    pattern !== "monthly"
      ? ""
      : unsupported
        ? monthlyStartUnsupportedMessage
        : `On the ${nth} ${new Date(start).toLocaleDateString(undefined, { weekday: "long" })}`,
    unsupported ? "detail error" : "detail"
  );
  const until = input("bookRepeatUntil");
  until.min = dateValue(start);
  until.max = dateValue(Date.now() + maxFutureDays() * 24 * 60 * minute);
  if (!until.value || parseDate(until.value) < parseDate(until.min)) {
    const defaultUntil = new Date(start);
    defaultUntil.setMonth(defaultUntil.getMonth() + (pattern === "daily" ? 1 : 3));
    until.value = dateValue(Math.min(defaultUntil.getTime(), parseDate(until.max)));
  }
}

// ---------------------------------------------------------------------------------------------------------
// Attendees
// ---------------------------------------------------------------------------------------------------------

let searchTimer: number | undefined;
let searchSeq = 0;

function setupAttendeeSearch() {
  const search = input("bookAttendeeSearch");
  const results = el("bookAttendeeResults");
  search.addEventListener("input", () => {
    window.clearTimeout(searchTimer);
    const query = search.value.trim();
    if (query.length < 2) {
      show(results, false);
      return;
    }
    searchTimer = window.setTimeout(() => runSearch(query), 300);
  });
  search.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      const first = results.querySelector("button");
      if (isEmail(search.value)) {
        addAttendee({ name: search.value.trim().split("@")[0], email: search.value.trim() });
      } else if (first && results.style.display !== "none") {
        (first as HTMLButtonElement).click();
      }
    } else if (event.key === "Escape") {
      show(results, false);
    }
  });
}

async function runSearch(query: string) {
  const results = el("bookAttendeeResults");
  const seq = ++searchSeq;
  let people: Person[] = [];
  try {
    people = await searchPeople(api, query, settings);
  } catch (error) {
    console.error(error);
  }
  if (seq !== searchSeq) {
    return;
  }
  results.textContent = "";
  const known = new Set(attendees.map((a) => a.email.toLowerCase()));
  const matches = people.filter((person) => !known.has(person.email.toLowerCase()));
  if (isEmail(query) && !matches.some((p) => p.email.toLowerCase() === query.toLowerCase())) {
    matches.push({ name: query.split("@")[0], email: query });
  }
  for (const person of matches) {
    const item = document.createElement("li");
    const button = document.createElement("button");
    button.type = "button";
    button.className = "suggestion";
    button.textContent =
      person.name && person.name !== person.email
        ? `${person.name} · ${person.email}`
        : person.email;
    button.addEventListener("click", () => addAttendee(person));
    item.appendChild(button);
    results.appendChild(item);
  }
  if (!matches.length) {
    const item = document.createElement("li");
    item.className = "empty";
    item.textContent = "No one found. Type a full email address to invite a guest.";
    results.appendChild(item);
  }
  show(results, true);
}

function addAttendee(person: Person) {
  if (!org) {
    return;
  }
  const email = person.email.trim();
  if (!attendees.some((a) => a.email.toLowerCase() === email.toLowerCase())) {
    attendees.push({ ...person, email, external: org.isExternal(email, settings) });
  }
  input("bookAttendeeSearch").value = "";
  show(el("bookAttendeeResults"), false);
  renderAttendees();
  input("bookAttendeeSearch").focus();
}

function renderAttendees() {
  const list = el("bookAttendees");
  list.textContent = "";
  for (const person of attendees) {
    const chip = document.createElement("li");
    chip.className = "chip";
    chip.textContent = person.name || person.email;
    if (person.external) {
      const tag = document.createElement("span");
      tag.className = "badge";
      tag.textContent = "Guest";
      chip.appendChild(tag);
    }
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "link chip-remove";
    remove.textContent = "×";
    remove.setAttribute("aria-label", `Remove ${person.name || person.email}`);
    remove.addEventListener("click", () => {
      attendees = attendees.filter((a) => a !== person);
      renderAttendees();
    });
    chip.appendChild(remove);
    list.appendChild(chip);
  }
  show(list, attendees.length > 0);
}

// ---------------------------------------------------------------------------------------------------------
// Rooms
// ---------------------------------------------------------------------------------------------------------

function readForm(): Omit<BookingRequest, "room"> | string {
  const title = input("bookTitle").value.trim();
  const allDay = allDayChecked();
  const { start, duration } = formTimes();
  if (!title) {
    return "Add a title.";
  }
  if (allDay && start && start + duration * minute <= Date.now()) {
    return "The all-day period today has ended. Choose another date.";
  }
  // As postForm, bookable hours apply to the all-day period too, in the building's timezone. A start too
  // late for the shortest booking leaves no end times, so check it with the minimum.
  const hoursError =
    start && (duration || !allDay)
      ? checkBookableHours(
          settings,
          start,
          duration || durationLimits().min,
          undefined,
          "Rooms",
          buildingTimezone()
        )
      : "";
  if (hoursError) {
    return hoursError;
  }
  if (!start || !duration) {
    return allDay ? "Choose a date." : "Choose a date and time.";
  }
  if (start + duration * minute < Date.now()) {
    return "That time has already passed.";
  }
  const pattern = select("bookRepeat").value as RecurrencePattern;
  let recurrence: BookingRequest["recurrence"] = null;
  if (pattern !== "none" && el("bookRepeatFields").style.display !== "none") {
    const until = parseDate(input("bookRepeatUntil").value);
    if (isNaN(until) || until < start - 24 * 60 * minute) {
      return "Choose when the meeting stops repeating.";
    }
    const days = Array.from(el("bookRepeatDays").querySelectorAll("input:checked")).map(
      (box) => +(box as HTMLInputElement).value
    );
    if (pattern === "weekly" && !days.length) {
      return "Choose at least one day of the week.";
    }
    if (pattern === "monthly" && monthlyStartUnsupported(new Date(start))) {
      return monthlyStartUnsupportedMessage;
    }
    if (
      allDay &&
      pattern === "daily" &&
      !settings.get<boolean>("events.allow_daily_allday_recurrence")
    ) {
      return "All-day meetings can't repeat daily. Choose another repeat.";
    }
    // A whole day is 23 or 25 hours when daylight saving starts or ends.
    if (duration > 24 * 60 && !allDay) {
      return "Bookings longer than a day can't repeat.";
    }
    recurrence = {
      pattern,
      interval: Math.max(1, Math.min(99, Math.floor(+input("bookRepeatInterval").value) || 1)),
      daysOfWeek: days,
      until,
    };
  }
  return {
    title,
    start,
    duration,
    attendees: [...attendees],
    notes: el<HTMLTextAreaElement>("bookNotes").value.trim(),
    recurrence,
    allDay,
    buildingId,
  };
}

async function findRooms() {
  const form = readForm();
  if (typeof form === "string") {
    setText(el("bookFormError"), form, "error");
    return;
  }
  setText(el("bookFormError"), "");
  request = { ...form, room: null as unknown as Space };
  const button = el<HTMLButtonElement>("bookFindButton");
  button.disabled = true;
  setText(el("bookFormError"), "Finding available rooms...", "detail");
  try {
    const zoneId = settings.get<boolean>("use_region")
      ? org?.buildings.find((b) => b.id === buildingId)?.parent_id || buildingId
      : buildingId;
    searchedBuildingIds =
      zoneId === buildingId
        ? [buildingId]
        : (org?.buildings || []).filter((b) => b.parent_id === zoneId).map((b) => b.id);
    const [spaces, rules] = await Promise.all([
      getSpaces(api, zoneId),
      getBookingRules(api, buildingId),
    ]);
    const groups = org?.user.groups || [];
    const allowed = spaces.filter(
      (space) => !hiddenByRules(rules, space, form.start, form.duration, groups)
    );
    const free = allowed.length
      ? await getAvailableSpaceIds(api, allowed, form.start, form.duration, settings)
      : new Set<string>();
    searchedRooms = allowed;
    candidates = allowed
      .filter((space) => free.has(space.id))
      .sort((a, b) => spaceName(a).localeCompare(spaceName(b)));
    mapSelection = null;
    setText(el("bookFormError"), "");
    prepareFilters();
    renderRooms();
    showStep("rooms");
  } catch (error) {
    console.error(error);
    setText(el("bookFormError"), describeError(error), "error");
  } finally {
    button.disabled = false;
  }
}

function prepareFilters() {
  if (!org || !request) {
    return;
  }
  const levels = select("bookLevel");
  const previousLevel = levels.value;
  levels.textContent = "";
  option(levels, "", "All levels");
  const seen = new Map<string, string>();
  for (const space of candidates) {
    const level = org.levelFor(space.zones);
    if (level) {
      seen.set(level.id, zoneLabel(level));
    }
  }
  Array.from(seen.entries())
    .sort((a, b) => a[1].localeCompare(b[1], undefined, { numeric: true }))
    .forEach(([id, name]) => option(levels, id, name));
  levels.value = seen.has(previousLevel) ? previousLevel : "";
  show(el("bookFavouritesField"), org.favouriteSpaces.length > 0);
  const capacity = input("bookCapacity");
  if (!+capacity.value) {
    capacity.value = `${request.attendees.length ? request.attendees.length + 1 : 0}`;
  }

  const features = el("bookFeatures");
  const checked = new Set(
    Array.from(features.querySelectorAll("input:checked")).map(
      (box) => (box as HTMLInputElement).value
    )
  );
  features.textContent = "";
  const all = Array.from(new Set(candidates.flatMap(spaceFeatures))).sort();
  for (const feature of all) {
    const label = document.createElement("label");
    label.className = "checkbox";
    const box = document.createElement("input");
    box.type = "checkbox";
    box.value = feature;
    box.checked = checked.has(feature);
    box.addEventListener("change", renderRooms);
    label.appendChild(box);
    label.appendChild(document.createTextNode(` ${feature}`));
    features.appendChild(label);
  }
  setText(el("bookRoomsSummary"), `${whenText(request)}${request.recurrence ? ", repeating" : ""}`);
  prepareMapLevels();
}

/** Free rooms that match the filters. The map shows one level itself, so it ignores the Level filter. */
function filteredRooms(ignoreLevel = roomView === "map"): Space[] {
  if (!org) {
    return [];
  }
  const level = ignoreLevel ? "" : select("bookLevel").value;
  const capacity = +input("bookCapacity").value || 0;
  const favourites = input("bookFavourites").checked ? org.favouriteSpaces : null;
  const features = Array.from(el("bookFeatures").querySelectorAll("input:checked")).map(
    (box) => (box as HTMLInputElement).value
  );
  return candidates.filter(
    (space) =>
      (!level || (space.zones || []).includes(level)) &&
      // As Workplace's Space, a capacity of 0 means unknown, so it matches any filter.
      (capacity <= 0 || (space.capacity || -1) < 0 || (space.capacity || 0) >= capacity) &&
      (!favourites || favourites.includes(space.id)) &&
      features.every((feature) => spaceFeatures(space).includes(feature))
  );
}

function roomLocation(space: Space): string {
  const level = org?.levelFor(space.zones);
  const building = org?.buildings.find((b) => b.id === buildingId);
  return [level ? zoneLabel(level) : "", building ? zoneLabel(building) : ""]
    .filter(Boolean)
    .join(", ");
}

function renderRooms() {
  const list = el("bookRoomList");
  list.textContent = "";
  const rooms = filteredRooms();
  const filtersOn = rooms.length !== candidates.length;
  setText(
    el("bookFiltersToggle"),
    filtersOn ? `Filters (${candidates.length - rooms.length} hidden)` : "Filters"
  );
  if (roomView === "map") {
    renderMap();
    return;
  }
  if (!rooms.length) {
    const item = document.createElement("li");
    item.className = "empty";
    item.textContent = candidates.length
      ? "No rooms match these filters."
      : "No rooms are free at this time. Try another time or building.";
    list.appendChild(item);
    return;
  }
  const favourites = org?.favouriteSpaces || [];
  for (const space of rooms) {
    const item = document.createElement("li");
    const button = document.createElement("button");
    button.type = "button";
    button.className = "room";
    const title = document.createElement("div");
    title.className = "item-title";
    title.textContent = `${favourites.includes(space.id) ? "★ " : ""}${spaceName(space)}`;
    button.appendChild(title);
    const details = [
      roomLocation(space),
      (space.capacity ?? -1) > 0 ? `Seats ${space.capacity}` : "",
      spaceFeatures(space).slice(0, 4).join(", "),
    ].filter(Boolean);
    for (const line of details) {
      const detail = document.createElement("div");
      detail.className = "item-detail";
      detail.textContent = line;
      button.appendChild(detail);
    }
    button.addEventListener("click", () => chooseRoom(space));
    item.appendChild(button);
    list.appendChild(item);
  }
}

// ---------------------------------------------------------------------------------------------------------
// Rooms map (legacy interactive-map): one level's floor plan, free rooms coloured and tappable
// ---------------------------------------------------------------------------------------------------------

function setRoomView(view: "list" | "map") {
  roomView = view;
  const list = view === "list";
  for (const [id, pressed] of [
    ["bookListViewButton", list],
    ["bookMapViewButton", !list],
  ] as [string, boolean][]) {
    el(id).setAttribute("aria-pressed", `${pressed}`);
    el(id).className = pressed ? "" : "secondary";
  }
  show(el("bookRoomList"), list);
  show(el("bookMapView"), !list);
  // The map has its own level picker.
  show(el("bookLevelField"), list);
  if (!list) {
    chooseDefaultMapLevel();
  }
  renderRooms();
}

/** Fills the map's level picker with the searched buildings' levels that have a floor plan. */
function prepareMapLevels() {
  if (!org) {
    return;
  }
  mapLevels = org.mapLevels(searchedBuildingIds);
  const picker = select("bookMapLevel");
  const previous = picker.value;
  picker.textContent = "";
  mapLevels.forEach((level) => option(picker, level.id, ""));
  picker.value = mapLevels.some((level) => level.id === previous) ? previous : "";
  updateMapLevelLabels();
  show(el("bookViewToggle"), mapLevels.length > 0);
  if (!mapLevels.length && roomView === "map") {
    setRoomView("list");
  } else if (roomView === "map" && !picker.value) {
    chooseDefaultMapLevel();
  }
}

/** "Level 3 (2 free)", with the building when rooms come from a region. */
function updateMapLevelLabels() {
  const free = filteredRooms(true);
  const options = select("bookMapLevel").options;
  mapLevels.forEach((level, index) => {
    const count = free.filter((space) => (space.zones || []).includes(level.id)).length;
    const building =
      searchedBuildingIds.length > 1
        ? org?.buildings.find((b) => b.id === level.parent_id)
        : undefined;
    options[index].textContent =
      `${building ? `${zoneLabel(building)}, ` : ""}${zoneLabel(level)} (${count} free)`;
  });
}

/** The Level filter's level if it has a map, else the current pick, else the level with most free rooms. */
function chooseDefaultMapLevel() {
  const picker = select("bookMapLevel");
  const filterLevel = select("bookLevel").value;
  if (filterLevel && mapLevels.some((level) => level.id === filterLevel)) {
    picker.value = filterLevel;
    return;
  }
  if (picker.value) {
    return;
  }
  const free = filteredRooms(true);
  let best: Zone | undefined;
  let bestCount = -1;
  for (const level of mapLevels) {
    const count = free.filter((space) => (space.zones || []).includes(level.id)).length;
    if (count > bestCount) {
      best = level;
      bestCount = count;
    }
  }
  picker.value = best?.id || "";
}

/** Colours from app.explore.colors, as Workplace's space map, else its defaults. */
function mapColours(): Record<MapRoomStatus, string> {
  const colours = settings.get<Record<string, string>>("explore.colors", {}) || {};
  const pick = (status: MapRoomStatus, key: string) =>
    colours[`space-${key}`] || colours[key] || defaultMapColours[status];
  return {
    free: pick("free", "free"),
    busy: pick("busy", "busy"),
    filtered: pick("filtered", "not-bookable"),
  };
}

const mapElementId = (space: Space) => space.map_id || space.id;

async function renderMap() {
  if (!floorMap) {
    return;
  }
  updateMapLevelLabels();
  const mapStatus = el("bookMapStatus");
  const level = mapLevels.find((l) => l.id === select("bookMapLevel").value);
  if (!level?.map_id) {
    floorMap.clear();
    shownMapUrl = "";
    setText(mapStatus, "No floor plans are set up for this building.", "detail");
    renderMapSelection();
    return;
  }
  const seq = ++mapLoadSeq;
  if (shownMapUrl !== level.map_id) {
    setText(mapStatus, "Loading the floor plan...", "detail");
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

  const free = new Set(filteredRooms(true).map((space) => space.id));
  const freeAnyFilter = new Set(candidates.map((space) => space.id));
  const onLevel = searchedRooms.filter((space) => (space.zones || []).includes(level.id));
  const rooms: MapRoom[] = onLevel.map((space) => {
    const status: MapRoomStatus = free.has(space.id)
      ? "free"
      : freeAnyFilter.has(space.id)
        ? "filtered"
        : "busy";
    const seats = (space.capacity ?? -1) > 0 ? `, seats ${space.capacity}` : "";
    const label = { free: "free", busy: "busy", filtered: "doesn't match the filters" }[status];
    return {
      elementId: mapElementId(space),
      status,
      title: `${spaceName(space)}${seats} (${label})`,
      selectable: status === "free",
    };
  });
  const colours = mapColours();
  const accent =
    window.getComputedStyle(document.documentElement).getPropertyValue("--accent").trim() ||
    "#0078d4";
  const missing = new Set(floorMap.setRooms(rooms, colours, accent));
  if (mapSelection && !(free.has(mapSelection.id) && onLevel.includes(mapSelection))) {
    mapSelection = null;
  }
  floorMap.select(mapSelection ? mapElementId(mapSelection) : null);

  const freeHere = onLevel.filter((space) => free.has(space.id));
  const unplaced = freeHere.filter((space) => missing.has(mapElementId(space))).length;
  setText(
    mapStatus,
    [
      freeHere.length
        ? "Tap a green room to choose it."
        : candidates.length
          ? "No free rooms on this level match your filters. Try another level."
          : "No rooms are free at this time. Try another time or building.",
      unplaced
        ? `${unplaced} free room${unplaced === 1 ? " isn't" : "s aren't"} on this floor plan; use the list to see ${unplaced === 1 ? "it" : "them"}.`
        : "",
    ]
      .filter(Boolean)
      .join(" "),
    "detail"
  );
  renderMapLegend(
    colours,
    rooms.some((room) => room.status === "filtered")
  );
  renderMapSelection();
}

function renderMapLegend(colours: Record<MapRoomStatus, string>, showFiltered: boolean) {
  const legend = el("bookMapLegend");
  legend.textContent = "";
  const entries: [MapRoomStatus, string][] = [
    ["free", "Free"],
    ["busy", "Busy"],
  ];
  if (showFiltered) {
    entries.push(["filtered", "Doesn't match filters"]);
  }
  for (const [status, text] of entries) {
    const item = document.createElement("li");
    const swatch = document.createElement("span");
    swatch.className = "swatch";
    swatch.style.backgroundColor = colours[status];
    item.appendChild(swatch);
    item.appendChild(document.createTextNode(text));
    legend.appendChild(item);
  }
}

function selectOnMap(elementId: string) {
  const level = select("bookMapLevel").value;
  mapSelection =
    filteredRooms(true).find(
      (space) => mapElementId(space) === elementId && (space.zones || []).includes(level)
    ) || null;
  floorMap?.select(mapSelection ? elementId : null);
  renderMapSelection();
  if (mapSelection) {
    el("bookMapSelection").scrollIntoView({ block: "nearest" });
  }
}

/** The tapped room's details and a button to choose it (the legacy add-in's room tile). */
function renderMapSelection() {
  const card = el("bookMapSelection");
  card.textContent = "";
  const space = mapSelection;
  show(card, !!space);
  if (!space) {
    return;
  }
  const favourite = org?.favouriteSpaces.includes(space.id);
  line(card, `${favourite ? "★ " : ""}${spaceName(space)}`, "item-title");
  line(card, roomLocation(space));
  line(card, (space.capacity ?? -1) > 0 ? `Seats ${space.capacity}` : "");
  line(card, spaceFeatures(space).slice(0, 6).join(", "));
  const button = document.createElement("button");
  button.type = "button";
  button.className = "card-action";
  button.textContent = "Choose this room";
  button.addEventListener("click", () => chooseRoom(space));
  card.appendChild(button);
}

// ---------------------------------------------------------------------------------------------------------
// Confirm: room details, summary, catering, assets
// ---------------------------------------------------------------------------------------------------------

function line(parent: HTMLElement, text: string, className = "item-detail") {
  if (!text) {
    return;
  }
  const div = document.createElement("div");
  div.className = className;
  div.textContent = text;
  parent.appendChild(div);
}

function recurrenceText(): string {
  const recurrence = request?.recurrence;
  if (!recurrence || !request) {
    return "";
  }
  const every = recurrence.interval > 1 ? `Every ${recurrence.interval} ` : "Every ";
  const unit = { daily: "day", weekly: "week", monthly: "month", none: "" }[recurrence.pattern];
  const days =
    recurrence.pattern === "weekly"
      ? ` on ${recurrence.daysOfWeek.map((d) => weekdays[d]).join(", ")}`
      : recurrence.pattern === "monthly"
        ? ` (${el("bookRepeatMonthly").textContent?.toLowerCase()})`
        : "";
  return `${every}${unit}${recurrence.interval > 1 ? "s" : ""}${days}, until ${new Date(
    recurrence.until
  ).toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" })}`;
}

function chooseRoom(space: Space) {
  if (!request) {
    return;
  }
  request.room = space;
  acceptClashes = false;
  setText(el("bookConfirmError"), "");
  show(el("bookAnywayButton"), false);

  const details = el("bookRoomDetails");
  details.textContent = "";
  if (space.images?.[0] && /^https:\/\//.test(space.images[0])) {
    const image = document.createElement("img");
    image.className = "room-image";
    image.src = space.images[0];
    image.alt = "";
    details.appendChild(image);
  }
  line(details, spaceName(space), "item-title");
  line(details, roomLocation(space));
  line(details, (space.capacity ?? -1) > 0 ? `Seats ${space.capacity}` : "");
  line(details, spaceFeatures(space).join(", "));
  line(details, space.description || "");

  const summary = el("bookSummary");
  summary.textContent = "";
  line(summary, request.title, "item-title");
  line(summary, whenText(request));
  line(summary, recurrenceText());
  const guests = request.attendees.filter((a) => a.external).length;
  line(
    summary,
    request.attendees.length
      ? `${request.attendees.length} attendee(s)${guests && !draft ? `, ${guests} guest(s) will be expected at reception` : ""}`
      : "No other attendees"
  );
  // Catering, assets and visitors are PlaceOS bookings linked to the event, which doesn't exist until
  // the draft is sent and synced.
  line(
    summary,
    draft
      ? "Catering, assets and visitor registration aren't available when adding a room to an Outlook meeting."
      : ""
  );

  showStep("confirm");
  if (draft) {
    cateringSelections = [];
    assetSelections = [];
    assetsAllowed = false;
    show(el("bookCatering"), false);
    show(el("bookAssets"), false);
    return;
  }
  loadCatering(space);
  loadAssets(space);
}

/** Delivery time choices (deliveryOffsets) as "N min after the start (time)". */
function offsetOptions(
  target: HTMLSelectElement,
  { min, max, step: offsetStep }: { min: number; max: number; step: number }
) {
  const previous = target.value;
  target.textContent = "";
  for (let offset = min; offset <= max; offset += offsetStep) {
    option(
      target,
      `${offset}`,
      offset === 0
        ? `At the start (${formatTime(request!.start)})`
        : `${formatDuration(offset)} after the start (${formatTime(request!.start + offset * minute)})`
    );
  }
  target.value = previous && +previous <= max && +previous >= min ? previous : `${min}`;
  if (!target.value && target.options.length) {
    target.selectedIndex = 0;
  }
}

let cateringSeq = 0;

/** The room's building (its level's parent), where Workplace looks for its menu, assets and rules. */
const roomBuildingId = (space: Space) => org?.levelFor(space.zones)?.parent_id || buildingId;

async function loadCatering(space: Space) {
  const section = el("bookCatering");
  const menu = el("bookCateringMenu");
  cateringSelections = [];
  cateringSettings = null;
  menu.textContent = "";
  show(el("bookCateringExtras"), false);
  show(el("bookCateringFilterBar"), false);
  show(el("bookCateringFilters"), false);
  if (!request || !org || !cateringEnabled(settings)) {
    show(section, false);
    return;
  }
  const seq = ++cateringSeq;
  const { start, duration } = request;
  show(section, true);
  setText(el("bookCateringStatus"), "Loading the menu...", "detail");
  try {
    const zoneId = roomBuildingId(space);
    const building = org.buildings.find((b) => b.id === buildingId) as
      { currency?: string } | undefined;
    const provider = cateringProvider(settings);
    const [menuItems, cateringConfig] = await Promise.all([
      getCateringMenu(api, zoneId, space, start, duration),
      getCateringSettings(api, buildingId, settings, building?.currency),
    ]);
    if (seq !== cateringSeq) {
      return;
    }
    // With app.catering_provider set, only that caterer's items are offered.
    const items = provider ? menuItems.filter((item) => item.caterer === provider) : menuItems;
    if (!items.length || cateringConfig.disabledRooms.includes(space.id)) {
      // No menu for this building (or nothing orderable this close to the meeting), or the room is in
      // "catering-settings" disabled_rooms.
      show(section, false);
      return;
    }
    cateringSettings = cateringConfig;
    setText(el("bookCateringStatus"), "Optional. Choose quantities to add catering.", "detail");
    cateringSelections = items.map((item) => ({ item, quantity: 0, optionIds: [] }));
    resetCateringFilters();
    renderCateringMenu();
    offsetOptions(select("bookCateringDeliver"), deliveryOffsets(settings, "catering", duration));
    const codes = select("bookChargeCode");
    codes.textContent = "";
    option(codes, "", "None");
    cateringConfig.chargeCodes.forEach((code) => option(codes, code, code));
    show(el("bookChargeCodeField"), cateringConfig.chargeCodes.length > 0);
    show(el("bookCateringNotesRequired"), cateringConfig.requireNotes);
    renderCateringTotal();
  } catch (error) {
    if (seq !== cateringSeq) {
      return;
    }
    console.error(error);
    setText(
      el("bookCateringStatus"),
      `Catering isn't available right now. ${describeError(error)}`,
      "detail error"
    );
  }
}

// ---------------------------------------------------------------------------------------------------------
// Catering and asset filters and favourites
// ---------------------------------------------------------------------------------------------------------

/** Search box, Filters toggle and favourites checkbox for "bookCatering" or "bookAssets". */
function setupMenuFilters(prefix: string, render: () => void) {
  input(`${prefix}Search`).addEventListener("input", render);
  input(`${prefix}Favourites`).addEventListener("change", render);
  el(`${prefix}FiltersToggle`).addEventListener("click", () => {
    const filters = el(`${prefix}Filters`);
    show(filters, filters.style.display === "none");
  });
}

/** Clears the search, filters and favourites checkbox, and shows the filter bar. */
function resetMenuFilters(prefix: string) {
  input(`${prefix}Search`).value = "";
  input(`${prefix}Favourites`).checked = false;
  show(el(`${prefix}Filters`), false);
  show(el(`${prefix}FilterBar`), true);
}

/** Replaces a container's checkboxes with one per value. The field is hidden if there's nothing to pick. */
function checkboxGroup(
  field: HTMLElement,
  container: HTMLElement,
  values: string[],
  onChange: () => void
) {
  container.textContent = "";
  for (const value of values) {
    const label = document.createElement("label");
    label.className = "checkbox";
    const box = document.createElement("input");
    box.type = "checkbox";
    box.value = value;
    box.addEventListener("change", onChange);
    label.appendChild(box);
    label.appendChild(document.createTextNode(` ${value}`));
    container.appendChild(label);
  }
  show(field, values.length > 0);
}

const checkedValues = (container: HTMLElement) =>
  Array.from(container.querySelectorAll("input:checked")).map(
    (box) => (box as HTMLInputElement).value
  );

const unique = (values: string[]) =>
  Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b));

/** "Filters" or "Filters (3 hidden)". */
function updateFiltersToggle(prefix: string, hidden: number) {
  setText(el(`${prefix}FiltersToggle`), hidden ? `Filters (${hidden} hidden)` : "Filters");
}

/**
 * A star button that adds or removes the id in a favourites list in the user's PlaceOS settings, as
 * Workplace's heart buttons do. onChange runs after the change (and again if saving fails and it's undone).
 */
function favouriteButton(
  listName: string,
  id: string,
  name: string,
  statusElement: HTMLElement,
  onChange: () => void
): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "favourite";
  const update = () => {
    const on = !!org?.userList(listName).includes(id);
    button.textContent = on ? "★" : "☆";
    button.setAttribute("aria-pressed", `${on}`);
    button.title = on ? "Remove from favourites" : "Add to favourites";
    button.setAttribute("aria-label", `Favourite ${name}`);
  };
  update();
  button.addEventListener("click", async () => {
    if (!org) {
      return;
    }
    const saving = org.toggleFavourite(listName, id);
    update();
    onChange();
    try {
      await saving;
    } catch (error) {
      console.error(error);
      update();
      onChange();
      setText(
        statusElement,
        `Couldn't save your favourites. ${describeError(error)}`,
        "detail error"
      );
    }
  });
  return button;
}

function resetCateringFilters() {
  resetMenuFilters("bookCatering");
  const items = cateringSelections.map((s) => s.item);
  const caterers = unique(items.map((item) => item.caterer));
  const caterer = select("bookCaterer");
  caterer.textContent = "";
  option(caterer, "", "All caterers");
  caterers.forEach((name) => option(caterer, name, name));
  // Only one caterer when app.catering_provider is set, as items were already filtered to it.
  show(el("bookCatererField"), caterers.length > 1);
  const categories = unique(items.map((item) => item.category));
  checkboxGroup(
    el("bookCateringCategoriesField"),
    el("bookCateringCategories"),
    categories.length > 1 ? categories : [],
    renderCateringMenu
  );
  checkboxGroup(
    el("bookCateringTagsField"),
    el("bookCateringTags"),
    unique(items.flatMap((item) => item.tags)),
    renderCateringMenu
  );
}

/** Menu items matching the filters. Items with a quantity always stay, so nothing chosen is hidden. */
function visibleCatering(): CateringSelection[] {
  const matching = new Set(
    filterCateringMenu(
      cateringSelections.map((s) => s.item),
      {
        search: input("bookCateringSearch").value,
        caterer: select("bookCaterer").value,
        categories: checkedValues(el("bookCateringCategories")),
        tags: checkedValues(el("bookCateringTags")),
        favouritesOnly: input("bookCateringFavourites").checked,
      },
      org?.userList(favouriteMenuItems) || []
    ).map((item) => item.id)
  );
  return cateringSelections.filter((s) => matching.has(s.item.id) || s.quantity > 0);
}

function renderCateringMenu() {
  const menu = el("bookCateringMenu");
  menu.textContent = "";
  const currency = cateringSettings?.currency || "USD";
  const visible = visibleCatering();
  updateFiltersToggle("bookCatering", cateringSelections.length - visible.length);
  if (!visible.length) {
    line(menu, "No menu items match these filters.", "detail");
    return;
  }
  const categories = Array.from(new Set(visible.map((s) => s.item.category))).sort();
  for (const category of categories) {
    const heading = document.createElement("h3");
    heading.className = "menu-category";
    heading.textContent = category;
    menu.appendChild(heading);
    const list = document.createElement("ul");
    list.className = "items";
    for (const selection of visible.filter((s) => s.item.category === category)) {
      const item = document.createElement("li");
      const row = document.createElement("div");
      row.className = "menu-row";
      const text = document.createElement("div");
      line(text, selection.item.name, "item-title");
      line(text, selection.item.description || "");
      line(text, formatPrice(selection.item.unit_price, currency));
      row.appendChild(text);
      const actions = document.createElement("div");
      actions.className = "menu-row-actions";
      actions.appendChild(
        favouriteButton(
          favouriteMenuItems,
          selection.item.id,
          selection.item.name,
          el("bookCateringStatus"),
          () => input("bookCateringFavourites").checked && renderCateringMenu()
        )
      );
      const quantity = document.createElement("input");
      quantity.type = "number";
      quantity.min = "0";
      quantity.max = "99";
      quantity.className = "quantity";
      quantity.value = `${selection.quantity}`;
      quantity.setAttribute("aria-label", `Quantity of ${selection.item.name}`);
      quantity.addEventListener("change", () => {
        selection.quantity = Math.max(0, Math.min(99, Math.floor(+quantity.value) || 0));
        quantity.value = `${selection.quantity}`;
        renderOptions(selection, options);
        renderCateringTotal();
      });
      actions.appendChild(quantity);
      row.appendChild(actions);
      item.appendChild(row);
      const options = document.createElement("div");
      options.className = "menu-options";
      item.appendChild(options);
      renderOptions(selection, options);
      list.appendChild(item);
    }
    menu.appendChild(list);
  }
}

function renderOptions(selection: CateringSelection, container: HTMLElement) {
  container.textContent = "";
  if (!selection.quantity || !selection.item.options.length) {
    return;
  }
  const currency = cateringSettings?.currency || "USD";
  for (const group of optionGroups(selection.item)) {
    const fieldset = document.createElement("div");
    line(fieldset, group.name, "item-detail");
    for (const choice of group.options) {
      const label = document.createElement("label");
      label.className = "checkbox";
      const box = document.createElement("input");
      box.type = group.multiple ? "checkbox" : "radio";
      box.name = `${selection.item.id}-${group.name}`;
      box.checked = selection.optionIds.includes(choice.id);
      box.addEventListener("change", () => {
        const others = group.multiple ? [] : group.options.map((o) => o.id);
        selection.optionIds = selection.optionIds.filter(
          (id) => id !== choice.id && !others.includes(id)
        );
        if (box.checked) {
          selection.optionIds.push(choice.id);
        }
        renderCateringTotal();
      });
      label.appendChild(box);
      label.appendChild(
        document.createTextNode(
          ` ${choice.name}${choice.unit_price ? ` (+${formatPrice(choice.unit_price, currency)})` : ""}`
        )
      );
      fieldset.appendChild(label);
    }
    container.appendChild(fieldset);
  }
}

function renderCateringTotal() {
  const chosen = cateringSelections.filter((s) => s.quantity > 0);
  show(el("bookCateringExtras"), chosen.length > 0);
  const total = chosen.reduce((sum, s) => sum + unitPriceWithOptions(s) * s.quantity, 0);
  const count = chosen.reduce((sum, s) => sum + s.quantity, 0);
  setText(
    el("bookCateringTotal"),
    chosen.length
      ? `${count} item(s), ${formatPrice(total, cateringSettings?.currency || "USD")}`
      : ""
  );
}

let assetsSeq = 0;

async function loadAssets(space: Space) {
  const section = el("bookAssets");
  const list = el("bookAssetList");
  assetSelections = [];
  assetsAllowed = false;
  list.textContent = "";
  show(el("bookAssetsDeliverField"), false);
  show(el("bookAssetsFilterBar"), false);
  show(el("bookAssetsFilters"), false);
  if (!request || !assetsEnabled(settings)) {
    show(section, false);
    return;
  }
  const seq = ++assetsSeq;
  const { start, duration } = request;
  show(section, true);
  setText(el("bookAssetsStatus"), "Checking what's available...", "detail");
  try {
    if (await assetsDisabledForRoom(api, buildingId, space)) {
      if (seq === assetsSeq) {
        setText(el("bookAssetsStatus"), "Assets can't be requested for this room.", "detail");
      }
      return;
    }
    const types = await getAvailableAssetTypes(api, roomBuildingId(space), space, start, duration);
    if (seq !== assetsSeq) {
      return;
    }
    if (!types.length) {
      setText(el("bookAssetsStatus"), "No assets are available at this time.", "detail");
      return;
    }
    assetsAllowed = true;
    setText(el("bookAssetsStatus"), "Optional. Choose quantities to request equipment.", "detail");
    assetSelections = types.map((type) => ({ type, quantity: 0 }));
    resetMenuFilters("bookAssets");
    const categories = unique(types.map((type) => type.category));
    checkboxGroup(
      el("bookAssetsCategoriesField"),
      el("bookAssetsCategories"),
      categories.length > 1 ? categories : [],
      renderAssetList
    );
    renderAssetList();
    offsetOptions(select("bookAssetsDeliver"), deliveryOffsets(settings, "assets", duration));
  } catch (error) {
    if (seq !== assetsSeq) {
      return;
    }
    console.error(error);
    setText(
      el("bookAssetsStatus"),
      `Assets aren't available right now. ${describeError(error)}`,
      "detail error"
    );
  }
}

/** Asset types matching the filters. Requested ones always stay, so nothing chosen is hidden. */
function renderAssetList() {
  const list = el("bookAssetList");
  list.textContent = "";
  const matching = new Set(
    filterAssetTypes(
      assetSelections.map((s) => s.type),
      {
        search: input("bookAssetsSearch").value,
        categories: checkedValues(el("bookAssetsCategories")),
        favouritesOnly: input("bookAssetsFavourites").checked,
      },
      org?.userList(favouriteAssets) || []
    ).map((type) => type.id)
  );
  const visible = assetSelections.filter((s) => matching.has(s.type.id) || s.quantity > 0);
  updateFiltersToggle("bookAssets", assetSelections.length - visible.length);
  if (!visible.length) {
    const item = document.createElement("li");
    item.className = "empty";
    item.textContent = "No assets match these filters.";
    list.appendChild(item);
    return;
  }
  visible.forEach((selection) => list.appendChild(assetRow(selection)));
}

function assetRow(selection: AssetSelection): HTMLElement {
  const item = document.createElement("li");
  const row = document.createElement("div");
  row.className = "menu-row";
  const text = document.createElement("div");
  line(text, selection.type.name, "item-title");
  line(text, [selection.type.category, selection.type.brand].filter(Boolean).join(" · "));
  line(text, `${selection.type.available.length} available`);
  row.appendChild(text);
  const actions = document.createElement("div");
  actions.className = "menu-row-actions";
  actions.appendChild(
    favouriteButton(
      favouriteAssets,
      selection.type.id,
      selection.type.name,
      el("bookAssetsStatus"),
      () => input("bookAssetsFavourites").checked && renderAssetList()
    )
  );
  const quantity = document.createElement("input");
  quantity.type = "number";
  quantity.min = "0";
  quantity.max = `${selection.type.available.length}`;
  quantity.className = "quantity";
  quantity.value = `${selection.quantity}`;
  quantity.setAttribute("aria-label", `Quantity of ${selection.type.name}`);
  quantity.addEventListener("change", () => {
    selection.quantity = Math.max(
      0,
      Math.min(selection.type.available.length, Math.floor(+quantity.value) || 0)
    );
    quantity.value = `${selection.quantity}`;
    show(
      el("bookAssetsDeliverField"),
      assetSelections.some((s) => s.quantity > 0)
    );
  });
  actions.appendChild(quantity);
  row.appendChild(actions);
  item.appendChild(row);
  return item;
}

// ---------------------------------------------------------------------------------------------------------
// Submit
// ---------------------------------------------------------------------------------------------------------

async function submit() {
  if (!request || !org || submitting) {
    return;
  }
  const errorText = el("bookConfirmError");
  const room = request.room;
  const roomName = spaceName(room);
  const chosenCatering = cateringSelections.filter((s) => s.quantity > 0);
  const cateringNotes = el<HTMLTextAreaElement>("bookCateringNotes").value.trim();
  if (chosenCatering.length && cateringSettings?.requireNotes && !cateringNotes) {
    setText(errorText, "Add catering notes.", "error");
    return;
  }
  submitting = true;
  const buttons = ["bookSubmitButton", "bookAnywayButton", "bookOtherRoom"].map((id) =>
    el<HTMLButtonElement>(id)
  );
  buttons.forEach((button) => (button.disabled = true));
  setText(errorText, draft ? "Adding the room..." : "Booking...", "detail");
  try {
    const { start, duration } = request;
    // Someone may have taken the room since the list loaded.
    const stillFree = await getAvailableSpaceIds(api, [room], start, duration, settings);
    if (!stillFree.has(room.id)) {
      throw new Error("This room has just been booked by someone else. Choose another room.");
    }
    if (draft) {
      await draft.fill(request);
      showDone(0, 0);
      return;
    }
    const orders = chosenCatering.length
      ? buildCateringOrders(
          chosenCatering,
          room,
          start,
          +select("bookCateringDeliver").value || 0,
          cateringNotes,
          select("bookChargeCode").value
        )
      : [];
    const chosenAssets = assetsAllowed ? assetSelections.filter((s) => s.quantity > 0) : [];
    const assetRequest = chosenAssets.length
      ? buildAssetRequest(
          chosenAssets,
          await getAvailableAssetTypes(api, roomBuildingId(room), room, start, duration),
          start,
          +select("bookAssetsDeliver").value || 0
        )
      : null;

    const body = buildEventBody(request, org, settings, {
      catering: orders,
      assets: assetRequest ? [assetRequest] : [],
    });

    if (request.recurrence && !acceptClashes) {
      const clashes = await findClashes(api, body);
      if (clashes.length) {
        if (clashes[0].booking_start === body.event_start) {
          throw new Error("The room is already booked for the first meeting in this series.");
        }
        if (!settings.get<boolean>("events.allow_recurring_instance_clashes")) {
          throw new Error(
            `The room is already booked for ${clashes.length} of the repeating meetings. Change the time or the repeat.`
          );
        }
        const dates = clashes
          .slice(0, 5)
          .map((clash) => new Date(clash.booking_start * 1000).toLocaleDateString())
          .join(", ");
        setText(
          errorText,
          `The room isn't free for ${clashes.length} of the repeating meetings (${dates}${
            clashes.length > 5 ? ", ..." : ""
          }). Those meetings won't have the room.`,
          "error"
        );
        show(el("bookAnywayButton"), true);
        return;
      }
    }

    const event = await createEvent(api, body);
    const linked = [
      ...visitorBookings(request, org, body.host),
      ...cateringBookings(orders, room, roomName),
      ...(assetRequest ? [assetBooking(assetRequest, room, roomName, org.user.email)] : []),
    ];
    if (linked.length) {
      try {
        await createLinkedBookings(
          api,
          { ...event, title: event.title || body.title },
          {
            user_email: body.host,
            zones: bookingZones(org, buildingId, room),
            timezone: body.timezone,
          },
          linked
        );
      } catch (error) {
        console.error(error);
        // As postForm's rollback: the organiser's calendar, not a forced or room host.
        await removeEvent(api, event, org.user.email, room.id).catch((e) => console.error(e));
        throw new Error(
          error instanceof PlaceosRequestError && error.status === 409
            ? "Some of the catering or assets were just booked by someone else, so the booking was cancelled. Try again."
            : "The visitor, catering or asset bookings couldn't be made, so the booking was cancelled. Try again."
        );
      }
    }
    showDone(orders.length, assetRequest?.item_count || 0);
  } catch (error) {
    console.error(error);
    setText(errorText, describeError(error), "error");
  } finally {
    submitting = false;
    buttons.forEach((button) => (button.disabled = false));
  }
}

function showDone(cateringOrders: number, assetCount: number) {
  if (!request) {
    return;
  }
  const summary = el("bookDoneSummary");
  summary.textContent = "";
  line(summary, draft ? "✓ Room added to your meeting" : "✓ Room booked", "item-title ok");
  line(summary, `${request.title} in ${spaceName(request.room)}`);
  line(summary, whenText(request));
  line(summary, recurrenceText());
  line(summary, cateringOrders ? "Catering ordered." : "");
  line(summary, assetCount ? `${assetCount} asset(s) requested.` : "");
  line(
    summary,
    draft
      ? "Send the invitation to book the room. The room accepts or declines it by email."
      : "The meeting will appear in your calendar shortly."
  );
  request = null;
  showStep("done");
}
