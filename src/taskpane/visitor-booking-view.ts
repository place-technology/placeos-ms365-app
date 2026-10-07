/*
 * The Invite a visitor view: visitor details and time → done. Mirrors PlaceOS Workplace's invite visitor form
 * for one visitor; the data side is in visitor-booking-data.ts.
 */

/* global document, console, window, HTMLElement, HTMLInputElement, HTMLSelectElement, HTMLButtonElement */

import {
  loadOrganisation,
  zoneLabel,
  type AppSettings,
  type Organisation,
} from "./booking-settings";
import type { PlaceosApi } from "./placeos-data";
import { PlaceosRequestError } from "./placeos-helper";
import {
  allDayPeriod,
  checkBookableHours,
  dateValue,
  isEmail,
  isFullDay,
} from "./room-booking-data";
import {
  buildVisitorBooking,
  checkBeforeInvite,
  createVisitorBooking,
  pastVisitors,
  rememberVisitor,
  visitorAllDayAllowed,
  visitorDurationLimits,
  visitorSetting,
  type PastVisitor,
  type VisitorRequest,
} from "./visitor-booking-data";

const el = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
const input = (id: string) => el<HTMLInputElement>(id);
const select = (id: string) => el<HTMLSelectElement>(id);
const show = (element: HTMLElement, visible: boolean) =>
  (element.style.display = visible ? "" : "none");

const status = el("visitorStatus");
const steps = {
  form: el("visitorForm"),
  done: el("visitorDone"),
};
type Step = keyof typeof steps;

const minute = 60000;
// Start times are offered every 15 minutes, end times every 15 minutes after the shortest visit.
const startStep = 15;
const durationStep = 15;
// How far ahead visitors can be invited (Workplace's date field allows a year).
const maxFutureDays = 365;

let api: PlaceosApi;
let onClose: (booked: boolean) => void = () => undefined;
let org: Organisation | null = null;
let settings: AppSettings;
let buildingId = "";
let submitting = false;
let formReady = false;

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

/** The "error" or "message" text of a JSON error body, as Workplace's errorMessage reads it, or "". */
function serverMessage(body: string): string {
  try {
    const json: unknown = JSON.parse(body);
    if (json && typeof json === "object") {
      const { error, message } = json as { error?: unknown; message?: unknown };
      return typeof error === "string" ? error : typeof message === "string" ? message : "";
    }
  } catch {
    // Not JSON (e.g. a proxy error page).
  }
  return "";
}

function describeError(error: unknown): string {
  if (error instanceof PlaceosRequestError) {
    const denied = error.status === 401 || error.status === 403;
    const text = denied
      ? "PlaceOS didn't allow this invite."
      : `PlaceOS couldn't complete this (${error.status}).`;
    // Workplace shows the Staff API's message, so add it when there is one.
    return [text, serverMessage(error.body) || (denied ? "" : "Try again.")]
      .filter(Boolean)
      .join(" ");
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
function whenText({
  start,
  duration,
  allDay,
}: Pick<VisitorRequest, "start" | "duration" | "allDay">) {
  const times =
    allDay && isFullDay(start, duration)
      ? "all day"
      : `${formatTime(start)} – ${formatTime(start + duration * minute)}`;
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

/** Wires up the view once. onDone is called when the user leaves it (booked = an invite was sent). */
export function initVisitorBooking(placeosApi: PlaceosApi, onDone: (booked: boolean) => void) {
  api = placeosApi;
  onClose = onDone;
  el("visitorForm").addEventListener("submit", (event) => {
    event.preventDefault();
    submit();
  });
  el("visitorCancelButton").addEventListener("click", () => onClose(false));
  select("visitorBuilding").addEventListener("change", () =>
    changeBuilding(select("visitorBuilding").value)
  );
  input("visitorDate").addEventListener("change", fillTimes);
  input("visitorAllDay").addEventListener("change", updateAllDay);
  select("visitorStart").addEventListener("change", () => fillEndTimes());
  for (const id of ["visitorName", "visitorEmail"]) {
    input(id).addEventListener("input", () => showSuggestions(input(id).value));
    input(id).addEventListener("focus", () => showSuggestions(input(id).value));
    input(id).addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        show(el("visitorSuggestions"), false);
      }
    });
  }
  // Hide suggestions when focus moves elsewhere in the form, but not while clicking one.
  el("visitorForm").addEventListener("focusin", (event) => {
    const target = event.target as HTMLElement;
    if (
      !["visitorName", "visitorEmail"].includes(target.id) &&
      !target.closest("#visitorSuggestions")
    ) {
      show(el("visitorSuggestions"), false);
    }
  });
  el("visitorDoneButton").addEventListener("click", () => onClose(true));
  el("visitorAnotherButton").addEventListener("click", () => {
    resetForm();
    showStep("form");
  });
}

/** Opens the view at the form. Loads the org and settings the first time. */
export async function openVisitorBooking() {
  showStep("form");
  show(steps.form, false);
  setText(status, "Loading...", "detail");
  try {
    org = await loadOrganisation(api);
  } catch (error) {
    console.error(error);
    setText(status, `Couldn't load your organisation. ${describeError(error)}`, "detail error");
    return;
  }
  if (!org.buildings.length) {
    setText(status, "No buildings are set up in PlaceOS for visitors.", "detail error");
    return;
  }
  const buildings = select("visitorBuilding");
  if (!buildings.options.length) {
    org.buildings.forEach((building) => option(buildings, building.id, zoneLabel(building)));
  }
  show(el("visitorBuildingField"), org.buildings.length > 1);
  await changeBuilding(buildingId || (await org.defaultBuilding())?.id || org.buildings[0].id);
  if (!formReady) {
    resetForm();
  }
  setText(status, "");
  show(steps.form, true);
}

async function changeBuilding(id: string) {
  if (!org) {
    return;
  }
  buildingId = id;
  select("visitorBuilding").value = id;
  org.rememberBuilding(id);
  settings = await org.settings(id);
  const building = org.buildings.find((b) => b.id === id);
  setText(el("visitorSubtitle"), building ? zoneLabel(building) : "");
  show(el("visitorAllDayField"), visitorAllDayAllowed(settings));
  if (!visitorAllDayAllowed(settings)) {
    input("visitorAllDay").checked = false;
  }
  show(el("visitorPassField"), settings.get<boolean>("visitors.allow_pass_number") === true);
  show(
    el("visitorInternationalField"),
    settings.get<boolean>("visitors.allow_international") === true
  );
  const date = input("visitorDate");
  date.min = dateValue(Date.now());
  date.max = dateValue(Date.now() + maxFutureDays * 24 * 60 * minute);
  updateAllDay();
  // Duration limits are per building.
  if (select("visitorStart").value) {
    fillEndTimes();
  }
}

function resetForm() {
  formReady = true;
  input("visitorDate").value = dateValue(Date.now());
  input("visitorAllDay").checked =
    visitorAllDayAllowed(settings) &&
    visitorSetting<boolean>(settings, "all_day_default", false) === true;
  select("visitorStart").textContent = "";
  for (const id of ["visitorName", "visitorEmail", "visitorCompany", "visitorPass"]) {
    input(id).value = "";
  }
  // With app.visitors.reason_required the user must type a reason, so there's no default.
  input("visitorReason").value = reasonRequired() ? "" : "Visit";
  input("visitorInternational").checked = false;
  fillTimes();
  setText(el("visitorFormError"), "");
  show(el("visitorSuggestions"), false);
}

// ---------------------------------------------------------------------------------------------------------
// Date and time
// ---------------------------------------------------------------------------------------------------------

/** Start times for the chosen day; today starts from the next slot, other days at 9:00. */
function fillTimes() {
  const day = parseDate(input("visitorDate").value);
  const start = select("visitorStart");
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
  fillEndTimes(previous ? undefined : 60);
  updateAllDay();
}

const allDayChecked = () => visitorAllDayAllowed(settings) && input("visitorAllDay").checked;
const reasonRequired = () => settings.get<boolean>("visitors.reason_required") === true;

/**
 * The chosen start (ms) and duration (minutes): the all-day period, or the start and end pickers.
 * The all-day period is the whole day (or all_day_period hours), also today, as Workplace's
 * getAllDayTimeRange.
 */
function formTimes(): { start: number; duration: number } {
  if (allDayChecked()) {
    const day = parseDate(input("visitorDate").value);
    return isNaN(day)
      ? { start: 0, duration: 0 }
      : allDayPeriod(settings, day, visitorSetting(settings, "all_day_period", null));
  }
  return { start: +select("visitorStart").value, duration: +select("visitorEnd").value };
}

function updateAllDay() {
  const allDay = allDayChecked();
  show(el("visitorTimeFields"), !allDay);
  const { start, duration } = formTimes();
  setText(
    el("visitorAllDayTimes"),
    !allDay || !start
      ? ""
      : start + duration * minute <= Date.now()
        ? "The all-day period today has ended."
        : isFullDay(start, duration)
          ? ""
          : `${formatTime(start)} – ${formatTime(start + duration * minute)}`
  );
}

/** End times from the shortest to the longest visit (app.visitors.min_duration / max_duration). */
function fillEndTimes(keepDuration?: number) {
  const start = +select("visitorStart").value;
  const end = select("visitorEnd");
  const previous = keepDuration ?? (end.value ? +end.value : 60);
  end.textContent = "";
  if (!start) {
    return;
  }
  const { min, max } = visitorDurationLimits(settings);
  for (let minutes = min; minutes <= max; minutes += durationStep) {
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

// ---------------------------------------------------------------------------------------------------------
// Past visitors: suggestions from the user's "visitor-invitees" setting, as Workplace's form
// ---------------------------------------------------------------------------------------------------------

function showSuggestions(text: string) {
  const list = el("visitorSuggestions");
  list.textContent = "";
  if (!org) {
    show(list, false);
    return;
  }
  const query = text.trim().toLowerCase();
  const matches = pastVisitors(org)
    .reverse()
    .filter((visitor) =>
      [visitor.email, visitor.name, visitor.company].some((part) =>
        part.toLowerCase().includes(query)
      )
    )
    .slice(0, 8);
  for (const visitor of matches) {
    const item = document.createElement("li");
    const button = document.createElement("button");
    button.type = "button";
    button.className = "suggestion";
    button.textContent = [visitor.name, visitor.email, visitor.company]
      .filter((part, index, parts) => part && parts.indexOf(part) === index)
      .join(" · ");
    button.addEventListener("click", () => useVisitor(visitor));
    item.appendChild(button);
    list.appendChild(item);
  }
  show(list, matches.length > 0);
}

function useVisitor(visitor: PastVisitor) {
  input("visitorName").value = visitor.name;
  input("visitorEmail").value = visitor.email;
  input("visitorCompany").value = visitor.company;
  input("visitorInternational").checked = visitor.international;
  show(el("visitorSuggestions"), false);
  input("visitorReason").focus();
}

// ---------------------------------------------------------------------------------------------------------
// Send the invite
// ---------------------------------------------------------------------------------------------------------

function readForm(): VisitorRequest | string {
  const name = input("visitorName").value.trim();
  const email = input("visitorEmail").value.trim();
  if (!name) {
    return "Enter the visitor's name.";
  }
  if (!isEmail(email)) {
    return "Enter the visitor's email address.";
  }
  const day = parseDate(input("visitorDate").value);
  if (isNaN(day)) {
    return "Choose a date.";
  }
  if (day > Date.now() + maxFutureDays * 24 * 60 * minute) {
    return "Visitors can only be invited up to a year ahead.";
  }
  const allDay = allDayChecked();
  const { start, duration } = formTimes();
  if (!start || duration <= 0) {
    return "Choose a time.";
  }
  if (start + duration * minute <= Date.now()) {
    return allDay
      ? "The all-day period today has ended. Choose another day."
      : "Choose a time that hasn't passed.";
  }
  if (!allDay) {
    const hours = checkBookableHours(
      settings,
      start,
      duration,
      visitorSetting(settings, "bookable_hours", null),
      "Visitors"
    );
    if (hours) {
      return hours;
    }
  }
  const reason = input("visitorReason").value.trim();
  if (!reason && reasonRequired()) {
    return "Enter the reason for the visit.";
  }
  const showPass = el("visitorPassField").style.display !== "none";
  const showInternational = el("visitorInternationalField").style.display !== "none";
  return {
    start,
    duration,
    allDay,
    buildingId,
    name,
    email,
    company: input("visitorCompany").value.trim(),
    reason: reason || "Visit",
    passNumber: showPass ? input("visitorPass").value.trim() : "",
    international: showInternational && input("visitorInternational").checked,
  };
}

async function submit() {
  if (!org || submitting) {
    return;
  }
  const errorText = el("visitorFormError");
  const request = readForm();
  if (typeof request === "string") {
    setText(errorText, request, "error");
    return;
  }
  submitting = true;
  const button = el<HTMLButtonElement>("visitorSendButton");
  button.disabled = true;
  setText(errorText, "Sending invite...", "detail");
  try {
    const problem = await checkBeforeInvite(api, settings, org, request);
    if (problem) {
      throw new Error(problem);
    }
    await createVisitorBooking(api, buildVisitorBooking(request, org, settings));
    // Suggestions for next time; the invite is already sent, so a failure here doesn't matter.
    rememberVisitor(org, request).catch((error) => console.error(error));
    setText(errorText, "");
    showDone(request);
  } catch (error) {
    console.error(error);
    setText(errorText, describeError(error), "error");
  } finally {
    submitting = false;
    button.disabled = false;
  }
}

function showDone(request: VisitorRequest) {
  const summary = el("visitorDoneSummary");
  summary.textContent = "";
  const building = org?.buildings.find((b) => b.id === request.buildingId);
  line(summary, "✓ Visitor invited", "item-title ok");
  line(summary, [request.name, request.company].filter(Boolean).join(", "));
  line(summary, request.email);
  line(summary, whenText(request));
  line(summary, building ? zoneLabel(building) : "");
  formReady = false;
  showStep("done");
}
