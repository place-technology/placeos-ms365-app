/* Renders the Today overview cards. Each card loads and fails independently. */

/* global document, console, window, HTMLElement, HTMLButtonElement */

import { FloorMap, loadMapSvg } from "./floor-map";
import type { PlaceosApi } from "./placeos-data";
import { PlaceosRequestError } from "./placeos-helper";
import {
  deskItems,
  getEvents,
  nextMeetingItem,
  parkingItems,
  roomItems,
  visitorItems,
  PlaceosLookups,
  type SettingsForZones,
  type TodayAction,
  type TodayItem,
} from "./today-data";

export type TodayLists = {
  rooms: HTMLElement;
  desks: HTMLElement;
  parking: HTMLElement;
  visitors: HTMLElement;
  // Only set in the mail context: the next meeting card's section and list.
  nextMeeting?: { section: HTMLElement; list: HTMLElement };
};

function friendlyError(error: unknown): string {
  if (error instanceof PlaceosRequestError) {
    if (error.status === 511) {
      return "PlaceOS can't read your calendar yet. Sign in to PlaceOS on the web once, then refresh.";
    }
    if (error.status === 501) {
      return "Calendar access isn't set up for this PlaceOS domain.";
    }
    if (error.status === 401 || error.status === 403) {
      return "PlaceOS didn't allow this request.";
    }
  }
  return "Couldn't load this right now.";
}

function renderMessage(list: HTMLElement, text: string, className: string) {
  list.textContent = "";
  const item = document.createElement("li");
  item.className = className;
  item.textContent = text;
  list.appendChild(item);
}

// For loading floor plans from item buttons.
let api: PlaceosApi;
let mapCount = 0;

/**
 * Opens or closes a floor plan under the item with the desk or parking space highlighted and zoomed in on
 * (Workplace's "view location").
 */
function toggleMap(
  item: HTMLElement,
  toggle: HTMLButtonElement,
  location: NonNullable<TodayItem["location"]>,
  name: string
) {
  const existing = item.querySelector(".item-map") as HTMLElement | null;
  if (existing) {
    item.removeChild(existing);
    toggle.setAttribute("aria-expanded", "false");
    return;
  }
  toggle.setAttribute("aria-expanded", "true");
  const panel = document.createElement("div");
  panel.className = "item-map map-view";
  const status = document.createElement("div");
  status.className = "item-detail";
  status.textContent = "Loading the floor plan...";
  const frame = document.createElement("div");
  frame.className = "floor-map compact";
  const canvas = document.createElement("div");
  canvas.className = "floor-map-canvas";
  // FloorMap scopes the plan's styles to the canvas id.
  canvas.id = `todayMap${++mapCount}`;
  const controls = document.createElement("div");
  controls.className = "floor-map-controls";
  frame.appendChild(canvas);
  frame.appendChild(controls);
  panel.appendChild(status);
  panel.appendChild(frame);
  const actions = item.querySelector(".item-actions");
  item.insertBefore(panel, actions || null);

  const map = new FloorMap(canvas);
  const control = (text: string, label: string, onClick: () => void) => {
    const element = button(text, true, onClick);
    element.setAttribute("aria-label", label);
    controls.appendChild(element);
  };
  control("+", "Zoom in", () => map.zoomBy(1.5));
  control("−", "Zoom out", () => map.zoomBy(1 / 1.5));
  control("Fit", "Show the whole floor", () => map.reset());

  loadMapSvg(location.mapUrl, api)
    .then((svg) => {
      if (!panel.parentNode) {
        return;
      }
      map.show(svg);
      const accent =
        window.getComputedStyle(document.documentElement).getPropertyValue("--accent").trim() ||
        "#0078d4";
      const missing = map.setRooms(
        [{ elementId: location.elementId, status: "free", title: name, selectable: false }],
        { free: accent, busy: accent, filtered: accent },
        accent
      );
      if (missing.length) {
        status.textContent = `${name} isn't on the ${location.level} floor plan.`;
        return;
      }
      map.select(location.elementId);
      map.focus(location.elementId);
      status.textContent = `${name} is highlighted on ${location.level}.`;
    })
    .catch((error) => {
      console.error(error);
      status.className = "item-detail error";
      status.textContent = "Couldn't load the floor plan.";
    });
}

function button(text: string, secondary: boolean, onClick: () => void): HTMLButtonElement {
  const element = document.createElement("button");
  element.type = "button";
  element.className = secondary ? "secondary" : "";
  element.textContent = text;
  element.addEventListener("click", onClick);
  return element;
}

/**
 * An item's buttons. An action with confirm first swaps the buttons for its question with Yes and No.
 * After an action succeeds the card reloads; if it fails, the error shows under the buttons.
 */
function renderActions(item: HTMLElement, actions: TodayAction[], reload: () => void) {
  const row = document.createElement("div");
  row.className = "item-actions";
  const message = document.createElement("div");
  message.className = "item-detail";
  item.appendChild(row);
  item.appendChild(message);

  const run = async (action: TodayAction) => {
    row.querySelectorAll("button").forEach((b) => ((b as HTMLButtonElement).disabled = true));
    message.className = "item-detail";
    message.textContent = `${action.label}...`;
    try {
      await action.run();
      reload();
    } catch (error) {
      console.error(error);
      showButtons();
      message.className = "item-detail error";
      message.textContent =
        error instanceof PlaceosRequestError && (error.status === 401 || error.status === 403)
          ? "PlaceOS didn't allow this."
          : `Couldn't ${action.label.toLowerCase()}. Try again.`;
    }
  };
  const ask = (action: TodayAction) => {
    row.textContent = "";
    message.textContent = "";
    const question = document.createElement("span");
    question.className = "item-detail";
    question.textContent = action.confirm || "";
    row.appendChild(question);
    row.appendChild(button("Yes", false, () => run(action)));
    row.appendChild(button("No", true, showButtons));
  };
  function showButtons() {
    row.textContent = "";
    for (const action of actions) {
      row.appendChild(
        button(action.label, !!action.secondary, () => (action.confirm ? ask(action) : run(action)))
      );
    }
  }
  showButtons();
}

function renderItems(list: HTMLElement, items: TodayItem[], emptyText: string, reload: () => void) {
  if (items.length === 0) {
    renderMessage(list, emptyText, "empty");
    return;
  }
  list.textContent = "";
  for (const entry of items) {
    const item = document.createElement("li");
    const title = document.createElement("div");
    title.className = "item-title";
    const location = entry.location;
    if (location) {
      // The name opens the floor plan.
      const toggle = button(entry.title, false, () =>
        toggleMap(item, toggle, location, entry.title)
      );
      toggle.className = "link";
      toggle.title = "Show on map";
      toggle.setAttribute("aria-expanded", "false");
      title.appendChild(toggle);
    } else {
      title.textContent = entry.title;
    }
    if (entry.badge) {
      const badge = document.createElement("span");
      badge.className = `badge ${entry.badge.kind}`;
      badge.textContent = entry.badge.text;
      title.appendChild(badge);
    }
    item.appendChild(title);
    for (const line of entry.details) {
      const detail = document.createElement("div");
      detail.className = "item-detail";
      detail.textContent = line;
      item.appendChild(detail);
    }
    if (entry.link) {
      const link = document.createElement("a");
      link.className = "item-link";
      link.href = entry.link.url;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.textContent = entry.link.text;
      item.appendChild(link);
    }
    if (entry.actions?.length) {
      renderActions(item, entry.actions, reload);
    }
    list.appendChild(item);
  }
}

async function loadCard(
  list: HTMLElement,
  emptyText: string,
  load: () => Promise<TodayItem[]>
): Promise<void> {
  renderMessage(list, "Loading...", "empty");
  try {
    renderItems(list, await load(), emptyText, () => loadCard(list, emptyText, load));
  } catch (error) {
    console.error(error);
    renderMessage(list, friendlyError(error), "empty error");
  }
}

/**
 * Loads all four cards. The calendar is fetched once and shared by Rooms and Visitors.
 */
export async function loadTodayOverview(
  placeosApi: PlaceosApi,
  lists: TodayLists,
  settingsFor?: SettingsForZones
): Promise<void> {
  api = placeosApi;
  const lookups = new PlaceosLookups(api);
  const events = getEvents(api);
  const nextMeeting = lists.nextMeeting;
  if (nextMeeting) {
    nextMeeting.section.style.display = "";
  }
  await Promise.all([
    nextMeeting
      ? loadCard(nextMeeting.list, "Nothing on your calendar in the next 7 days", async () => {
          const item = await nextMeetingItem(api, await events);
          return item ? [item] : [];
        })
      : Promise.resolve(),
    loadCard(lists.rooms, "No room bookings today", async () => roomItems(await events, lookups)),
    loadCard(lists.desks, "No desk booked today", () => deskItems(api, lookups, settingsFor)),
    loadCard(lists.parking, "No parking booked today", () =>
      parkingItems(api, lookups, settingsFor)
    ),
    loadCard(lists.visitors, "No visitors expected today", async () =>
      // If the calendar fails, still show visitor bookings.
      visitorItems(api, await events.catch(() => []))
    ),
  ]);
}
