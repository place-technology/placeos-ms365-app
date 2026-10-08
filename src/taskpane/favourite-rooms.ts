/* global document, HTMLElement, HTMLButtonElement */

import { loadOrganisation, zoneLabel } from "./booking-settings";
import type { PlaceosApi } from "./placeos-data";
import {
  checkBookableHours,
  getAvailableSpaceIds,
  getBookingRules,
  getSpaces,
  hiddenByRules,
  spaceName,
  type MeetingDraft,
} from "./room-booking-data";

export type FavouriteRoom = { id: string; buildingId: string };

export function favouriteRoomsView(
  api: PlaceosApi,
  getDraft: () => MeetingDraft | null,
  choose: (room: FavouriteRoom) => void
) {
  const status = document.getElementById("favouriteRoomsStatus") as HTMLElement;
  const list = document.getElementById("favouriteRoomsList") as HTMLElement;
  const refreshButton = document.getElementById("refreshFavourites") as HTMLButtonElement;
  let generation = 0;

  async function refresh() {
    const current = ++generation;
    list.textContent = "";
    status.className = "detail loading";
    status.textContent = "Checking your favourite rooms...";
    refreshButton.disabled = true;
    try {
      const org = await loadOrganisation(api);
      const favourites = new Set(org.favouriteSpaces);
      const draft = getDraft();
      const details = draft ? await draft.read() : null;
      const duration = details ? (details.end - details.start) / 60000 : 0;
      const validTime =
        !!details &&
        Number.isFinite(details.start) &&
        Number.isFinite(duration) &&
        details.start > Date.now() &&
        duration > 0;
      const rooms = favourites.size
        ? (
            await Promise.all(
              org.buildings.map(async (building) => {
                const spaces = (await getSpaces(api, building.id)).filter((space) =>
                  favourites.has(space.id)
                );
                if (!spaces.length) return [];
                const settings = await org.settings(building.id);
                const rules = details && validTime ? await getBookingRules(api, building.id) : [];
                const allowed = spaces.filter(
                  (space) =>
                    !details ||
                    !validTime ||
                    !hiddenByRules(rules, space, details.start, duration, org.user.groups || [])
                );
                const min = settings.get<number>("events.min_duration", 30) || 30;
                const max = Math.max(min, settings.get<number>("events.max_duration", 480) || 480);
                const hoursError =
                  details && validTime
                    ? checkBookableHours(
                        settings,
                        details.start,
                        duration,
                        undefined,
                        "Rooms",
                        building.timezone || ""
                      )
                    : "";
                const maxDays = settings.get<number>("events.allowed_future_days", 180) || 180;
                const restricted =
                  !!hoursError ||
                  duration < min ||
                  duration > max ||
                  (!!details && details.start > Date.now() + maxDays * 86400000);
                let available: Set<string> | null = null;
                if (
                  details &&
                  validTime &&
                  !restricted &&
                  !details.recurring &&
                  details.recurrenceKnown !== false &&
                  allowed.length
                ) {
                  try {
                    available = await getAvailableSpaceIds(
                      api,
                      allowed,
                      details.start,
                      duration,
                      settings
                    );
                  } catch {
                    available = null;
                  }
                }
                return allowed.map((room) => {
                  const added = details?.attendees.some(
                    (person) => person.email.toLowerCase() === room.email?.toLowerCase()
                  );
                  const state = added
                    ? "Already on this meeting"
                    : !details
                      ? "Select a calendar event"
                      : !validTime
                        ? "Choose a future meeting time"
                        : details.recurrenceKnown === false
                          ? "Meeting recurrence could not be checked"
                          : details.recurring
                            ? "Series availability not checked"
                            : restricted
                              ? "Outside booking limits"
                              : available === null
                                ? "Availability unknown"
                                : available.has(room.id)
                                  ? "Available"
                                  : "Unavailable";
                  return { room, building, state };
                });
              })
            )
          ).flat()
        : [];
      if (current !== generation) return;
      if (draft?.isCurrent?.() === false) {
        await refresh();
        return;
      }
      status.className = "detail";
      status.textContent =
        details && validTime
          ? `${new Date(details.start).toLocaleString([], { weekday: "short", day: "numeric", month: "short", hour: "numeric", minute: "2-digit" })} – ${new Date(details.end).toLocaleString([], { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" })}. Availability is checked again before adding.`
          : "Open an editable Outlook calendar event to check availability and add a room.";
      if (!rooms.length) {
        const empty = document.createElement("li");
        empty.className = "empty";
        empty.textContent = favourites.size
          ? "No favourite rooms are available to show for this meeting."
          : "No favourite rooms saved in PlaceOS yet.";
        list.appendChild(empty);
      }
      const seen = new Set<string>();
      for (const { room, building, state } of rooms) {
        if (seen.has(room.id)) continue;
        seen.add(room.id);
        const item = document.createElement("li");
        const title = document.createElement("div");
        title.className = "item-title";
        title.textContent = spaceName(room);
        const location = document.createElement("div");
        location.className = "item-detail";
        const level = org.levelFor(room.zones);
        location.textContent = [level ? zoneLabel(level) : "", zoneLabel(building)]
          .filter(Boolean)
          .join(", ");
        const footer = document.createElement("div");
        footer.className = "result-footer";
        const indicator = document.createElement("span");
        indicator.className = `result-availability ${state === "Available" ? "ok" : state === "Unavailable" || state === "Outside booking limits" ? "error" : "unknown"}`;
        indicator.textContent = state;
        footer.appendChild(indicator);
        if (state === "Available") {
          const button = document.createElement("button");
          button.type = "button";
          button.className = "link";
          button.textContent = "Add room to this booking";
          button.addEventListener("click", () => {
            if (current === generation) choose({ id: room.id, buildingId: building.id });
          });
          footer.appendChild(button);
        }
        item.appendChild(title);
        item.appendChild(location);
        item.appendChild(footer);
        list.appendChild(item);
      }
    } catch {
      if (current !== generation) return;
      status.className = "detail error";
      status.textContent = "Couldn't check your favourite rooms. Try Refresh.";
    } finally {
      if (current === generation) refreshButton.disabled = false;
    }
  }

  refreshButton.addEventListener("click", refresh);
  return { refresh };
}
