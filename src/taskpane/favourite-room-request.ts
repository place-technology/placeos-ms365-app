import type { AppSettings, Organisation } from "./booking-settings";
import type { FavouriteRoom } from "./favourite-rooms";
import type { PlaceosApi } from "./placeos-data";
import {
  checkBookableHours,
  getAvailableSpaceIds,
  getBookingRules,
  getSpaces,
  hiddenByRules,
  type BookingRequest,
  type DraftDetails,
} from "./room-booking-data";

export async function favouriteRoomRequest(
  api: PlaceosApi,
  org: Organisation,
  settings: AppSettings,
  favourite: FavouriteRoom,
  details: DraftDetails
): Promise<BookingRequest> {
  const building = org.buildings.find((entry) => entry.id === favourite.buildingId);
  const duration = (details.end - details.start) / 60000;
  const min = settings.get<number>("events.min_duration", 30) || 30;
  const max = Math.max(min, settings.get<number>("events.max_duration", 480) || 480);
  const maxDays = settings.get<number>("events.allowed_future_days", 180) || 180;
  if (!building || !org.favouriteSpaces.includes(favourite.id)) {
    throw new Error(
      "This favourite room is no longer available. Go back and refresh your favourites."
    );
  }
  if (details.recurring || details.recurrenceKnown === false) {
    throw new Error(
      "Quick add cannot check this meeting's recurrence. Use Rooms to review the booking."
    );
  }
  if (
    !Number.isFinite(details.start) ||
    !Number.isFinite(duration) ||
    details.start <= Date.now() ||
    duration < min ||
    duration > max ||
    details.start > Date.now() + maxDays * 86400000
  ) {
    throw new Error(
      "The meeting time is outside this room's booking limits. Update the event and refresh your favourites."
    );
  }
  const hoursError = checkBookableHours(
    settings,
    details.start,
    duration,
    undefined,
    "Rooms",
    building.timezone || ""
  );
  if (hoursError) throw new Error(hoursError);
  const [spaces, rules] = await Promise.all([
    getSpaces(api, building.id),
    getBookingRules(api, building.id),
  ]);
  const room = spaces.find((space) => space.id === favourite.id);
  if (!room || hiddenByRules(rules, room, details.start, duration, org.user.groups || [])) {
    throw new Error(
      "This room can't be booked for this meeting. Go back and choose another favourite."
    );
  }
  if (
    details.attendees.some((person) => person.email.toLowerCase() === room.email?.toLowerCase())
  ) {
    throw new Error("This room is already on the meeting.");
  }
  const available = await getAvailableSpaceIds(api, [room], details.start, duration, settings);
  if (!available.has(room.id)) {
    throw new Error(
      "This room is no longer available at the meeting time. Go back and refresh your favourites."
    );
  }
  return {
    title: details.title,
    start: details.start,
    duration,
    attendees: details.attendees,
    notes: "",
    recurrence: null,
    allDay: false,
    room,
    buildingId: building.id,
  };
}
