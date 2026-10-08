# Room booking: follow-ups

The Book a room flow (Today → Rooms card → **Book a room**) ports the legacy Angular add-in's room flow and meeting form (PlaceOS/user-interfaces `apps/outlook-addin`, `libs/events` `EventFormService`). How it works is summarised in `CLAUDE.md`. This file lists what is still open.

Status (2026-10-01): implemented. It builds, type-checks and passes lint, but hasn't been tested end to end in Outlook against a PlaceOS domain.

## Outstanding questions

1. **Settings metadata name.** The legacy add-in reads its settings from zone metadata named after its URL path (for example `/outlook/` gives `outlook_app`). This add-in uses fixed names instead: `outlook_app`, then `outlook-addin_app`, then `workplace_app`, taking the first one found on each zone (`booking-settings.ts`). `outlook_app` is the documented name (`entra-setup.md`, "Room booking settings"). Check which names existing customers have populated, and drop `outlook-addin_app` if nobody uses it.
2. **Monthly recurrence week.** Staff API ignores `nth_of_month`; PlaceOS/calendar works out the week from the start date as `day // 7`, which books monthly series starting on the 7th, 14th, 21st or 28th a week late. The form blocks monthly repeats starting on those days (`monthlyStartUnsupported`); on every other day the form's label (`weekOfMonth`) matches what Office 365 books. Google-backed domains are still wrong for the 22nd to 27th (sent as second-to-last). A fix is parked on PlaceOS/calendar branch `fix/monthly-recurrence-week` (pushed, no PR) until the calendar owners confirm nothing relies on the current behaviour. Remove the guard once Staff API ships with a fixed calendar version.
3. **Linked booking bodies.** The bodies for visitor, `catering-order` and `asset-request` bookings follow Workplace's `createBookingsForEvent` and `validateAssetRequestsForResource` but with fewer fields. Visitor and catering bookings are made for the event host. Asset requests are made for the user, as in Workplace. Check they show up in the catering and asset management apps, and at reception for visitors.
4. **Catering order status.** Orders are sent with `status: "accepted"`, which is the legacy class default. Confirm that's right for new orders.
5. **Timezones.** Bookable hours use the building's timezone, as Workplace's `postForm` does (local time if the building has no timezone, or if the webview can't convert timezones). The end times offered stop at the end of the bookable hours. Booking rules (`is_between`) and catering/asset `after_hour`/`before_hour` rules use the browser's local time, as in Workplace. With `app.use_region`, a room in another building is checked against the selected building's hours. Workplace checks each room's own building.
6. **Favourites saving.** Starring a menu item or asset saves `favourite_menu_items` / `favourite_assets` in the user's `settings` metadata with a PUT of the whole details, as Workplace does. The add-in re-reads the settings first to avoid overwriting newer changes, but Workplace open in another tab can still overwrite the add-in's change (and the other way round). Check the user's PlaceOS role is allowed to write their own metadata.
7. **Caterer filter.** Workplace limits one order to a single caterer and defaults its caterer picker to the first one. The add-in already makes one order per caterer, so its picker defaults to "All caterers". `app.catering_provider` restricts the menu to that caterer, as in Workplace. Searches match descriptions as well as names (Workplace's catering search only matches names).
8. **Calendar access (511).** Booking needs PlaceOS to have a delegated calendar token for the user, the same as the Today view. Confirm what the user experience should be when it's missing.
9. **All day and bookable hours.** As in Workplace, bookable hours also apply to an all-day booking. If a building has `events.bookable_hours`, a whole-day booking fails the check unless `events.all_day_period` is inside those hours. Today's all-day period starts at its normal time, not from now (Workplace's `getAllDayTimeRange`). Check that customers who use both settings also set `events.all_day_period`.
10. **Settings that changed to match Workplace.** Repeats are only offered with `events.allow_recurrence: true` (it was on unless set to `false`). All-day meetings can't repeat daily unless `events.allow_daily_allday_recurrence` is set. End times use `events.duration_step` (default 30 minutes) and `events.custom_duration_options`, and the default length is `events.default_duration` (60). Assets are offered with `events.allow_assets` or Workplace's `events.has_assets`. Catering and asset delivery times use `catering.step_interval` / `assets.step_interval` (default 5 minutes), up to the end of the meeting less `end_offset`. Rooms in the building's `catering-settings` `disabled_rooms` get no catering, and menu items are hidden for rooms in their `hide_for_zones`. A room capacity of 0 counts as unknown, so the capacity filter doesn't hide it.

## Adding a room to an Outlook meeting

Status (2026-10-06): implemented. It builds, type-checks and passes lint, but hasn't been tested in Outlook.

When the task pane is open on a meeting the user is organising (calendar compose), the Rooms card button reads **Add a room to this meeting**. The flow then fills in that draft instead of having PlaceOS create an event (`outlook-draft.ts`, `MeetingDraft` in `room-booking-data.ts`). The form is prefilled from the draft's subject, start, end and attendees. Choosing a room writes back the start, end, subject, any added attendees, and the room. On Mailbox 1.8 clients the room is added as an `enhancedLocation` of type Room, which also adds its mailbox as a resource. Older clients add the room's mailbox as a required attendee and set the location text. Nothing is booked until the user sends the invitation, and then the room mailbox accepts or declines it. This needed `ReadWriteItem` in `manifest.xml` (version 1.0.1.0), so customers have to redeploy the manifest.

In this mode:

* Notes, repeat and all day are hidden, because the draft keeps its own body, recurrence and all-day setting. If the draft repeats (known on Mailbox 1.7), a note says that availability is only checked for the first meeting.
* Catering, assets and visitors are hidden. They are PlaceOS bookings linked by `event_id`/`ical_uid`, and the event doesn't exist until the draft is sent and PlaceOS has synced it.
* Attendees are only added. Removing a prefilled attendee in the form doesn't remove them from the draft.

Open questions:

1. **Editing a meeting that already has the room.** The availability check counts the meeting's own booking as busy, so the room isn't offered again.
2. **Room mailbox policy.** Whether the room accepts depends on its Exchange booking settings (for example conflicts in a series, booking window, delegates), not on PlaceOS booking rules. The add-in still applies booking rules and bookable hours, but only to filter the list.
3. **enhancedLocation on every client.** Check that the Room location adds the room as a resource on new Outlook for Windows, classic Outlook, Outlook on the web and Mac.
4. **Catering and assets.** These would need the event first: either save the draft (`saveAsync`) and wait for PlaceOS to see it, or let the user add them from the Today view once the meeting is sent.
5. **Opening straight into the flow.** The pane still opens on Today. It could open Add a room directly in this context.

## Map view

Status (2026-10-06): implemented. It builds, type-checks and passes lint, and the map component was checked in headless Chromium with a sample SVG (sanitising, style scoping, tap to select, pan and zoom). Real S3 floor plans load and work in local development (2026-10-06, dev server against placeos-dev, through the S3 map proxy). Not yet checked in production hosting, where maps come straight from S3.

The rooms step has a **List / Map** toggle when a searched building has levels with a floor plan (`map_id` on the level zone), and then opens on the map. Once the user picks List, later searches stay on the list. The map shows one level at a time, picked from a list that shows how many free rooms match the filters on each level. It opens on the Level filter's level if that has a map, else the level with the most free rooms. Free rooms that match the filters are green and can be tapped (or focused and Enter pressed) to show their details and **Choose this room**. Busy rooms are red, and free rooms hidden by the filters are grey. Rooms hidden by booking rules aren't coloured. Colours can be overridden with `explore.colors` (`space-free`/`free`, `space-busy`/`busy`, `not-bookable`), as in Workplace. The legacy add-in's `#zones` layer is hidden. Free rooms whose element isn't on the floor plan are counted under the map, and are still in the list.

Floor plans are SVGs on S3, linked from the level zone's `map_id`. They're fetched without credentials, so the bucket's CORS rules must allow the customer domain (Workplace on that domain needs the same). The bucket doesn't allow `localhost`, so in local development the dev server proxies S3 maps (`/__map-proxy/<s3 host>/…` in `webpack.config.js`).

Differences from the legacy add-in: there's no "All levels" option (it stacked a small map per level, which doesn't fit a task pane), rooms aren't marked with pins, and MapsIndoors and Cisco maps aren't supported.

Open questions:

1. **Large floor plans.** The SVG is inlined in the page rather than drawn to a canvas as Workplace does. Check performance with large plans in the Outlook webviews, especially IE-based classic Outlook.
2. **Dark theme.** The map keeps a white background, because floor plans are drawn for one.

## Not ported yet

* Booking for someone else (`events.can_book_for_others`, host picker).
* More than one room per meeting (`events.multiple_spaces`), including picking which room a catering order goes to.
* Attendee availability finder (legacy "Availability" link).
* CSV attendee import / template download.
* Guest search (`/api/staff/v1/guests`) and suggestions from past visitors (`visitor-invitees` user setting). Workplace also saves a meeting's external attendees to `visitor-invitees` when it books.
* Start times limited to the bookable hours (Workplace's time field). The add-in offers every time and shows an error.
* Catering and asset delivery on a specific day for multi-day meetings (`deliver_day_offset`), and the "deliver at exact time" toggle (`deliver_time`).
* Booking rule `auto_approve` and room approval display.
* Room alerts (`room_alerts` org metadata).
* Saving favourite rooms (favourites are read only).
* Upcoming bookings list, and editing or cancelling existing bookings.
* Catering, assets and visitor registration when adding a room to an Outlook meeting (see below).
