# Desk and parking booking: follow-ups

The Book a desk flow (Today → Desks card → **Book a desk**) ports PlaceOS Workplace's desk booking with `app.desks.use_assets` (user-interfaces `libs/assets` `desk-assets.fn.ts`, `libs/bookings` `BookingFormService`, develop as of 2026-10-06). **Book parking** (Today → Parking card) uses the same view for parking spaces (see [Parking](#parking) below). How both work is summarised in `CLAUDE.md`. This file lists what is still open.

Status (2026-10-06): implemented. It builds, type-checks and passes lint, but hasn't been tested against a PlaceOS domain. No domain has desk assets yet (desks are still in zone metadata everywhere), so until one is migrated the flow shows "No desks are set up in PlaceOS for this building".

## Assets only

Desks are read only as Assets: asset type `_DESKS_` in the hidden asset category `_DESKS_`, with `zone_id` = the desk's level. The legacy `desks` zone metadata isn't read, on purpose, because Assets are the way forward. Workplace creates the category and type when they're missing. The add-in only reads them and never creates them. The Today view still shows names for bookings of either kind.

## Outstanding questions

1. **Testing data.** Migrate one test building on placeos-dev to desk assets (Backoffice or Concierge with `desks.use_assets`), with a floor plan on a level, to test the list, the map and booking.
2. **Approval.** Bookings are sent with `approved: true` only when `bookings.no_approval` is set, as Workplace does. Otherwise they're created unapproved. Check how that looks in Concierge and on the Today view.
3. **Rules host.** Booking rules use the signed-in user's groups. The add-in can't book for someone else, so Workplace's `force_current_user_for_booking_rules` doesn't apply.
4. **Assigned desks.** With `assigned_resource_booking` other than `allow`, a user with an assigned desk (`assigned_to` on a desk asset in the searched buildings) sees that desk instead of the form, with **View on map** (its level's floor plan, the desk highlighted and zoomed in on). The check covers the selected building, or its region with `use_region`, when the view opens. Workplace checks every building the user can book in, so a desk assigned in another building stops the user booking in Workplace but not in the add-in. With `allow`, the assigned desk isn't shown at all.
5. **Timezones.** Times are picked in the browser's time zone. `use_building_timezone` only changes the `timezone` sent, as with rooms.
6. **Questionnaire.** Workplace can ask health questions first (`desks.ignore_questions === false`). It isn't ported.

## Check in, check out and cancel

Desk bookings on the Today view have **Check in**, **Check out** and **Cancel** buttons (see `CLAUDE.md` for the rules). They work for any desk booking, whether the desk is an asset or in metadata. Open questions:

1. **Check-in window.** Workplace offers check-in from 15 minutes before the start, and the add-in does the same. Staff API may apply its own window, so an early check-in could be refused. If it is, the item shows "Couldn't check in".
2. **Series.** Cancel only removes that day's occurrence of a recurring booking. Workplace can also cancel the whole series (`allow_series_delete`), but the add-in can't.
3. **Show on map.** Tapping the desk name opens its level's floor plan under the item, with the desk highlighted and zoomed in. This works for metadata desks too, using `extension_data.map_id`, else the `asset_id`. Bookings made before `map_id` was saved, where the desk id isn't the map element id, show "isn't on the floor plan".
4. **Parking.** Parking bookings have the same buttons, using the `parking.` settings. Check-in is hidden for `unallocated…` parking requests, as Workplace does.

## Not ported yet

* Recurring desk bookings (`desks.allow_recurrence`, `clashing-assets`).
* Booking for someone else and group desk bookings.
* Auto-allocation (`desks.auto_allocation`), the nearby-desk preselect, and favourite desks.
* Lockers (`desks.can_book_lockers`), assets with a desk (`desks.allow_assets`) and payments.
* Custom duration options (`custom_duration_options`).

## Parking

Status (2026-10-06): implemented from Workplace's parking flow (`apps/workplace` `parking-flow`, `libs/bookings` `ParkingService`, `libs/assets` `parking-assets.fn.ts`). It builds, type-checks and passes lint, but hasn't been tested against a PlaceOS domain. Until a building has parking levels with `_PARKING_SPACES_` assets, it shows "No parking spaces are set up in PlaceOS for this building".

Outstanding questions:

1. **Testing data.** A building with a level tagged `parking` (with a floor plan), some spaces in Concierge, and a parking user entry for a test account (to try `deny` and the plate number).
2. **Title.** Workplace's parking form asks for a title. The add-in sends "Parking Booking", as it sends "Desk Booking" for desks.
3. **Assigned space.** Handled like an assigned desk: with `assigned_resource_booking` other than `allow`, the user sees their space and can't book another. Workplace's parking page only shows the assigned space instead of the form when the user also has a parking booking today, but its submit check blocks the booking either way.
4. **Home location.** With `parking.restrict_home_location`, only the selected building's desk assets are checked for a desk assigned to the user. Workplace looks for the first building with one, so someone with assigned desks in two buildings is treated slightly differently.
5. **Plate number.** It isn't saved to the user's settings. Workplace reads the `plate_number` user setting but doesn't write it either.
6. **`allow_all_day`.** Workplace reads `parking.allow_all_day || bookings.allow_all_day`, defaulting to on, so a `false` on parking is ignored unless bookings is also `false`. The add-in uses `parking.allow_all_day ?? bookings.allow_all_day ?? true`, as for desks.

Not ported: parking requests (`parking-request-flow`, unallocated spaces and the waitlist), booking for someone else, fleet vehicles, `space_restrictions`, favourite spaces, and the legacy `parking-spaces` zone metadata.
