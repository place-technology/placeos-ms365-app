# Visitor invites: follow-ups

The Invite a visitor flow (Today → Visitors card → **Invite a visitor**) ports PlaceOS Workplace's invite visitor form for one visitor (user-interfaces `libs/bookings` `invite-visitor-form.component.ts` and `BookingFormService.postForm`, develop as of 2026-10-06). How it works is summarised in `CLAUDE.md`. This file lists what is still open.

Status (2026-10-06): implemented. It builds, type-checks and passes lint, but hasn't been tested against a PlaceOS domain.

## Outstanding questions

1. **One visitor per host at a time.** Before posting, the add-in runs Workplace's check: the user's visitor bookings overlapping the visit are counted against `bookings.allowed_daily_visitor_count`, which defaults to **1**. So with default settings a host can't invite two visitors for overlapping times, in Workplace or the add-in. Check whether customers set this (0 or less turns it off), or whether the add-in should default to no limit.
2. **Invite email.** The add-in only creates the visitor booking. Whether the visitor gets an email, and what it says, depends on the domain's Staff API and notification setup, the same as for Workplace. Check one arrives on placeos-dev.
3. **Approval.** Bookings are sent with `approved: true` only when `bookings.no_approval` is set, as Workplace does. Check how unapproved invites look in Concierge.
4. **Pass number.** Workplace shows the pass number field (`visitors.allow_pass_number`) but doesn't save it. The add-in saves it as `extension_data.pass_number`, which is where Workplace reads it from when editing. Check Concierge shows it.
5. **Timezones.** Times are picked in the browser's time zone. `bookings.use_building_timezone` / `visitors.use_building_timezone` only change the `timezone` sent, as with rooms and desks.

## Not ported yet

* Several visitors in one invite (`bookings.multiple_visitors`): Workplace makes a group container booking plus one booking per visitor. With that setting on, the add-in still invites one visitor at a time.
* Inviting on behalf of another host (`visitors.can_book_for_others` / `can_book_for_anyone`). The signed-in user is always the host.
* Calendar links on the done screen (`visitors.show_calendar_links`).
* Editing or cancelling a visitor invite from the Today view.
* Guest search (`/api/staff/v1/guests`). Suggestions come only from the user's own `visitor-invitees` setting.
