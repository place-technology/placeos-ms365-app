# Visitor invites: follow-ups

The Invite a visitor flow (Today → Visitors card → **Invite a visitor**) ports PlaceOS Workplace's invite visitor form for one visitor (user-interfaces `libs/bookings` `invite-visitor-form.component.ts` and `BookingFormService.postForm`, develop as of 2026-10-06). How it works is summarised in `CLAUDE.md`. This file lists what is still open.

Status (2026-10-07): implemented and aligned with Workplace (user-interfaces develop `4214bc6ca`). It builds and type-checks, but nobody has tested it against a PlaceOS domain.

## Outstanding questions

1. **One visitor per host at a time.** Before posting, the add-in runs Workplace's check: the user's visitor bookings overlapping the visit are counted against `bookings.allowed_daily_visitor_count`, which defaults to **1**. So with default settings a host can't invite two visitors for overlapping times, in Workplace or the add-in. Check whether customers set this (0 or less turns it off), or whether the add-in should default to no limit. Bookings that are cancelled, declined, checked out or ended are not counted, as in Workplace. If the add-in can't read the bookings, the check passes and the Staff API decides, as in Workplace (`queryBookings` returns an empty list on errors).
2. **Invite email.** The add-in only creates the visitor booking. Whether the visitor gets an email, and what it says, depends on the domain's Staff API and notification setup, the same as for Workplace. Check one arrives on placeos-dev.
3. **Approval.** Bookings are sent with `approved: true` only when `bookings.no_approval` is set, as Workplace does. Check how unapproved invites look in Concierge.
4. **Pass number.** With `visitors.allow_pass_number`, the pass number is sent as `extension_data.pass_number`, as Workplace does. Concierge's guest list shows it from there. Check this on placeos-dev.
5. **Timezones.** Times are picked in the browser's time zone. `use_building_timezone` (`visitor.` → `visitors.` → `bookings.`, as Workplace's `BookingFormService`) only changes the `timezone` sent. Workplace also shows the times and the all-day period in the building's time zone. The add-in doesn't do this yet.
6. **All day today.** An all-day visit is the whole day (or the `all_day_period` hours), also when it is today, as Workplace's `getAllDayTimeRange`. So an all-day visit today starts at midnight, not now. Workplace also compares the all-day start with `bookable_hours`, so with bookable hours set it refuses all-day visits that start at midnight. The add-in doesn't apply bookable hours to all-day visits, which looks like Workplace's intent.
7. **Saved visitors.** Workplace saves the visitor to `visitor-invitees` before it posts the booking. The add-in saves the visitor only after the invite succeeds, so a failed invite doesn't add a suggestion.
8. **Visitor name.** Workplace marks the name as required but doesn't check it, and uses the email when the name is empty. The add-in requires a name.

Other Workplace behaviour that the add-in follows:

* With `visitors.reason_required`, the reason has no "Visit" default and the user must type one.
* Errors from the Staff API show its `error` or `message` text, as Workplace's `errorMessage`.

## Not ported yet

* Several visitors in one invite (`bookings.multiple_visitors`): Workplace makes a group container booking plus one booking per visitor. With that setting on, the add-in still invites one visitor at a time.
* Inviting on behalf of another host (`visitors.can_book_for_others` / `can_book_for_anyone`). The signed-in user is always the host.
* Calendar links on the done screen (`visitors.show_calendar_links`).
* Editing or cancelling a visitor invite from the Today view.
* Guest search (`/api/staff/v1/guests`). Suggestions come only from the user's own `visitor-invitees` setting.
