# Shared presentation

The Outlook task pane and Teams / Outlook app-bar app use the same HTML template and stylesheet. The `full-page` host class enables wider layouts; narrow embedded views retain a single column of results.

## Launcher

The `.launcher` navigation contains repeatable `.launcher-tile` buttons with a decorative SVG, `.launcher-label`, and `.launcher-detail`. Existing booking button IDs and click handlers are retained. The calendar draft context updates `bookRoomActionLabel`, preserving the icon and Rooms label.

Add future services, such as Catering, as another tile wired to its own flow. The grid wraps automatically; no fixed feature count is assumed. Catering remains in the existing room confirmation flow until a dedicated section is implemented. Today's bookings, check-in, cancellation, floor-plan links, and Refresh remain below the launcher.

## Favourite room shortcuts

The duplicate in-app PlaceOS title is removed; a compact navigation link retains Diagnostics and Back to today. Outlook's own add-in title bar is controlled by the host.

Your Favourites appears between the launcher and Your Bookings. It reads the existing user's `favourite_spaces` settings and existing room availability APIs, respects hidden-room rules, building hours, and duration limits, and checks the current editable Outlook meeting. Green means available; red means unavailable or outside booking limits. Unknown availability, unsupported recurrence checks, recurring meetings, and non-calendar contexts never show an Add action. Rooms already invited to the meeting are labelled accordingly.

The shortcut rereads the draft and opens confirmation for the selected favourite directly, without showing the search form or requiring a title on an unfinished Outlook draft. It validates the selected room with the existing room, rule, and availability APIs, retaining the exact event times. It does not create an event or send an invitation. Confirmation rechecks availability and rejects changed meeting times or a switched Outlook item; current title and attendees are preserved. Sending the Outlook invitation remains the step that books the room. Time, recurrence, recipient, and pinned-item changes refresh the section on supported hosts; manual Refresh is also available.

Verified in the live Outlook calendar preview with saved Sydney rooms and the event's actual time. Isolated regression checks cover available/unavailable results, unknown availability, missing calendar context, past times, recurrence, booking limits, already-invited rooms, empty favourites, hidden rooms, and out-of-order responses. No live meeting was submitted during verification.

## Result styling

`result-card.ts` renders the shared room, desk, and parking result button using text nodes. Names, location, capacity where supplied, and features remain visible. The availability label describes the existing filtered available candidates; unavailable resources remain excluded from lists and retain the existing map treatment. Choosing a result invokes the original handler and opens confirmation.

Light surfaces use the supplied Nestlé palette. Host dark and high-contrast themes retain distinct tokens. Hover uses a warm highlight; keyboard focus uses an offset outline; selected resource summaries have a stronger border and explicit label. Loading indicators respect reduced motion. Empty and error messages remain textual, with distinct boundaries. Native disabled controls retain their disabled behavior.

## Verification

- Development and production builds, TypeScript checking, and repository lint.
- Isolated Chromium rendering with the actual template, stylesheet, and shared result component, using sample data and blocked network requests.
- Widths: 280, 320, 360, 480, and 1100px; long labels and emails; a fifth launcher tile; room and desk results; narrow forms; selected summaries; loading, empty, and error states.
- Result hover, visible keyboard focus, Enter and Space activation, disabled styling, dark theme, contrast theme, forced colors, and reduced motion.

The isolated review does not authenticate to Microsoft or PlaceOS, submit bookings, or reproduce Office / Teams host SDKs. Sideloaded Outlook mail/calendar and Teams checks remain necessary for host authentication, draft synchronization, live maps, and end-to-end booking behavior. No service, authentication, permission, or booking data logic was changed. The repository has no automated test suite.
