/*
 * The meeting being composed in Outlook, for Book a room's "add to this meeting" mode. Reads the draft's
 * details to prefill the form and writes the chosen time, attendees and room back to it. Nothing is booked
 * until the user sends the invitation: Exchange's room mailbox then accepts or declines it.
 * Writing needs ReadWriteItem in manifest.xml.
 */

import {
  spaceName,
  type BookingRequest,
  type DraftDetails,
  type MeetingDraft,
  type Person,
} from "./room-booking-data";

/** Wraps an Office.js callback API as a promise. */
function officeAsync<T>(call: (callback: (result: Office.AsyncResult<T>) => void) => void) {
  return new Promise<T>((resolve, reject) =>
    call((result) =>
      result.status === Office.AsyncResultStatus.Succeeded
        ? resolve(result.value)
        : reject(new Error(result.error?.message || "Outlook couldn't update the meeting."))
    )
  );
}

const toPeople = (recipients: Office.EmailAddressDetails[]): Person[] =>
  recipients
    .filter((r) => r.emailAddress)
    .map((r) => ({ name: r.displayName || r.emailAddress, email: r.emailAddress }));

const supports = (version: string) =>
  Office.context.requirements.isSetSupported("Mailbox", version);

/** The current item as a meeting draft, or null if it isn't a meeting the user is organising. */
export function outlookMeetingDraft(): MeetingDraft | null {
  const item = Office.context.mailbox?.item;
  // In read mode subject is a string; in compose it's an object with getAsync/setAsync.
  if (
    !item ||
    item.itemType !== Office.MailboxEnums.ItemType.Appointment ||
    typeof (item.subject as unknown as Office.Subject)?.setAsync !== "function"
  ) {
    return null;
  }
  const draft = item as Office.AppointmentCompose;

  const attendees = async () => {
    const [required, optional] = await Promise.all([
      officeAsync<Office.EmailAddressDetails[]>((cb) => draft.requiredAttendees.getAsync(cb)),
      officeAsync<Office.EmailAddressDetails[]>((cb) => draft.optionalAttendees.getAsync(cb)),
    ]);
    return toPeople([...required, ...optional]);
  };

  return {
    async read(): Promise<DraftDetails> {
      const [title, start, end, people, recurrence] = await Promise.all([
        officeAsync<string>((cb) => draft.subject.getAsync(cb)),
        officeAsync<Date>((cb) => draft.start.getAsync(cb)),
        officeAsync<Date>((cb) => draft.end.getAsync(cb)),
        attendees(),
        supports("1.7")
          ? officeAsync<Office.Recurrence | null>((cb) => draft.recurrence.getAsync(cb)).catch(
              () => null
            )
          : Promise.resolve(null),
      ]);
      return {
        title: title || "",
        start: start.getTime(),
        end: end.getTime(),
        attendees: people,
        recurring: !!recurrence?.recurrenceType,
      };
    },

    async fill(request: BookingRequest) {
      const { room } = request;
      // Setting the start keeps the duration, so set the end afterwards.
      await officeAsync<void>((cb) => draft.start.setAsync(new Date(request.start), cb));
      await officeAsync<void>((cb) =>
        draft.end.setAsync(new Date(request.start + request.duration * 60000), cb)
      );
      const subject = await officeAsync<string>((cb) => draft.subject.getAsync(cb));
      if (request.title && request.title !== subject) {
        await officeAsync<void>((cb) => draft.subject.setAsync(request.title, cb));
      }
      // Only adds people; anyone already on the invitation stays on it.
      const known = new Set((await attendees()).map((p) => p.email.toLowerCase()));
      const added = request.attendees.filter((p) => !known.has(p.email.toLowerCase()));
      if (added.length) {
        await officeAsync<void>((cb) =>
          draft.requiredAttendees.addAsync(
            added.map((p) => ({ displayName: p.name || p.email, emailAddress: p.email })),
            cb
          )
        );
      }
      if (!room.email) {
        throw new Error("This room has no mailbox, so it can't be added to an Outlook meeting.");
      }
      if (supports("1.8")) {
        // A Room location also adds the room's mailbox to the invitation as a resource.
        await officeAsync<void>((cb) =>
          draft.enhancedLocation.addAsync(
            [{ id: room.email!, type: Office.MailboxEnums.LocationType.Room }],
            cb
          )
        );
      } else {
        // Older clients can't add resources: invite the room's mailbox and set the location text.
        if (!known.has(room.email.toLowerCase())) {
          await officeAsync<void>((cb) =>
            draft.requiredAttendees.addAsync(
              [{ displayName: spaceName(room), emailAddress: room.email! }],
              cb
            )
          );
        }
        await officeAsync<void>((cb) => draft.location.setAsync(spaceName(room), cb));
      }
    },
  };
}
