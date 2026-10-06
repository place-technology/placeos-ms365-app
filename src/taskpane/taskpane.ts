/* Entry point for the Outlook add-in task pane. */

/* global Office */

import { AccountManager } from "./authConfig";
import { startApp } from "./app-ui";
import { outlookMeetingDraft } from "./outlook-draft";
import { applyTheme, themeFromBackground } from "./theme";

function applyOfficeTheme() {
  applyTheme(themeFromBackground(Office.context.officeTheme?.bodyBackgroundColor));
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    return;
  }
  applyOfficeTheme();
  try {
    // Not supported by every Outlook client; the initial theme still applies.
    Office.context.mailbox.addHandlerAsync(Office.EventType.OfficeThemeChanged, applyOfficeTheme);
  } catch {
    /* empty */
  }
  const accountManager = new AccountManager({
    isNestedAppAuthSupported: () =>
      Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"),
    // Older Outlook clients can't open the MSAL popup, so fall back to the Office dialog API.
    officeDialogFallback: true,
  });
  startApp(accountManager, "Outlook add-in", {
    // Mail context: reading or writing an email, or no item selected (pinned / no-item context).
    // Calendar events have their own context, so they don't show the next meeting.
    showNextMeeting: () => {
      const item = Office.context.mailbox.item;
      return !item || item.itemType === Office.MailboxEnums.ItemType.Message;
    },
    // Composing a meeting: Book a room adds the room to it.
    meetingDraft: outlookMeetingDraft,
  });
});
