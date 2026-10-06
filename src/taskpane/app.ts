/* Entry point for the personal tab app in the Teams, Outlook and Microsoft 365 app bar. */

/* global console, document */

import { app, nestedAppAuth } from "@microsoft/teams-js";
import { AccountManager } from "./authConfig";
import { startApp } from "./app-ui";
import { applyTheme } from "./theme";

const hostNames: Record<string, string> = {
  Teams: "Teams app",
  TeamsModern: "Teams app",
  Outlook: "Outlook app",
  OutlookWin32: "Outlook app",
  Office: "Microsoft 365 app",
};

async function start() {
  try {
    await app.initialize();
  } catch (error) {
    // Not running inside Teams, Outlook or Microsoft 365; leave the "open from" message showing.
    console.warn(`TeamsJS did not initialize: ${error}`);
    return;
  }
  document.body.classList.add("full-page");
  const context = await app.getContext();
  applyTheme(context.app.theme);
  app.registerOnThemeChangeHandler(applyTheme);
  const accountManager = new AccountManager({
    isNestedAppAuthSupported: () => nestedAppAuth.isNAAChannelRecommended(),
    // The Office dialog API isn't available outside Office add-ins.
    officeDialogFallback: false,
  });
  startApp(accountManager, hostNames[context.app.host.name] ?? `${context.app.host.name} app`, {
    // Full-page app with no item context: always show the next meeting.
    showNextMeeting: () => true,
  });
}

start();
