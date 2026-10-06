// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

/* global console, document*/

import {
  BrowserAuthError,
  createNestablePublicClientApplication,
  type IPublicClientApplication,
} from "@azure/msal-browser";
import { createMsalConfig } from "./msalconfig";
import type { AddinConfig } from "./addin-config";
import { createLocalUrl } from "./util";
import { getTokenRequest } from "./msalcommon";

export type AuthDialogResult = {
  accessToken?: string;
  username?: string;
  error?: string;
};

// Which path produced the token. Only "silent" counts as passing the POC's silent criterion.
export type AuthPath = "silent" | "popup" | "dialog";

export type TokenResult = {
  accessToken: string;
  username: string;
  authPath: AuthPath;
};

type DialogEventMessage = { message: string; origin: string | undefined };
type DialogEventError = { error: number };
type DialogEventArg = DialogEventMessage | DialogEventError;

// What the hosting app (Office add-in or TeamsJS personal tab) supports.
export type AuthHost = {
  isNestedAppAuthSupported: () => boolean;
  // Office add-ins only: fall back to the Office dialog API when the MSAL popup can't open.
  officeDialogFallback: boolean;
};

// Encapsulate functions for getting user account and token information.
export class AccountManager {
  constructor(private readonly host: AuthHost) {}

  private pca: IPublicClientApplication | undefined = undefined;
  private _dialogApiResult: Promise<TokenResult> | null = null;
  private _usingFallbackDialog = false;

  private getSignOutButton() {
    return document.getElementById("signOutButton");
  }

  private setSignOutButtonVisibility(isVisible: boolean) {
    const signOutButton = this.getSignOutButton();
    if (signOutButton) {
      signOutButton.style.visibility = isVisible ? "visible" : "hidden";
    }
  }

  isNestedAppAuthSupported() {
    return this.host.isNestedAppAuthSupported();
  }

  // Initialize MSAL public client application.
  async initialize(config: AddinConfig) {
    // If auth is not working, enable debug logging to help diagnose.
    this.pca = await createNestablePublicClientApplication(createMsalConfig(config));

    // If Office does not support Nested App Auth provide a sign-out button since the user selects account
    if (!this.isNestedAppAuthSupported() && this.pca.getActiveAccount()) {
      this.setSignOutButtonVisibility(true);
    }
    this.getSignOutButton()?.addEventListener("click", () => this.signOut());
  }

  private async signOut() {
    if (this._usingFallbackDialog) {
      await this.signOutWithDialogApi();
    } else {
      await this.pca?.logoutPopup();
    }

    this.setSignOutButtonVisibility(false);
  }

  /**
   * Tries to get a token without any UI. Returns null if user interaction would be needed.
   * @param scopes the minimum scopes needed.
   */
  async acquireTokenSilentOnly(scopes: string[]): Promise<TokenResult | null> {
    if (this._dialogApiResult) {
      return this._dialogApiResult;
    }

    if (this.pca === undefined) {
      throw new Error("AccountManager is not initialized!");
    }

    const selectAccount = this.pca.getActiveAccount() ? false : true;
    try {
      console.log("Trying to acquire token silently...");
      const authResult = await this.pca.acquireTokenSilent(getTokenRequest(scopes, selectAccount));
      console.log("Acquired token silently.");
      return {
        accessToken: authResult.accessToken,
        username: authResult.account.username,
        authPath: "silent",
      };
    } catch (error) {
      console.warn(`Unable to acquire token silently: ${error}`);
      return null;
    }
  }

  /**
   * Gets a token silently, falling back to a popup and then to the Office dialog.
   * @param scopes the minimum scopes needed.
   */
  async ssoGetAccessToken(scopes: string[]): Promise<TokenResult> {
    const silentResult = await this.acquireTokenSilentOnly(scopes);
    if (silentResult) {
      return silentResult;
    }

    if (this.pca === undefined) {
      throw new Error("AccountManager is not initialized!");
    }

    // Acquire token silent failure. Send an interactive request via popup.
    const selectAccount = this.pca.getActiveAccount() ? false : true;
    try {
      console.log("Trying to acquire token interactively...");
      const authResult = await this.pca.acquireTokenPopup(getTokenRequest(scopes, selectAccount));
      console.log("Acquired token interactively.");
      if (selectAccount) {
        this.pca.setActiveAccount(authResult.account);
      }
      if (!this.isNestedAppAuthSupported()) {
        this.setSignOutButtonVisibility(true);
      }
      return {
        accessToken: authResult.accessToken,
        username: authResult.account.username,
        authPath: "popup",
      };
    } catch (popupError) {
      // Optional fallback if about:blank popup should not be shown
      if (
        this.host.officeDialogFallback &&
        popupError instanceof BrowserAuthError &&
        popupError.errorCode === "popup_window_error"
      ) {
        return this.getTokenWithDialogApi();
      } else {
        // Acquire token interactive failure.
        console.error(`Unable to acquire token interactively: ${popupError}`);
        throw new Error(`Unable to acquire access token: ${popupError}`);
      }
    }
  }

  /**
   * Gets an access token by using the Office dialog API to handle authentication. Used for fallback scenario.
   * @returns The access token.
   */
  async getTokenWithDialogApi(): Promise<TokenResult> {
    this._dialogApiResult = new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(
        createLocalUrl(`dialog.html`),
        { height: 60, width: 30 },
        (result) => {
          result.value.addEventHandler(
            Office.EventType.DialogEventReceived,
            (arg: DialogEventArg) => {
              const errorArg = arg as DialogEventError;
              if (errorArg.error == 12006) {
                this._dialogApiResult = null;
                reject("Dialog closed");
              }
            }
          );
          result.value.addEventHandler(
            Office.EventType.DialogMessageReceived,
            (arg: DialogEventArg) => {
              const messageArg = arg as DialogEventMessage;
              const parsedMessage: AuthDialogResult = JSON.parse(messageArg.message);
              result.value.close();

              if (parsedMessage.error || !parsedMessage.accessToken) {
                reject(parsedMessage.error ?? "No access token returned from dialog");
                this._dialogApiResult = null;
              } else {
                resolve({
                  accessToken: parsedMessage.accessToken,
                  username: parsedMessage.username ?? "",
                  authPath: "dialog",
                });
                this.setSignOutButtonVisibility(true);
                this._usingFallbackDialog = true;
              }
            }
          );
        }
      );
    });
    return this._dialogApiResult;
  }

  signOutWithDialogApi(): Promise<void> {
    return new Promise((resolve) => {
      Office.context.ui.displayDialogAsync(
        createLocalUrl(`dialog.html?logout=1`),
        { height: 60, width: 30 },
        (result) => {
          result.value.addEventHandler(Office.EventType.DialogMessageReceived, () => {
            this.setSignOutButtonVisibility(false);
            this._dialogApiResult = null;
            this._usingFallbackDialog = false;
            resolve();
            result.value.close();
          });
        }
      );
    });
  }
}
