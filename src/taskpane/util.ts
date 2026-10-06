// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* global window, URL */

/**
 * Constructs a URL for a page served alongside the current one, e.g. "auth.html".
 * Relative to the current page so it works at the domain root (local dev) and under a path
 * such as https://<domain>/outlook-addin/ (production).
 * @param path The path relative to the add-in's folder.
 */
export function createLocalUrl(path: string) {
  return new URL(path, window.location.href).toString();
}
