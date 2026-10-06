/*
 * DEV ONLY: lets a developer reveal and copy the raw Entra access token, e.g. to test auth.cr before
 * the Phase 4a validation endpoint exists. Only called when __DEV_TOOLS__ is true (development
 * builds), so production builds drop it. The token is a bearer credential: it is never logged, and
 * it is only put in the DOM while revealed.
 */

/* global document, navigator, setTimeout, HTMLElement, HTMLTextAreaElement */

let currentToken = "";

export function renderRawTokenDevOnly(container: HTMLElement, accessToken: string) {
  currentToken = accessToken;
  container.textContent = "";

  const toggle = document.createElement("button");
  toggle.className = "secondary";
  toggle.textContent = "Show raw token (dev only)";

  const copy = document.createElement("button");
  copy.className = "secondary";
  copy.textContent = "Copy";
  copy.style.display = "none";

  const textarea = document.createElement("textarea") as HTMLTextAreaElement;
  textarea.className = "raw-token";
  textarea.readOnly = true;
  textarea.rows = 6;
  textarea.style.display = "none";

  toggle.addEventListener("click", () => {
    const show = textarea.style.display === "none";
    textarea.value = show ? currentToken : "";
    textarea.style.display = show ? "" : "none";
    copy.style.display = show ? "" : "none";
    toggle.textContent = show ? "Hide raw token" : "Show raw token (dev only)";
  });

  copy.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(currentToken);
    } catch {
      // The Outlook iframe may block the async clipboard API; fall back to selecting the text.
      textarea.focus();
      textarea.select();
      document.execCommand("copy");
    }
    copy.textContent = "Copied";
    setTimeout(() => (copy.textContent = "Copy"), 2000);
  });

  container.appendChild(toggle);
  container.appendChild(copy);
  container.appendChild(textarea);
}
