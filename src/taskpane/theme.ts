/* Applies the host's theme to <html data-theme>, which selects the colour variables in taskpane.css. */

/* global document */

/**
 * @param theme TeamsJS theme names ("default", "dark", "contrast"). Anything else counts as light.
 * Undefined leaves the attribute unset, so the OS preference (prefers-color-scheme) applies.
 */
export function applyTheme(theme: string | undefined) {
  const root = document.documentElement;
  if (theme === undefined) {
    root.removeAttribute("data-theme");
  } else {
    root.setAttribute("data-theme", theme === "dark" || theme === "contrast" ? theme : "light");
  }
}

/**
 * Outlook doesn't support officeTheme.isDarkTheme, so infer dark mode from the theme's background colour.
 * @param color e.g. "#1F1F1F". Undefined if the client doesn't report a theme.
 */
export function themeFromBackground(color: string | undefined): string | undefined {
  const match = color && /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(color.trim());
  if (!match) {
    return undefined;
  }
  const [r, g, b] = match.slice(1).map((hex) => parseInt(hex, 16) / 255);
  const luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return luminance < 0.5 ? "dark" : "default";
}
