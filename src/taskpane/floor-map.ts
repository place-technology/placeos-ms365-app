/*
 * Floor plan map for the Book a room view, a lightweight port of the legacy `interactive-map`
 * (PlaceOS user-interfaces libs/components, with libs/events space-map.component for the room styling).
 *
 * A level's map is an SVG at the level zone's `map_id`. Rooms are the SVG elements whose id is the system's
 * `map_id`. The SVG is inlined (after removing scripts, event handlers and external links, and scoping its
 * <style> rules to the map) so rooms can be coloured and tapped. Pan and zoom change the SVG's viewBox,
 * which works in the IE-based Outlook webviews too.
 */

/* global document, window, fetch, DOMParser, Element, HTMLElement, HTMLStyleElement, SVGElement, SVGSVGElement, CSSStyleRule, CSSMediaRule, CSSRuleList, PointerEvent, WheelEvent, KeyboardEvent, EventTarget, URL */

import type { PlaceosApi } from "./placeos-data";

declare const __DEV_TOOLS__: boolean;

export type MapRoomStatus = "free" | "busy" | "filtered";

export type MapRoom = {
  // Id of the room's element in the SVG.
  elementId: string;
  status: MapRoomStatus;
  title: string;
  // Free rooms that match the filters can be chosen.
  selectable: boolean;
};

/** Workplace's DEFAULT_COLOURS (libs/explore explore-spaces.service): free, busy and not-bookable. */
export const defaultMapColours: Record<MapRoomStatus, string> = {
  free: "#43a047",
  busy: "#e53935",
  filtered: "#757575",
};

type ViewBox = { x: number; y: number; width: number; height: number };

const maxZoom = 10;
// Pointer movement (px) before a press counts as a drag rather than a tap.
const dragThreshold = 5;
// Marks what this module adds to or sets on room elements, so it can be undone on the next render.
const styledAttribute = "data-placeos-room";

// ---------------------------------------------------------------------------------------------------------
// Loading
// ---------------------------------------------------------------------------------------------------------

const mapCache = new Map<string, Promise<string>>();

/**
 * The SVG text of a map, cached per URL. Floor plans are normally on S3 (the level zone's map_id), fetched
 * without credentials, so the bucket's CORS rules must allow the domain (as Workplace needs too). Maps on this
 * domain are fetched with the PlaceOS token, as Workplace does.
 */
export function loadMapSvg(url: string, api: PlaceosApi): Promise<string> {
  let cached = mapCache.get(url);
  if (!cached) {
    cached = fetchMapSvg(url, api);
    mapCache.set(url, cached);
    cached.catch(() => mapCache.delete(url));
  }
  return cached;
}

async function fetchMapSvg(url: string, api: PlaceosApi): Promise<string> {
  // Relative map_ids are taken as paths on this domain.
  const resolved = new URL(url, window.location.origin);
  if (resolved.origin === window.location.origin) {
    return api<string>(resolved.pathname + resolved.search, {
      text: true,
      headers: { Accept: "image/svg+xml, */*" },
    });
  }
  // The S3 bucket's CORS rules don't allow localhost, so the dev server proxies maps (webpack.config.js).
  const devProxy =
    typeof __DEV_TOOLS__ !== "undefined" &&
    __DEV_TOOLS__ &&
    window.location.hostname === "localhost" &&
    /(^|\.)amazonaws\.com$/i.test(resolved.hostname);
  const href = devProxy
    ? `/__map-proxy/${resolved.host}${resolved.pathname}${resolved.search}`
    : resolved.href;
  const response = await fetch(href, { credentials: "omit" });
  if (!response.ok) {
    throw new Error(`The floor plan request failed with ${response.status}.`);
  }
  return response.text();
}

// ---------------------------------------------------------------------------------------------------------
// Sanitising
// ---------------------------------------------------------------------------------------------------------

// Elements that can run script, embed other documents or change attributes (animations can set an href).
const blockedElements = [
  "script",
  "foreignObject",
  "iframe",
  "object",
  "embed",
  "audio",
  "video",
  "animate",
  "animateMotion",
  "animateTransform",
  "set",
  "handler",
  "listener",
];

function sanitise(svg: Element) {
  for (const name of blockedElements) {
    Array.from(svg.getElementsByTagName(name)).forEach((node) =>
      node.parentNode?.removeChild(node)
    );
  }
  const all = [svg, ...Array.from(svg.getElementsByTagName("*"))];
  for (const node of all) {
    for (const attribute of Array.from(node.attributes)) {
      const name = attribute.name.toLowerCase();
      const value = attribute.value.trim().toLowerCase();
      if (name.indexOf("on") === 0) {
        node.removeAttribute(attribute.name);
      } else if (name === "href" || name === "xlink:href") {
        // Internal references are fine; images may also load pictures. Nothing else may link out.
        const allowed =
          value.charAt(0) === "#" ||
          (node.localName === "image" &&
            (value.indexOf("data:image/") === 0 ||
              !/^[a-z][a-z0-9+.-]*:/.test(value) ||
              value.indexOf("https:") === 0));
        if (!allowed) {
          node.removeAttribute(attribute.name);
        }
      }
    }
  }
}

/**
 * Prefixes every selector in the SVG's <style> rules with `scope`. Inline SVG styles apply to the whole page,
 * so unscoped rules (e.g. `text { … }`) would restyle the add-in.
 */
function scopeStyles(svg: Element, scope: string) {
  for (const style of Array.from(svg.getElementsByTagName("style"))) {
    const parsing = document.createElement("style");
    // Parsed by the browser but never applied.
    parsing.media = "not all";
    parsing.textContent = style.textContent || "";
    document.head.appendChild(parsing);
    try {
      const rules = (parsing as HTMLStyleElement).sheet?.cssRules;
      style.textContent = rules ? scopeRules(rules, scope) : "";
    } catch {
      style.textContent = "";
    } finally {
      document.head.removeChild(parsing);
    }
  }
}

function scopeRules(rules: CSSRuleList, scope: string): string {
  let css = "";
  for (const rule of Array.from(rules)) {
    if (rule.type === 1) {
      const styleRule = rule as CSSStyleRule;
      const selectors = styleRule.selectorText
        .split(",")
        .map((selector) => `${scope} ${selector.trim()}`)
        .join(", ");
      css += `${selectors} { ${styleRule.style.cssText} }\n`;
    } else if (rule.type === 4) {
      const media = rule as CSSMediaRule;
      css += `@media ${media.media.mediaText} { ${scopeRules(media.cssRules, scope)} }\n`;
    } else if (rule.type !== 3) {
      // Not @import; @font-face and @keyframes don't select page elements.
      css += `${rule.cssText}\n`;
    }
  }
  return css;
}

function quoteAttribute(value: string) {
  return `"${value.replace(/["\\]/g, "\\$&")}"`;
}

// ---------------------------------------------------------------------------------------------------------
// The map
// ---------------------------------------------------------------------------------------------------------

export class FloorMap {
  private svg: SVGSVGElement | null = null;
  private base: ViewBox = { x: 0, y: 0, width: 1, height: 1 };
  private view: ViewBox = { x: 0, y: 0, width: 1, height: 1 };
  private rooms = new Map<Element, MapRoom>();
  private selected: Element | null = null;
  private selectedColour = "#0078d4";
  // Active pointers (for pan and pinch), the press target and whether the press has become a drag.
  private pointers = new Map<number, { x: number; y: number }>();
  private pressTarget: Element | null = null;
  private pressStart = { x: 0, y: 0 };
  private dragged = false;

  /** Called with the element id of a selectable room the user tapped or pressed Enter on. */
  onSelect: (elementId: string) => void = () => undefined;

  constructor(private readonly container: HTMLElement) {
    container.addEventListener("pointerdown", (event) => this.pointerDown(event));
    container.addEventListener("pointermove", (event) => this.pointerMove(event));
    container.addEventListener("pointerup", (event) => this.pointerUp(event, true));
    container.addEventListener("pointercancel", (event) => this.pointerUp(event, false));
    container.addEventListener("wheel", (event) => this.wheel(event as WheelEvent), {
      passive: false,
    });
    container.addEventListener("keydown", (event) => this.keyDown(event as KeyboardEvent));
  }

  /** Replaces the map with this SVG. Throws if it isn't a valid SVG. */
  show(svgText: string) {
    const parsed = new DOMParser().parseFromString(svgText, "image/svg+xml");
    const root = parsed.documentElement;
    if (!root || root.localName !== "svg" || parsed.getElementsByTagName("parsererror").length) {
      throw new Error("The floor plan isn't a valid SVG.");
    }
    sanitise(root);
    scopeStyles(root, `#${this.container.id}`);
    const svg = document.importNode(root, true) as unknown as SVGSVGElement;

    const box = (svg.getAttribute("viewBox") || "").split(/[\s,]+/).map(parseFloat);
    if (box.length === 4 && box.every((n) => !isNaN(n)) && box[2] > 0 && box[3] > 0) {
      this.base = { x: box[0], y: box[1], width: box[2], height: box[3] };
    } else {
      const width = parseFloat(svg.getAttribute("width") || "") || 1000;
      const height = parseFloat(svg.getAttribute("height") || "") || 1000;
      this.base = { x: 0, y: 0, width, height };
    }
    svg.setAttribute("width", "100%");
    svg.setAttribute("height", "100%");
    svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
    svg.setAttribute("focusable", "false");
    // The legacy add-in hides the zones layer.
    for (const id of ["zones", "Zones"]) {
      const layer = svg.querySelector(`[id=${quoteAttribute(id)}]`) as SVGElement | null;
      if (layer) {
        layer.style.display = "none";
      }
    }

    this.container.textContent = "";
    this.container.appendChild(svg);
    this.svg = svg;
    this.rooms.clear();
    this.selected = null;
    this.reset();
  }

  clear() {
    this.container.textContent = "";
    this.svg = null;
    this.rooms.clear();
    this.selected = null;
  }

  /**
   * Colours rooms by status; selectable ones get a tooltip and can be tapped or focused.
   * @returns The element ids that aren't on the map.
   */
  setRooms(
    rooms: MapRoom[],
    colours: Record<MapRoomStatus, string>,
    selectedColour: string
  ): string[] {
    this.selectedColour = selectedColour;
    this.rooms.forEach((_, element) => this.unstyle(element));
    this.rooms.clear();
    const missing: string[] = [];
    if (!this.svg) {
      return rooms.map((room) => room.elementId);
    }
    for (const room of rooms) {
      const element = this.svg.querySelector(
        `[id=${quoteAttribute(room.elementId)}]`
      ) as SVGElement | null;
      if (!element) {
        missing.push(room.elementId);
        continue;
      }
      this.rooms.set(element, room);
      element.setAttribute(styledAttribute, room.status);
      element.style.fill = colours[room.status];
      element.style.opacity = "0.6";
      if (room.selectable) {
        element.style.cursor = "pointer";
        element.setAttribute("tabindex", "0");
        element.setAttribute("role", "button");
        element.setAttribute("aria-label", room.title);
      }
      const title = document.createElementNS("http://www.w3.org/2000/svg", "title");
      title.setAttribute(styledAttribute, "");
      title.textContent = room.title;
      element.insertBefore(title, element.firstChild);
    }
    if (this.selected && !this.rooms.has(this.selected)) {
      this.selected = null;
    }
    this.highlight();
    return missing;
  }

  /** Outlines a room, or none. */
  select(elementId: string | null) {
    this.selected = null;
    this.rooms.forEach((room, element) => {
      if (room.elementId === elementId) {
        this.selected = element;
      }
    });
    this.highlight();
  }

  zoomBy(factor: number) {
    const rect = this.container.getBoundingClientRect();
    this.zoomAt(factor, rect.left + rect.width / 2, rect.top + rect.height / 2);
  }

  reset() {
    this.view = { ...this.base };
    this.apply();
  }

  /**
   * Centres the map on a room and zooms in on it (to `zoom` times, or further to fit it about three times
   * over). Uses its on-screen box, so transforms in the SVG count; the map must be visible.
   */
  focus(elementId: string, zoom = 4) {
    let target: Element | null = null;
    this.rooms.forEach((room, element) => {
      if (room.elementId === elementId) {
        target = element;
      }
    });
    if (!this.svg || !target) {
      return;
    }
    this.reset();
    const box = (target as Element).getBoundingClientRect();
    const { rect, scale, offsetX, offsetY } = this.layout(this.view);
    if (!box.width && !box.height) {
      return;
    }
    const toSvgX = (clientX: number) => this.view.x + (clientX - rect.left - offsetX) / scale;
    const toSvgY = (clientY: number) => this.view.y + (clientY - rect.top - offsetY) / scale;
    const centreX = toSvgX(box.left + box.width / 2);
    const centreY = toSvgY(box.top + box.height / 2);
    const width = Math.min(
      this.base.width,
      Math.max(this.base.width / maxZoom, this.base.width / zoom, (box.width / scale) * 3 || 0)
    );
    const height = width * (this.base.height / this.base.width);
    this.view = { x: centreX - width / 2, y: centreY - height / 2, width, height };
    this.clampView();
    this.apply();
  }

  private unstyle(element: Element) {
    const svgElement = element as SVGElement;
    svgElement.style.fill = "";
    svgElement.style.opacity = "";
    svgElement.style.cursor = "";
    svgElement.style.stroke = "";
    svgElement.style.strokeWidth = "";
    svgElement.style.removeProperty("vector-effect");
    for (const name of [styledAttribute, "tabindex", "role", "aria-label"]) {
      element.removeAttribute(name);
    }
    Array.from(element.childNodes)
      .filter((child) => child instanceof Element && child.hasAttribute(styledAttribute))
      .forEach((child) => element.removeChild(child));
  }

  private highlight() {
    this.rooms.forEach((_, element) => {
      const svgElement = element as SVGElement;
      const selected = element === this.selected;
      svgElement.style.stroke = selected ? this.selectedColour : "";
      svgElement.style.strokeWidth = selected ? "3px" : "";
      svgElement.style.opacity = selected ? "0.9" : "0.6";
      if (selected) {
        svgElement.style.setProperty("vector-effect", "non-scaling-stroke");
      } else {
        svgElement.style.removeProperty("vector-effect");
      }
    });
  }

  /** The selectable room an event target is in, if any. (No Element.closest on SVG in IE.) */
  private roomAt(target: EventTarget | null): MapRoom | null {
    let node = target as Element | null;
    while (node && node !== this.container) {
      const room = this.rooms.get(node);
      if (room) {
        return room.selectable ? room : null;
      }
      node = node.parentNode as Element | null;
    }
    return null;
  }

  // Pan and zoom ------------------------------------------------------------------------------------------

  private apply() {
    const { x, y, width, height } = this.view;
    this.svg?.setAttribute("viewBox", `${x} ${y} ${width} ${height}`);
  }

  /** Pixels per SVG unit and the letterbox offset, for a view in the container. */
  private layout(view: ViewBox) {
    const rect = this.container.getBoundingClientRect();
    const scale = Math.min(rect.width / view.width, rect.height / view.height) || 1;
    return {
      rect,
      scale,
      offsetX: (rect.width - view.width * scale) / 2,
      offsetY: (rect.height - view.height * scale) / 2,
    };
  }

  private zoomAt(factor: number, clientX: number, clientY: number) {
    if (!this.svg) {
      return;
    }
    const current = this.layout(this.view);
    // The SVG point under the pointer stays where it is.
    const pointX = this.view.x + (clientX - current.rect.left - current.offsetX) / current.scale;
    const pointY = this.view.y + (clientY - current.rect.top - current.offsetY) / current.scale;
    const width = Math.min(
      this.base.width,
      Math.max(this.base.width / maxZoom, this.view.width / factor)
    );
    const height = width * (this.base.height / this.base.width);
    const next = this.layout({ x: 0, y: 0, width, height });
    this.view = {
      x: pointX - (clientX - next.rect.left - next.offsetX) / next.scale,
      y: pointY - (clientY - next.rect.top - next.offsetY) / next.scale,
      width,
      height,
    };
    this.clampView();
    this.apply();
  }

  private panBy(dx: number, dy: number) {
    const { scale } = this.layout(this.view);
    this.view.x -= dx / scale;
    this.view.y -= dy / scale;
    this.clampView();
    this.apply();
  }

  /** Keeps the centre of the view on the map. */
  private clampView() {
    const centreX = Math.min(
      this.base.x + this.base.width,
      Math.max(this.base.x, this.view.x + this.view.width / 2)
    );
    const centreY = Math.min(
      this.base.y + this.base.height,
      Math.max(this.base.y, this.view.y + this.view.height / 2)
    );
    this.view.x = centreX - this.view.width / 2;
    this.view.y = centreY - this.view.height / 2;
  }

  private pointerDown(event: PointerEvent) {
    if (!this.svg || (event.pointerType === "mouse" && event.button !== 0)) {
      return;
    }
    if (this.pointers.size === 0) {
      this.pressTarget = event.target as Element;
      this.pressStart = { x: event.clientX, y: event.clientY };
      this.dragged = false;
    } else {
      // A second finger: pinch, not a tap.
      this.dragged = true;
    }
    this.pointers.set(event.pointerId, { x: event.clientX, y: event.clientY });
    try {
      this.container.setPointerCapture(event.pointerId);
    } catch {
      // Not supported everywhere; panning still works while the pointer stays over the map.
    }
  }

  private pointerMove(event: PointerEvent) {
    const previous = this.pointers.get(event.pointerId);
    if (!previous) {
      return;
    }
    const current = { x: event.clientX, y: event.clientY };
    if (this.pointers.size === 1) {
      if (
        !this.dragged &&
        Math.abs(current.x - this.pressStart.x) + Math.abs(current.y - this.pressStart.y) >
          dragThreshold
      ) {
        this.dragged = true;
      }
      if (this.dragged) {
        this.panBy(current.x - previous.x, current.y - previous.y);
      }
    } else if (this.pointers.size === 2) {
      const other = Array.from(this.pointers.entries()).find(([id]) => id !== event.pointerId)?.[1];
      if (other) {
        const before = Math.hypot(previous.x - other.x, previous.y - other.y);
        const after = Math.hypot(current.x - other.x, current.y - other.y);
        if (before > 0) {
          this.zoomAt(after / before, (current.x + other.x) / 2, (current.y + other.y) / 2);
        }
      }
    }
    this.pointers.set(event.pointerId, current);
  }

  private pointerUp(event: PointerEvent, completed: boolean) {
    if (!this.pointers.has(event.pointerId)) {
      return;
    }
    this.pointers.delete(event.pointerId);
    if (completed && this.pointers.size === 0 && !this.dragged) {
      const room = this.roomAt(this.pressTarget);
      if (room) {
        this.onSelect(room.elementId);
      }
    }
    if (this.pointers.size === 0) {
      this.pressTarget = null;
    }
  }

  private wheel(event: WheelEvent) {
    if (!this.svg) {
      return;
    }
    event.preventDefault();
    this.zoomAt(event.deltaY < 0 ? 1.2 : 1 / 1.2, event.clientX, event.clientY);
  }

  private keyDown(event: KeyboardEvent) {
    if (event.key !== "Enter" && event.key !== " ") {
      return;
    }
    const room = this.roomAt(event.target);
    if (room) {
      event.preventDefault();
      this.onSelect(room.elementId);
    }
  }
}
