/* global document, HTMLElement */

export function resultCard(options: {
  name: string;
  details: string[];
  action: string;
  onChoose: () => void;
}) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "room";
  const title = document.createElement("span");
  title.className = "item-title";
  title.textContent = options.name;
  button.appendChild(title);
  for (const text of options.details.filter(Boolean)) {
    const detail = document.createElement("span");
    detail.className = "item-detail";
    detail.textContent = text;
    button.appendChild(detail);
  }
  const footer = document.createElement("span");
  footer.className = "result-footer";
  const availability = document.createElement("span");
  availability.className = "result-availability";
  availability.textContent = "Available";
  const action = document.createElement("span");
  action.className = "result-action";
  action.textContent = options.action;
  footer.appendChild(availability);
  footer.appendChild(action);
  button.appendChild(footer);
  button.addEventListener("click", options.onChoose);
  return button;
}

export function selectedResult(card: HTMLElement, label: string): void {
  const status = document.createElement("div");
  status.className = "selection-label";
  status.textContent = `✓ ${label}`;
  card.appendChild(status);
}
