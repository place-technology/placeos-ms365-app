/*
 * DEV ONLY: Phase 4 failure tests. Sends deliberately bad tokens to the auth.cr exchange and checks they're
 * rejected, so no one has to copy tokens around. Only rendered when __DEV_TOOLS__ is true; production
 * builds drop it. Tokens are never logged or displayed.
 */

/* global document, window, setInterval, HTMLElement, HTMLButtonElement */

import {
  exchangeForPlaceosToken,
  PlaceosRequestError,
  refreshPlaceosToken,
} from "./placeos-helper";
import { decodeJwtPayload } from "./token-inspector";

export type FailureTestDeps = {
  placeosClientId: string;
  // A fresh, valid Entra token for the PlaceOS scope (silent).
  getEntraToken: () => Promise<string>;
  // A Microsoft Graph token from the same sign-in, or null if one can't be had silently.
  getGraphToken: () => Promise<string | null>;
};

type Outcome = { status: number; code: string };
type TestCase = {
  name: string;
  expected: string;
  run: () => Promise<Outcome | string>; // a string means "skipped: <reason>"
  passes: (outcome: Outcome) => boolean;
};

// The first valid token seen, kept (in memory) until it expires for the expired-token test.
let heldToken: { token: string; exp: number } | null = null;

export function holdTokenForExpiryTest(token: string) {
  if (heldToken) {
    return;
  }
  const { exp } = decodeJwtPayload<{ exp?: number }>(token);
  if (exp) {
    heldToken = { token, exp };
  }
}

function base64UrlEncode(text: string): string {
  // Percent-decode to bytes so non-ASCII claims survive btoa.
  const binary = encodeURIComponent(text).replace(/%([0-9A-F]{2})/g, (_, hex) =>
    String.fromCharCode(parseInt(hex, 16))
  );
  return window.btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function withBadSignature(token: string): string {
  const [header, payload, signature] = token.split(".");
  // Flip the first character of the signature; still valid base64url, wrong signature.
  const first = signature.charAt(0) === "A" ? "B" : "A";
  return `${header}.${payload}.${first}${signature.slice(1)}`;
}

function withTamperedTenant(token: string): string {
  const [header, , signature] = token.split(".");
  const claims = decodeJwtPayload<Record<string, unknown>>(token);
  claims.tid = "00000000-0000-0000-0000-000000000000";
  return `${header}.${base64UrlEncode(JSON.stringify(claims))}.${signature}`;
}

async function outcomeOf(request: Promise<unknown>): Promise<Outcome> {
  try {
    await request;
    return { status: 200, code: "ok" };
  } catch (error) {
    if (error instanceof PlaceosRequestError) {
      return { status: error.status, code: error.body };
    }
    return { status: 0, code: error instanceof Error ? error.message : `${error}` };
  }
}

const rejected = (outcome: Outcome) => outcome.status === 400 && outcome.code === "invalid_grant";

export function renderFailureTestsDevOnly(container: HTMLElement, deps: FailureTestDeps) {
  const exchange = (token: string) =>
    outcomeOf(exchangeForPlaceosToken(token, deps.placeosClientId));

  const cases: TestCase[] = [
    {
      name: "Control: valid Entra token",
      expected: "200 (accepted)",
      run: async () => exchange(await deps.getEntraToken()),
      passes: (o) => o.status === 200,
    },
    {
      name: "Malformed token (not a JWT)",
      expected: "400 invalid_grant",
      run: () => exchange("not.a.jwt"),
      passes: rejected,
    },
    {
      name: "Invalid signature",
      expected: "400 invalid_grant",
      run: async () => exchange(withBadSignature(await deps.getEntraToken())),
      passes: rejected,
    },
    {
      name: "Tampered tid (other tenant, original signature)",
      expected: "400 invalid_grant",
      run: async () => exchange(withTamperedTenant(await deps.getEntraToken())),
      passes: rejected,
    },
    {
      name: "Wrong audience (Microsoft Graph token)",
      expected: "400 invalid_grant",
      run: async () => {
        const graphToken = await deps.getGraphToken();
        return graphToken
          ? exchange(graphToken)
          : "couldn't get a Graph token silently (no User.Read consent?)";
      },
      passes: rejected,
    },
    {
      name: "Expired token",
      expected: "400 invalid_grant",
      run: async () => {
        if (!heldToken) {
          return "no token held yet; sign in first";
        }
        if (heldToken.exp * 1000 > Date.now()) {
          return `held token expires at ${new Date(heldToken.exp * 1000).toLocaleTimeString()}; run again after that`;
        }
        return exchange(heldToken.token);
      },
      passes: rejected,
    },
    {
      name: "Bogus PlaceOS refresh token",
      expected: "400 invalid_grant",
      run: () => outcomeOf(refreshPlaceosToken("not-a-refresh-token", deps.placeosClientId)),
      passes: rejected,
    },
  ];

  container.textContent = "";
  const heading = document.createElement("h2");
  heading.textContent = "Failure tests (dev only)";
  const note = document.createElement("p");
  note.className = "detail";
  note.textContent =
    "Not testable here: a genuine token from another tenant (needs a user there) and an app-only / " +
    "missing-scope token (needs a client secret). Cover those with auth.cr's spec tests. " +
    "Reasons for rejections are in the auth.cr logs.";
  const runButton = document.createElement("button") as HTMLButtonElement;
  runButton.className = "secondary";
  runButton.textContent = "Run failure tests";
  const results = document.createElement("dl");
  results.className = "claims";
  const expiry = document.createElement("p");
  expiry.className = "detail";

  container.appendChild(heading);
  container.appendChild(note);
  container.appendChild(runButton);
  container.appendChild(expiry);
  container.appendChild(results);

  const updateExpiry = () => {
    if (heldToken) {
      const minutes = Math.ceil((heldToken.exp * 1000 - Date.now()) / 60000);
      expiry.textContent =
        minutes > 0
          ? `Expired-token test available in ${minutes} min (${new Date(heldToken.exp * 1000).toLocaleTimeString()}).`
          : "Expired-token test ready.";
    }
  };
  updateExpiry();
  setInterval(updateExpiry, 30000);

  runButton.addEventListener("click", async () => {
    runButton.disabled = true;
    results.textContent = "";
    for (const test of cases) {
      const dt = document.createElement("dt");
      dt.textContent = `${test.name}: running...`;
      const dd = document.createElement("dd");
      results.appendChild(dt);
      results.appendChild(dd);

      let result: Outcome | string;
      try {
        result = await test.run();
      } catch (error) {
        result = `error: ${error instanceof Error ? error.message : error}`;
      }
      if (typeof result === "string") {
        dt.textContent = `${test.name}: skipped`;
        dd.className = "expected";
        dd.textContent = result;
      } else {
        const passed = test.passes(result);
        dt.textContent = `${test.name} ${passed ? "✓" : "✗"}`;
        dt.className = passed ? "ok" : "error";
        dd.textContent = `got ${result.status} ${result.code}; expected ${test.expected}`;
      }
    }
    runButton.disabled = false;
  });
}
