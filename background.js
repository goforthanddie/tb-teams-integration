/* global browser */

const DEFAULT_SCOPES = [
  "OnlineMeetings.ReadWrite",
  "offline_access",
  "openid",
  "profile"
];

async function getSettings() {
  const data = await browser.storage.local.get({
    clientId: "REPLACE_WITH_CLIENT_ID",
    tenant: "organizations",
    authorityHost: "https://login.microsoftonline.com",
    scopes: DEFAULT_SCOPES.join(" "),
    debugEnabled: false
  });
  return data;
}

function isPlaceholder(value) {
  if (!value) {
    return true;
  }
  return value === "REPLACE_WITH_CLIENT_ID";
}

function decodeJwtPayload(token) {
  if (!token || !token.includes(".")) {
    return null;
  }
  const [, payload] = token.split(".");
  if (!payload) {
    return null;
  }
  const normalized = payload.replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized.padEnd(normalized.length + (4 - (normalized.length % 4)) % 4, "=");
  try {
    return JSON.parse(atob(padded));
  } catch (err) {
    return null;
  }
}

function getAccountTypeFromIdToken(idToken) {
  const payload = decodeJwtPayload(idToken);
  if (!payload) {
    return { type: "unknown", tenantId: "" };
  }
  const tenantId = payload.tid || "";
  if (tenantId === "9188040d-6c67-4c5b-b112-36a304b66dad") {
    return { type: "personal", tenantId };
  }
  return { type: "work", tenantId };
}

function base64UrlEncode(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (const b of bytes) {
    binary += String.fromCharCode(b);
  }
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

async function sha256(input) {
  const encoder = new TextEncoder();
  const data = encoder.encode(input);
  return crypto.subtle.digest("SHA-256", data);
}

function randomString(length) {
  const bytes = new Uint8Array(length);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, b => (b % 36).toString(36)).join("");
}

async function buildPkce() {
  const verifier = randomString(96);
  const challenge = base64UrlEncode(await sha256(verifier));
  return { verifier, challenge };
}

async function getAccessToken(interactive = true) {
  const settings = await getSettings();
  if (isPlaceholder(settings.clientId)) {
    await browser.runtime.openOptionsPage();
    throw new Error("Missing client ID. Configure the add-on options first.");
  }

  const tokenState = await browser.storage.local.get({
    accessToken: "",
    refreshToken: "",
    tokenExpiresAt: 0
  });

  const now = Date.now();
  if (tokenState.accessToken && tokenState.tokenExpiresAt > now + 60000) {
    return tokenState.accessToken;
  }

  if (tokenState.refreshToken) {
    const refreshed = await refreshAccessToken(settings, tokenState.refreshToken);
    if (refreshed) {
      return refreshed;
    }
  }

  if (!interactive) {
    throw new Error("No cached token available.");
  }

  const { verifier, challenge } = await buildPkce();
  const redirectUri = browser.identity.getRedirectURL();
  const scopes = settings.scopes || DEFAULT_SCOPES.join(" ");
  const authority = `${settings.authorityHost.replace(/\/$/, "")}/${settings.tenant}`;

  const authUrl = new URL(`${authority}/oauth2/v2.0/authorize`);
  authUrl.searchParams.set("client_id", settings.clientId);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("redirect_uri", redirectUri);
  authUrl.searchParams.set("response_mode", "query");
  authUrl.searchParams.set("scope", scopes);
  authUrl.searchParams.set("code_challenge", challenge);
  authUrl.searchParams.set("code_challenge_method", "S256");
  authUrl.searchParams.set("prompt", "select_account");

  const responseUrl = await browser.identity.launchWebAuthFlow({
    url: authUrl.toString(),
    interactive: true
  });

  const code = new URL(responseUrl).searchParams.get("code");
  if (!code) {
    throw new Error("Authorization failed: no code returned.");
  }

  const token = await exchangeCodeForToken(settings, code, verifier, redirectUri, scopes);
  return token;
}

async function exchangeCodeForToken(settings, code, verifier, redirectUri, scopes) {
  const authority = `${settings.authorityHost.replace(/\/$/, "")}/${settings.tenant}`;
  const body = new URLSearchParams();
  body.set("client_id", settings.clientId);
  body.set("grant_type", "authorization_code");
  body.set("code", code);
  body.set("redirect_uri", redirectUri);
  body.set("code_verifier", verifier);
  body.set("scope", scopes);

  const res = await fetch(`${authority}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
  });

  const json = await res.json();
  if (!res.ok) {
    throw new Error(json.error_description || "Token exchange failed.");
  }

  await persistToken(json);
  return json.access_token;
}

async function refreshAccessToken(settings, refreshToken) {
  const authority = `${settings.authorityHost.replace(/\/$/, "")}/${settings.tenant}`;
  const scopes = settings.scopes || DEFAULT_SCOPES.join(" ");
  const body = new URLSearchParams();
  body.set("client_id", settings.clientId);
  body.set("grant_type", "refresh_token");
  body.set("refresh_token", refreshToken);
  body.set("scope", scopes);

  const res = await fetch(`${authority}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
  });

  if (!res.ok) {
    return null;
  }

  const json = await res.json();
  await persistToken(json);
  return json.access_token;
}

async function persistToken(tokenResponse) {
  const expiresInMs = (tokenResponse.expires_in || 0) * 1000;
  const tokenExpiresAt = Date.now() + expiresInMs;
  await browser.storage.local.set({
    accessToken: tokenResponse.access_token || "",
    refreshToken: tokenResponse.refresh_token || "",
    tokenExpiresAt,
    idToken: tokenResponse.id_token || ""
  });
}

async function createOnlineMeeting(payload) {
  const accessToken = await getAccessToken(true);
  const body = {
    subject: payload.title || "Teams meeting",
    startDateTime: payload.startDateTime,
    endDateTime: payload.endDateTime
  };

  const res = await fetch("https://graph.microsoft.com/v1.0/me/onlineMeetings", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(body)
  });

  const json = await res.json();
  if (!res.ok) {
    const message = json?.error?.message || "Failed to create Teams meeting.";
    throw new Error(message);
  }

  return json.joinWebUrl || "";
}

browser.teamsDialog.onTeamsButtonClick.addListener(async (payload) => {
  try {
    if (!payload.startDateTime || !payload.endDateTime) {
      throw new Error("Event start/end time missing.");
    }
    const joinUrl = await createOnlineMeeting(payload);
    if (!joinUrl) {
      throw new Error("No join URL returned.");
    }
    await browser.teamsDialog.insertJoinInfo({
      dialogId: payload.dialogId,
      joinUrl,
      label: "Microsoft Teams meeting"
    });
  } catch (err) {
    console.error("Teams meeting creation failed:", err);
    await browser.teamsDialog.showError({
      dialogId: payload.dialogId,
      message: String(err.message || err)
    });
  }
});

async function applyDebugSetting() {
  const settings = await getSettings();
  if (browser.teamsDialog && typeof browser.teamsDialog.setDebug === "function") {
    await browser.teamsDialog.setDebug({ enabled: !!settings.debugEnabled });
  }
}

if (browser.teamsDialog && typeof browser.teamsDialog.register === "function") {
  browser.teamsDialog.register();
  applyDebugSetting();

  browser.storage.onChanged.addListener((changes, area) => {
    if (area === "local" && Object.prototype.hasOwnProperty.call(changes, "debugEnabled")) {
      applyDebugSetting();
    }
  });
} else {
  console.warn("teamsDialog experiment API not available.");
}

browser.runtime.onMessage.addListener(async (message) => {
  if (!message || !message.type) {
    return null;
  }
  if (message.type === "getStatus") {
    const settings = await getSettings();
    const tokenState = await browser.storage.local.get({ idToken: "" });
    const accountInfo = getAccountTypeFromIdToken(tokenState.idToken);
    return {
      configured: !isPlaceholder(settings.clientId),
      accountType: accountInfo.type,
      tenantId: accountInfo.tenantId
    };
  }
  if (message.type === "testConnection") {
    const token = await getAccessToken(true);
    const tokenState = await browser.storage.local.get({ idToken: "" });
    const accountInfo = getAccountTypeFromIdToken(tokenState.idToken);
    return {
      ok: !!token,
      accountType: accountInfo.type,
      tenantId: accountInfo.tenantId
    };
  }
  return null;
});
