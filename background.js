/* global browser, DEFAULT_APPLICATION_ID, DEFAULT_TENANT, DEFAULT_AUTHORITY_HOST, DEFAULT_SCOPES, isPlaceholder, validateSettings */

async function getSettings() {
  const data = await browser.storage.local.get({
    clientId: DEFAULT_APPLICATION_ID,
    tenant: DEFAULT_TENANT,
    authorityHost: DEFAULT_AUTHORITY_HOST,
    scopes: DEFAULT_SCOPES.join(" "),
    debugEnabled: false,
    allowCustomAuthorityHost: false
  });
  return data;
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

async function readResponsePayload(res) {
  const text = await res.text();
  if (!text) {
    return { json: null, text: "" };
  }
  try {
    return { json: JSON.parse(text), text };
  } catch (err) {
    return { json: null, text };
  }
}

async function getAccessToken(interactive = true) {
  const settings = await getSettings();
  if (settings.debugEnabled) {
    console.log("[tb-teams] getAccessToken start");
  }
  if (isPlaceholder(settings.clientId)) {
    await browser.runtime.openOptionsPage();
    throw new Error("Missing client ID. Configure the add-on options first.");
  }

  const validation = validateSettings(settings);
  if (!validation.ok) {
    throw new Error(validation.errors.join(" "));
  }
  if (validation.warnings.length && settings.debugEnabled) {
    console.log(`[tb-teams] Settings warning: ${validation.warnings.join(" ")}`);
  }

  const tokenState = await browser.storage.local.get({
    accessToken: "",
    refreshToken: "",
    tokenExpiresAt: 0
  });

  const now = Date.now();
  if (tokenState.accessToken && tokenState.tokenExpiresAt > now + 60000) {
    if (settings.debugEnabled) {
      console.log("[tb-teams] Using cached access token.");
    }
    return tokenState.accessToken;
  }

  if (tokenState.refreshToken) {
    if (settings.debugEnabled) {
      console.log("[tb-teams] Refreshing access token.");
    }
    const refreshed = await refreshAccessToken(settings, tokenState.refreshToken);
    if (refreshed) {
      if (settings.debugEnabled) {
        console.log("[tb-teams] Refresh succeeded.");
      }
      return refreshed;
    }
    if (settings.debugEnabled) {
      console.log("[tb-teams] Refresh failed; falling back to interactive login.");
    }
  }

  if (!interactive) {
    throw new Error("No cached token available.");
  }

  const { verifier, challenge } = await buildPkce();
  const state = randomString(32);
  const redirectUri = browser.identity.getRedirectURL();
  const scopes = settings.scopes || DEFAULT_SCOPES.join(" ");
  const authorityHost = validation.normalized.authorityHost || settings.authorityHost;
  const tenant = validation.normalized.tenant || settings.tenant;
  const authority = `${authorityHost.replace(/\/$/, "")}/${tenant}`;

  const authUrl = new URL(`${authority}/oauth2/v2.0/authorize`);
  authUrl.searchParams.set("client_id", settings.clientId);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("redirect_uri", redirectUri);
  authUrl.searchParams.set("response_mode", "query");
  authUrl.searchParams.set("scope", scopes);
  authUrl.searchParams.set("code_challenge", challenge);
  authUrl.searchParams.set("code_challenge_method", "S256");
  authUrl.searchParams.set("prompt", "select_account");
  authUrl.searchParams.set("state", state);

  if (settings.debugEnabled) {
    console.log("[tb-teams] Launching auth flow.");
  }
  const responseUrl = await browser.identity.launchWebAuthFlow({
    url: authUrl.toString(),
    interactive: true
  });

  const responseParams = new URL(responseUrl).searchParams;
  const authError = responseParams.get("error");
  if (authError) {
    const description = responseParams.get("error_description") || authError;
    throw new Error(`Authorization failed: ${description}`);
  }

  const returnedState = responseParams.get("state");
  if (returnedState !== state) {
    throw new Error("Authorization failed: invalid state.");
  }

  const code = responseParams.get("code");
  if (!code) {
    throw new Error("Authorization failed: no code returned.");
  }

  const token = await exchangeCodeForToken(settings, code, verifier, redirectUri, scopes);
  if (settings.debugEnabled) {
    console.log("[tb-teams] Token exchange succeeded.");
  }
  return token;
}

async function exchangeCodeForToken(settings, code, verifier, redirectUri, scopes) {
  const validation = validateSettings(settings);
  const authorityHost = validation.normalized.authorityHost || settings.authorityHost;
  const tenant = validation.normalized.tenant || settings.tenant;
  const authority = `${authorityHost.replace(/\/$/, "")}/${tenant}`;
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

  const { json, text } = await readResponsePayload(res);
  if (!res.ok) {
    const message = json?.error_description || json?.error?.message || text || `Token exchange failed (HTTP ${res.status}).`;
    throw new Error(message);
  }
  if (!json) {
    throw new Error("Token exchange failed: invalid response.");
  }

  await persistToken(json);
  return json.access_token;
}

async function refreshAccessToken(settings, refreshToken) {
  const validation = validateSettings(settings);
  const authorityHost = validation.normalized.authorityHost || settings.authorityHost;
  const tenant = validation.normalized.tenant || settings.tenant;
  const authority = `${authorityHost.replace(/\/$/, "")}/${tenant}`;
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

  const { json } = await readResponsePayload(res);
  if (!res.ok || !json) {
    return null;
  }

  await persistToken(json, refreshToken);
  return json.access_token;
}

async function persistToken(tokenResponse, existingRefreshToken = "") {
  const expiresInMs = (tokenResponse.expires_in || 0) * 1000;
  const tokenExpiresAt = Date.now() + expiresInMs;
  const refreshToken = tokenResponse.refresh_token || existingRefreshToken || "";
  await browser.storage.local.set({
    accessToken: tokenResponse.access_token || "",
    refreshToken,
    tokenExpiresAt,
    idToken: tokenResponse.id_token || ""
  });
}

async function createOnlineMeeting(payload) {
  const settings = await getSettings();
  if (settings.debugEnabled) {
    console.log("[tb-teams] Creating online meeting.");
  }
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

  if (settings.debugEnabled) {
    console.log("[tb-teams] Online meeting created.");
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
