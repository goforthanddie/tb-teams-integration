/* global browser, DEFAULT_APPLICATION_ID, DEFAULT_TENANT, DEFAULT_AUTHORITY_HOST, DEFAULT_SCOPES, DEFAULT_ACCOUNT_MODE, DEFAULT_MEETING_MODE, isPlaceholder, validateSettings */

const form = document.getElementById("settings-form");
const statusEl = document.getElementById("status");
const testButton = document.getElementById("testConnection");
const testStatusEl = document.getElementById("testStatus");
const configStatusEl = document.getElementById("configStatus");
const accountStatusEl = document.getElementById("accountStatus");
const accountWarningEl = document.getElementById("accountWarning");
const personalAccountNoteEl = document.getElementById("personalAccountNote");
const copyRedirectButton = document.getElementById("copyRedirect");
const redirectStatusEl = document.getElementById("redirectStatus");
const logoutButton = document.getElementById("logout");
const clientIdRow = document.getElementById("clientIdRow");
const accountModeInputs = Array.from(document.querySelectorAll("input[name=\"accountMode\"]"));
const meetingModeInputs = Array.from(document.querySelectorAll("input[name=\"meetingMode\"]"));
let lastWorkTenant = DEFAULT_TENANT;

function getSelectedAccountMode() {
  const selected = accountModeInputs.find(input => input.checked);
  return selected ? selected.value : DEFAULT_ACCOUNT_MODE;
}

function getSelectedMeetingMode() {
  const selected = meetingModeInputs.find(input => input.checked);
  return selected ? selected.value : DEFAULT_MEETING_MODE;
}

function setMeetingMode(value) {
  for (const input of meetingModeInputs) {
    input.checked = input.value === value;
  }
}

function updateModeState() {
  const accountMode = getSelectedAccountMode();
  const tenantInput = document.getElementById("tenant");
  if (accountMode === "personal") {
    if (!tenantInput.disabled) {
      lastWorkTenant = tenantInput.value.trim() || DEFAULT_TENANT;
    }
    tenantInput.value = "consumers";
    tenantInput.disabled = true;
    for (const input of meetingModeInputs) {
      if (input.value === "direct") {
        input.checked = false;
        input.disabled = true;
      } else {
        input.disabled = false;
        input.checked = true;
      }
    }
  } else {
    if (tenantInput.disabled) {
      tenantInput.value = lastWorkTenant || DEFAULT_TENANT;
    }
    tenantInput.disabled = false;
    for (const input of meetingModeInputs) {
      input.disabled = false;
    }
  }

  const meetingMode = getSelectedMeetingMode();
  const clientIdInput = document.getElementById("clientId");
  if (meetingMode === "direct") {
    clientIdRow.classList.remove("hidden");
    clientIdInput.required = true;
  } else {
    clientIdRow.classList.remove("hidden");
    clientIdInput.required = false;
  }
  clientIdInput.placeholder = "Required for all accounts";
}

function getScopesForMode(meetingMode) {
  if (meetingMode === "calendar") {
    return "Calendars.ReadWrite offline_access openid profile";
  }
  return "OnlineMeetings.ReadWrite offline_access openid profile";
}

async function loadSettings() {
  const data = await browser.storage.local.get({
    clientId: "",
    tenant: DEFAULT_TENANT,
    authorityHost: DEFAULT_AUTHORITY_HOST,
    scopes: DEFAULT_SCOPES.join(" "),
    debugEnabled: false,
    allowCustomAuthorityHost: false,
    accountMode: DEFAULT_ACCOUNT_MODE,
    meetingMode: DEFAULT_MEETING_MODE,
    useDefaultApplicationId: false
  });

  const clientIdEl = document.getElementById("clientId");
  const tenantEl = document.getElementById("tenant");
  const authorityHostEl = document.getElementById("authorityHost");
  clientIdEl.value = data.clientId;
  tenantEl.value = data.tenant;
  authorityHostEl.value = data.authorityHost;
  lastWorkTenant = tenantEl.value || DEFAULT_TENANT;

  clientIdEl.placeholder = "Required for all accounts";
  tenantEl.placeholder = DEFAULT_TENANT;
  authorityHostEl.placeholder = DEFAULT_AUTHORITY_HOST;
  document.getElementById("debugEnabled").checked = !!data.debugEnabled;
  document.getElementById("allowCustomAuthorityHost").checked = !!data.allowCustomAuthorityHost;

  for (const input of accountModeInputs) {
    input.checked = input.value === (data.accountMode || DEFAULT_ACCOUNT_MODE);
  }
  for (const input of meetingModeInputs) {
    input.checked = input.value === (data.meetingMode || DEFAULT_MEETING_MODE);
  }
  updateModeState();

  const redirectEl = document.getElementById("redirectUri");
  try {
    redirectEl.textContent = browser.identity.getRedirectURL();
  } catch (err) {
    redirectEl.textContent = "Unavailable";
  }

}

function showRedirectStatus(text) {
  redirectStatusEl.textContent = text;
  if (!text) {
    return;
  }
  setTimeout(() => {
    if (redirectStatusEl.textContent === text) {
      redirectStatusEl.textContent = "";
    }
  }, 1500);
}

async function copyRedirectUri() {
  const redirectEl = document.getElementById("redirectUri");
  const value = redirectEl.textContent || "";
  if (!value || value === "Unavailable") {
    showRedirectStatus("No redirect URI available.");
    return;
  }
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(value);
    } else {
      const range = document.createRange();
      range.selectNodeContents(redirectEl);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      document.execCommand("copy");
      selection.removeAllRanges();
    }
    showRedirectStatus("Copied.");
  } catch (err) {
    showRedirectStatus("Copy failed.");
  }
}

function setWarning(text) {
  if (!text) {
    accountWarningEl.textContent = "";
    accountWarningEl.classList.add("hidden");
    return;
  }
  accountWarningEl.textContent = text;
  accountWarningEl.classList.remove("hidden");
}

function setPersonalAccountNote(visible) {
  if (visible) {
    personalAccountNoteEl.classList.remove("hidden");
  } else {
    personalAccountNoteEl.classList.add("hidden");
  }
}

function setConfigStatus(text) {
  configStatusEl.textContent = text;
}

function updateLocalStatus() {
  const clientIdValue = document.getElementById("clientId").value.trim();
  if (isPlaceholder(clientIdValue)) {
    setConfigStatus("Setup status: missing Application ID");
  } else {
    setConfigStatus("Setup status: ready");
  }
}

async function refreshStatus() {
  try {
    const status = await browser.runtime.sendMessage({ type: "getStatus" });
    if (!status) {
      setConfigStatus("Setup status: unknown");
      accountStatusEl.textContent = "Signed in as: unknown";
      setWarning("");
      setPersonalAccountNote(false);
      return;
    }
    if (!status.configured) {
      setConfigStatus("Setup status: missing Application ID");
      accountStatusEl.textContent = "Signed in as: unknown";
      setWarning("");
      setPersonalAccountNote(false);
      return;
    }
    setConfigStatus("Setup status: ready");
    if (status.accountEmail || status.accountName) {
      const label = status.accountEmail || status.accountName;
      accountStatusEl.textContent = `Signed in as: ${label}`;
    } else {
      accountStatusEl.textContent = "Signed in as: unknown";
    }
    if (status.accountType === "personal" || status.accountMode === "personal") {
      setWarning("Warning: personal Microsoft accounts may not support Teams meeting creation.");
      setPersonalAccountNote(true);
    } else {
      setWarning("");
      setPersonalAccountNote(false);
    }
  } catch (err) {
    setConfigStatus("Setup status: unknown");
    accountStatusEl.textContent = "Signed in as: unknown";
    setWarning("");
    setPersonalAccountNote(false);
  }
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  statusEl.textContent = "Saving...";

  const previousSettings = await browser.storage.local.get({
    accountMode: DEFAULT_ACCOUNT_MODE,
    meetingMode: DEFAULT_MEETING_MODE
  });

  const accountMode = getSelectedAccountMode();
  const meetingMode = getSelectedMeetingMode();

  const payload = {
    clientId: document.getElementById("clientId").value.trim(),
    tenant: document.getElementById("tenant").value.trim() || DEFAULT_TENANT,
    authorityHost: document.getElementById("authorityHost").value.trim() || DEFAULT_AUTHORITY_HOST,
    debugEnabled: document.getElementById("debugEnabled").checked,
    allowCustomAuthorityHost: document.getElementById("allowCustomAuthorityHost").checked,
    accountMode,
    meetingMode,
    useDefaultApplicationId: false,
    scopes: getScopesForMode(meetingMode)
  };

  if (isPlaceholder(payload.clientId)) {
    statusEl.textContent = "Application ID is required.";
    return;
  }

  payload.useDefaultApplicationId = false;

  const validation = validateSettings(payload);
  if (!validation.ok) {
    statusEl.textContent = validation.errors.join(" ");
    return;
  }
  if (validation.normalized.authorityHost) {
    payload.authorityHost = validation.normalized.authorityHost;
  }
  if (validation.normalized.tenant) {
    payload.tenant = validation.normalized.tenant;
  }

  await browser.storage.local.set(payload);
  statusEl.textContent = validation.warnings.length
    ? `Saved (warning: ${validation.warnings.join(" ")})`
    : "Saved";

  if (previousSettings.accountMode !== accountMode || previousSettings.meetingMode !== meetingMode) {
    try {
      await browser.runtime.sendMessage({ type: "logout" });
      statusEl.textContent = "Saved (re-auth required).";
    } catch (err) {
      statusEl.textContent = "Saved (could not clear login).";
    }
  }
  setTimeout(() => {
    statusEl.textContent = "";
  }, 1200);

  refreshStatus();
});

testButton.addEventListener("click", async () => {
  testStatusEl.textContent = "Testing...";
  testButton.disabled = true;

  try {
    const settings = await browser.storage.local.get({
      clientId: "",
      tenant: DEFAULT_TENANT,
      authorityHost: DEFAULT_AUTHORITY_HOST,
      scopes: DEFAULT_SCOPES.join(" "),
      debugEnabled: false,
      allowCustomAuthorityHost: false,
      accountMode: DEFAULT_ACCOUNT_MODE,
      meetingMode: DEFAULT_MEETING_MODE,
      useDefaultApplicationId: false
    });
    const requiresCustomAppId = settings.accountMode === "work";
    if (requiresCustomAppId && isPlaceholder(settings.clientId)) {
      testStatusEl.textContent = "Set Application ID first.";
      return;
    }
    if (isPlaceholder(settings.clientId)) {
      testStatusEl.textContent = "Set Application ID first.";
      return;
    }
    const validation = validateSettings(settings);
    if (!validation.ok) {
      testStatusEl.textContent = validation.errors.join(" ");
      return;
    }
    if (validation.warnings.length) {
      testStatusEl.textContent = `Testing... (${validation.warnings.join(" ")})`;
    }
    const result = await browser.runtime.sendMessage({ type: "testConnection" });
    if (result?.ok) {
      testStatusEl.textContent = "Connection successful.";
      if (result.accountEmail || result.accountName) {
        const label = result.accountEmail || result.accountName;
        accountStatusEl.textContent = `Signed in as: ${label}`;
      }
      if (result.accountType === "personal") {
        setWarning("Warning: personal Microsoft accounts may not support Teams meeting creation.");
      } else {
        setWarning("");
      }
    } else {
      testStatusEl.textContent = "Connection failed.";
    }
  } catch (err) {
    testStatusEl.textContent = "Connection failed.";
  } finally {
    testButton.disabled = false;
  }
});

copyRedirectButton.addEventListener("click", () => {
  copyRedirectUri();
});

for (const input of accountModeInputs) {
  input.addEventListener("change", () => {
    updateModeState();
    updateLocalStatus();
  });
}

for (const input of meetingModeInputs) {
  input.addEventListener("change", () => {
    updateModeState();
    updateLocalStatus();
  });
}

document.getElementById("tenant").addEventListener("input", (event) => {
  if (!event.target.disabled) {
    lastWorkTenant = event.target.value.trim() || DEFAULT_TENANT;
  }
});

document.getElementById("clientId").addEventListener("input", () => {
  updateLocalStatus();
});

logoutButton.addEventListener("click", async () => {
  statusEl.textContent = "Signing out...";
  try {
    await browser.runtime.sendMessage({ type: "logout" });
    statusEl.textContent = "Signed out.";
  } catch (err) {
    statusEl.textContent = "Sign out failed.";
  } finally {
    setTimeout(() => {
      statusEl.textContent = "";
    }, 1200);
    refreshStatus();
  }
});

loadSettings().then(() => {
  updateLocalStatus();
  refreshStatus();
});
