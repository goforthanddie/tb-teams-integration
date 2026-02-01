/* global browser, DEFAULT_APPLICATION_ID, DEFAULT_TENANT, DEFAULT_AUTHORITY_HOST, DEFAULT_SCOPES, isPlaceholder, validateSettings */

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

async function loadSettings() {
  const data = await browser.storage.local.get({
    clientId: DEFAULT_APPLICATION_ID,
    tenant: DEFAULT_TENANT,
    authorityHost: DEFAULT_AUTHORITY_HOST,
    scopes: DEFAULT_SCOPES.join(" "),
    debugEnabled: false,
    allowCustomAuthorityHost: false
  });

  const clientIdEl = document.getElementById("clientId");
  const tenantEl = document.getElementById("tenant");
  const authorityHostEl = document.getElementById("authorityHost");
  clientIdEl.value = data.clientId;
  tenantEl.value = data.tenant;
  authorityHostEl.value = data.authorityHost;

  clientIdEl.placeholder = DEFAULT_APPLICATION_ID;
  tenantEl.placeholder = DEFAULT_TENANT;
  authorityHostEl.placeholder = DEFAULT_AUTHORITY_HOST;
  document.getElementById("debugEnabled").checked = !!data.debugEnabled;
  document.getElementById("allowCustomAuthorityHost").checked = !!data.allowCustomAuthorityHost;

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

async function refreshStatus() {
  try {
    const status = await browser.runtime.sendMessage({ type: "getStatus" });
    if (!status) {
      configStatusEl.textContent = "Setup status: unknown";
      accountStatusEl.textContent = "Signed in as: unknown";
      setWarning("");
      setPersonalAccountNote(false);
      return;
    }
    if (!status.configured) {
      configStatusEl.textContent = "Setup status: missing Application ID";
      accountStatusEl.textContent = "Signed in as: unknown";
      setWarning("");
      setPersonalAccountNote(false);
      return;
    }
    configStatusEl.textContent = "Setup status: ready";
    if (status.accountEmail || status.accountName) {
      const label = status.accountEmail || status.accountName;
      accountStatusEl.textContent = `Signed in as: ${label}`;
    } else {
      accountStatusEl.textContent = "Signed in as: unknown";
    }
    if (status.accountType === "personal") {
      setWarning("Warning: personal Microsoft accounts may not support Teams meeting creation.");
      setPersonalAccountNote(true);
    } else {
      setWarning("");
      setPersonalAccountNote(false);
    }
  } catch (err) {
    configStatusEl.textContent = "Setup status: unknown";
    accountStatusEl.textContent = "Signed in as: unknown";
    setWarning("");
    setPersonalAccountNote(false);
  }
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  statusEl.textContent = "Saving...";

  const payload = {
    clientId: document.getElementById("clientId").value.trim(),
    tenant: document.getElementById("tenant").value.trim() || DEFAULT_TENANT,
    authorityHost: document.getElementById("authorityHost").value.trim() || DEFAULT_AUTHORITY_HOST,
    debugEnabled: document.getElementById("debugEnabled").checked,
    allowCustomAuthorityHost: document.getElementById("allowCustomAuthorityHost").checked
  };

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
      clientId: DEFAULT_APPLICATION_ID,
      tenant: DEFAULT_TENANT,
      authorityHost: DEFAULT_AUTHORITY_HOST,
      scopes: DEFAULT_SCOPES.join(" "),
      debugEnabled: false,
      allowCustomAuthorityHost: false
    });
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

loadSettings().then(refreshStatus);
