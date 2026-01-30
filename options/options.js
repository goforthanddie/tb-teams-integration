/* global browser */

const form = document.getElementById("settings-form");
const statusEl = document.getElementById("status");
const testButton = document.getElementById("testConnection");
const testStatusEl = document.getElementById("testStatus");
const configStatusEl = document.getElementById("configStatus");
const accountWarningEl = document.getElementById("accountWarning");

function isPlaceholder(value) {
  if (!value) {
    return true;
  }
  return value === "REPLACE_WITH_CLIENT_ID";
}

async function loadSettings() {
  const data = await browser.storage.local.get({
    clientId: "REPLACE_WITH_CLIENT_ID",
    tenant: "organizations",
    authorityHost: "https://login.microsoftonline.com",
    scopes: "OnlineMeetings.ReadWrite offline_access openid profile",
    debugEnabled: false
  });

  document.getElementById("clientId").value = data.clientId;
  document.getElementById("tenant").value = data.tenant;
  document.getElementById("authorityHost").value = data.authorityHost;
  document.getElementById("scopes").value = data.scopes;
  document.getElementById("debugEnabled").checked = !!data.debugEnabled;

  const redirectEl = document.getElementById("redirectUri");
  try {
    redirectEl.textContent = browser.identity.getRedirectURL();
  } catch (err) {
    redirectEl.textContent = "Unavailable";
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

async function refreshStatus() {
  try {
    const status = await browser.runtime.sendMessage({ type: "getStatus" });
    if (!status) {
      configStatusEl.textContent = "Setup status: unknown";
      setWarning("");
      return;
    }
    if (!status.configured) {
      configStatusEl.textContent = "Setup status: missing Application ID";
      setWarning("");
      return;
    }
    configStatusEl.textContent = "Setup status: ready";
    if (status.accountType === "personal") {
      setWarning("Warning: personal Microsoft accounts may not support Teams meeting creation.");
    } else {
      setWarning("");
    }
  } catch (err) {
    configStatusEl.textContent = "Setup status: unknown";
    setWarning("");
  }
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  statusEl.textContent = "Saving...";

  const payload = {
    clientId: document.getElementById("clientId").value.trim(),
    tenant: document.getElementById("tenant").value.trim() || "organizations",
    authorityHost: document.getElementById("authorityHost").value.trim() || "https://login.microsoftonline.com",
    scopes: document.getElementById("scopes").value.trim(),
    debugEnabled: document.getElementById("debugEnabled").checked
  };

  await browser.storage.local.set(payload);
  statusEl.textContent = "Saved";
  setTimeout(() => {
    statusEl.textContent = "";
  }, 1200);

  refreshStatus();
});

testButton.addEventListener("click", async () => {
  testStatusEl.textContent = "Testing...";
  testButton.disabled = true;

  try {
    const settings = await browser.storage.local.get({ clientId: "" });
    if (isPlaceholder(settings.clientId)) {
      testStatusEl.textContent = "Set Application ID first.";
      return;
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

loadSettings().then(refreshStatus);
