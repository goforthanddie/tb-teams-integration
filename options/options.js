/* global browser, DEFAULT_APPLICATION_ID, DEFAULT_TENANT, DEFAULT_AUTHORITY_HOST, DEFAULT_SCOPES, isPlaceholder, validateSettings */

const form = document.getElementById("settings-form");
const statusEl = document.getElementById("status");
const testButton = document.getElementById("testConnection");
const testStatusEl = document.getElementById("testStatus");
const configStatusEl = document.getElementById("configStatus");
const accountWarningEl = document.getElementById("accountWarning");
const copyRedirectButton = document.getElementById("copyRedirect");
const redirectStatusEl = document.getElementById("redirectStatus");

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

loadSettings().then(refreshStatus);
