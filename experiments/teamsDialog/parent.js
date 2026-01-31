/* global ChromeUtils */

const { ExtensionCommon } = ChromeUtils.importESModule(
  "resource://gre/modules/ExtensionCommon.sys.mjs"
);

let Services = null;
try {
  ({ Services } = ChromeUtils.importESModule("resource://gre/modules/Services.sys.mjs"));
} catch (err) {
  try {
    ({ Services } = ChromeUtils.importESModule("resource:///modules/Services.sys.mjs"));
  } catch (err2) {
    Services = globalThis.Services || null;
  }
}

const dialogWindows = new Map();
let nextDialogId = 1;
let eventFire = null;
let windowListenerRegistered = false;
let teamsIconUrl = null;
let windowListener = null;
let debugEnabled = false;
const OBSERVER_TIMEOUT_MS = 10000;

function log(message) {
  if (!debugEnabled) {
    return;
  }
  const text = `[tb-teams] ${message}`;
  try {
    Services?.console?.logStringMessage(text);
  } catch (err) {
    // Ignore logging errors.
  }
  try {
    console.log(text);
  } catch (err) {
    // Ignore console errors.
  }
}

function isEventDialog(win) {
  try {
    const doc = win?.document;
    if (!doc) {
      return false;
    }
    const windowType = doc.documentElement?.getAttribute("windowtype");
    if (windowType === "Calendar:EventDialog") {
      return true;
    }
    const href = doc.location?.href || "";
    return href.includes("calendar-event-dialog.xhtml");
  } catch (err) {
    return false;
  }
}

function describeWindow(win) {
  try {
    const doc = win?.document;
    const windowType = doc?.documentElement?.getAttribute("windowtype") || "unknown";
    const href = doc?.location?.href || "unknown";
    return `windowtype=${windowType} href=${href}`;
  } catch (err) {
    return "windowtype=unknown href=unknown";
  }
}

function getDialogDoc(win) {
  const doc = win.document;
  const iframe = doc.getElementById("calendar-item-panel-iframe");
  if (iframe?.contentDocument) {
    return iframe.contentDocument;
  }
  return doc;
}

function findToolbar(doc) {
  if (!doc) {
    return null;
  }
  return (
    doc.getElementById("event-toolbar") ||
    doc.getElementById("event-toolbarbutton-bar") ||
    doc.querySelector("toolbar#event-toolbar, toolbar#event-toolbarbutton-bar")
  );
}

function readDateTime(picker) {
  if (!picker) {
    return null;
  }

  if (picker.dateValue) {
    if (picker.dateValue instanceof Date) {
      return picker.dateValue;
    }
    if (picker.dateValue.jsDate instanceof Date) {
      return picker.dateValue.jsDate;
    }
    if (typeof picker.dateValue.toISOString === "function") {
      const iso = picker.dateValue.toISOString();
      const parsed = new Date(iso);
      if (!Number.isNaN(parsed.getTime())) {
        return parsed;
      }
    }
  }

  const raw = picker.value || picker.getAttribute("value") || "";
  if (!raw) {
    return null;
  }
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }
  return parsed;
}

function formatDateTimeWithOffset(date) {
  const pad2 = value => String(value).padStart(2, "0");
  const year = date.getFullYear();
  const month = pad2(date.getMonth() + 1);
  const day = pad2(date.getDate());
  const hours = pad2(date.getHours());
  const minutes = pad2(date.getMinutes());
  const seconds = pad2(date.getSeconds());
  const offsetMinutes = date.getTimezoneOffset();
  const offsetTotal = Math.abs(offsetMinutes);
  const offsetHours = pad2(Math.floor(offsetTotal / 60));
  const offsetMins = pad2(offsetTotal % 60);
  const sign = offsetMinutes <= 0 ? "+" : "-";
  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}${sign}${offsetHours}:${offsetMins}`;
}

function normalizeAllDayTimes(start, end) {
  if (!start) {
    return { start, end };
  }
  const normalizedStart = new Date(start);
  normalizedStart.setHours(0, 0, 0, 0);
  let normalizedEnd = end ? new Date(end) : null;
  if (normalizedEnd) {
    normalizedEnd.setHours(0, 0, 0, 0);
  } else {
    normalizedEnd = new Date(normalizedStart);
    normalizedEnd.setDate(normalizedEnd.getDate() + 1);
  }
  if (normalizedEnd <= normalizedStart) {
    normalizedEnd = new Date(normalizedStart);
    normalizedEnd.setDate(normalizedEnd.getDate() + 1);
  }
  return { start: normalizedStart, end: normalizedEnd };
}

function buildPayload(win) {
  const doc = getDialogDoc(win);
  const title = doc.getElementById("item-title")?.value || "";
  const location = doc.getElementById("item-location")?.value || "";
  const allDay = !!doc.getElementById("event-all-day")?.checked;
  const startPicker = doc.getElementById("event-starttime");
  const endPicker = doc.getElementById("event-endtime");
  let start = readDateTime(startPicker);
  let end = readDateTime(endPicker);
  if (allDay) {
    const normalized = normalizeAllDayTimes(start, end);
    start = normalized.start;
    end = normalized.end;
  }

  let descriptionText = "";
  const editor = doc.getElementById("item-description");
  if (editor?.contentDocument?.body) {
    descriptionText = editor.contentDocument.body.textContent || "";
  }

  return {
    title,
    location,
    isAllDay: allDay,
    startDateTime: start ? formatDateTimeWithOffset(start) : "",
    endDateTime: end ? formatDateTimeWithOffset(end) : "",
    descriptionText
  };
}

function ensureDialogId(win) {
  if (!win.__teamsDialogId) {
    win.__teamsDialogId = nextDialogId++;
    dialogWindows.set(win.__teamsDialogId, win);
  }
  return win.__teamsDialogId;
}

function createTeamsButtonInDoc(doc, win, dialogId) {
  const toolbar = findToolbar(doc);
  if (!toolbar) {
    log("Toolbar not found in document.");
    return false;
  }

  if (doc.getElementById("tb-teams-create-button")) {
    return true;
  }

  const buttonFactory = doc.createXULElement ? doc.createXULElement.bind(doc) : doc.createElement.bind(doc);
  const button = buttonFactory("toolbarbutton");
  button.setAttribute("id", "tb-teams-create-button");
  button.setAttribute("class", "toolbarbutton-1");
  button.setAttribute("label", "Teams");
  button.setAttribute("tooltiptext", "Create Teams meeting");
  if (teamsIconUrl) {
    button.setAttribute("image", teamsIconUrl);
  }

  button.addEventListener("command", () => {
    if (!eventFire) {
      return;
    }
    const payload = buildPayload(win);
    eventFire.async({ dialogId, ...payload });
  });

  toolbar.appendChild(button);
  log("Teams button inserted.");
  return true;
}

function clearObserver(win) {
  if (win?.__teamsObserver) {
    win.__teamsObserver.disconnect();
    win.__teamsObserver = null;
  }
  if (win?.__teamsObserverTimer) {
    win.clearTimeout(win.__teamsObserverTimer);
    win.__teamsObserverTimer = null;
  }
}

function createTeamsButton(win) {
  const dialogId = ensureDialogId(win);
  const doc = win.document;
  const primaryDoc = getDialogDoc(win);

  clearObserver(win);
  log(`Attempting button injection. docHref=${doc?.location?.href || "unknown"} primaryDocHref=${primaryDoc?.location?.href || "unknown"}`);
  if (createTeamsButtonInDoc(doc, win, dialogId)) {
    return;
  }
  if (primaryDoc !== doc && createTeamsButtonInDoc(primaryDoc, win, dialogId)) {
    return;
  }

  const ObserverClass = win.MutationObserver || win.WebKitMutationObserver;
  if (!ObserverClass) {
    log("MutationObserver not available; skipping delayed injection.");
    return;
  }

  const observer = new ObserverClass(() => {
    if (createTeamsButtonInDoc(doc, win, dialogId)) {
      clearObserver(win);
      return;
    }
    if (primaryDoc !== doc && createTeamsButtonInDoc(primaryDoc, win, dialogId)) {
      clearObserver(win);
    }
  });

  observer.observe(doc.documentElement, { childList: true, subtree: true });
  if (primaryDoc !== doc) {
    observer.observe(primaryDoc.documentElement, { childList: true, subtree: true });
  }
  win.__teamsObserver = observer;
  win.__teamsObserverTimer = win.setTimeout(() => {
    log("MutationObserver timeout; stopping delayed injection.");
    clearObserver(win);
  }, OBSERVER_TIMEOUT_MS);
}

function ensureDialog(win) {
  if (!isEventDialog(win)) {
    return;
  }

  const readyState = win?.document?.readyState || "unknown";
  const windowType = win?.document?.documentElement?.getAttribute("windowtype") || "unknown";
  if (readyState === "complete" || readyState === "interactive") {
    log(`Event dialog ready (state=${readyState}, windowtype=${windowType}), attempting to inject Teams button.`);
    createTeamsButton(win);
    return;
  }

  win.addEventListener(
    "load",
    () => {
      const loadedWindowType = win?.document?.documentElement?.getAttribute("windowtype") || "unknown";
      log(`Event dialog loaded (windowtype=${loadedWindowType}), attempting to inject Teams button.`);
      createTeamsButton(win);
    },
    { once: true }
  );
}

function attachToExistingWindows() {
  log("Scanning existing windows.");
  const enumerator = Services.wm.getEnumerator(null);
  while (enumerator.hasMoreElements()) {
    const win = enumerator.getNext();
    log(`Seen window: ${describeWindow(win)}`);
    if (isEventDialog(win)) {
      log("Existing event dialog found, attempting to inject Teams button.");
      createTeamsButton(win);
    }
  }
}

function registerWindowListener() {
  if (windowListenerRegistered) {
    return;
  }
  windowListenerRegistered = true;

  if (!Services?.wm) {
    log("Services.wm unavailable; window listener not registered.");
    return;
  }

  windowListener = {
    onOpenWindow(xulWindow) {
      const win = xulWindow.docShell.domWindow;
      log(`Window opened: ${describeWindow(win)}`);
      win.addEventListener(
        "load",
        () => {
          log(`Window load fired: ${describeWindow(win)}`);
          ensureDialog(win);
        },
        { once: true }
      );
    },
    onCloseWindow(xulWindow) {
      const win = xulWindow.docShell.domWindow;
      if (win?.__teamsDialogId) {
        dialogWindows.delete(win.__teamsDialogId);
      }
      clearObserver(win);
    },
    onWindowTitleChange() {}
  };

  Services.wm.addListener(windowListener);
}

function insertJoinInfoInDialog(win, joinUrl, label) {
  const doc = getDialogDoc(win);
  const editor = doc.getElementById("item-description");
  if (editor?.contentDocument?.body) {
    const body = editor.contentDocument.body;
    const text = body.textContent || "";
    if (!text.includes(joinUrl)) {
      const p = editor.contentDocument.createElement("p");
      p.textContent = `${label || "Teams meeting"}: ${joinUrl}`;
      body.appendChild(p);
    }
  }

  const location = doc.getElementById("item-location");
  if (location && !location.value) {
    location.value = joinUrl;
  }
}

function showError(win, message) {
  try {
    Services.prompt.alert(win, "Teams meeting", message);
  } catch (err) {
    Services.prompt.alert(null, "Teams meeting", message);
  }
}

const TeamsDialogAPI = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    teamsIconUrl = context.extension.baseURI.resolve("icons/teams.svg");
    log("teamsDialog API initialized.");
    return {
      teamsDialog: {
        register: () => {
          log("teamsDialog.register called.");
          registerWindowListener();
          attachToExistingWindows();
        },
        setDebug: ({ enabled }) => {
          debugEnabled = Boolean(enabled);
          log(`Debug logging ${debugEnabled ? "enabled" : "disabled"}.`);
        },
        insertJoinInfo: ({ dialogId, joinUrl, label }) => {
          const win = dialogWindows.get(dialogId);
          if (win) {
            insertJoinInfoInDialog(win, joinUrl, label);
          }
        },
        showError: ({ dialogId, message }) => {
          const win = dialogWindows.get(dialogId);
          showError(win, message);
        },
        onTeamsButtonClick: new ExtensionCommon.EventManager({
          context,
          name: "teamsDialog.onTeamsButtonClick",
          register(fire) {
            eventFire = fire;
            return () => {
              eventFire = null;
            };
          }
        }).api()
      }
    };
  }

  onShutdown() {
    eventFire = null;
    if (windowListenerRegistered) {
      Services.wm.removeListener(windowListener);
      windowListener = null;
      windowListenerRegistered = false;
    }
  }
};

var teamsDialog = TeamsDialogAPI;
