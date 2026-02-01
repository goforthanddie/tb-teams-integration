/* global */

const APPLICATION_ID_PLACEHOLDER = "REPLACE_WITH_APPLICATION_ID";
const DEFAULT_APPLICATION_ID = APPLICATION_ID_PLACEHOLDER;
const DEFAULT_TENANT = "organizations";
const DEFAULT_AUTHORITY_HOST = "https://login.microsoftonline.com";
const DEFAULT_ACCOUNT_MODE = "work";
const DEFAULT_MEETING_MODE = "direct";
const DEFAULT_SCOPES = [
  "OnlineMeetings.ReadWrite",
  "Calendars.ReadWrite",
  "offline_access",
  "openid",
  "profile"
];

function isPlaceholder(value) {
  if (!value) {
    return true;
  }
  return value === APPLICATION_ID_PLACEHOLDER;
}

function resolveDefaultApplicationId() {
  return DEFAULT_APPLICATION_ID;
}
