/* global */

const DEFAULT_APPLICATION_ID = "REPLACE_WITH_APPLICATION_ID";
const DEFAULT_TENANT = "organizations";
const DEFAULT_AUTHORITY_HOST = "https://login.microsoftonline.com";
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
  return value === DEFAULT_APPLICATION_ID;
}
