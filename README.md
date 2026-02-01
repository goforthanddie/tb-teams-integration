# TB Teams Integration

Adds a **Create Teams meeting** button to the Thunderbird event dialog. When clicked, the add-on authenticates the user with Microsoft (OAuth 2.0 + PKCE), creates an online meeting via Microsoft Graph, and injects the join link into the event description (and location if empty).
All accounts require an app registration in Microsoft Entra. Personal Microsoft accounts can use calendar scheduling when the app allows consumer sign-in.

## Setup (Microsoft Entra ID)

1. Create a new App Registration in Microsoft Entra ID.
2. Note the **Application (client) ID**.
3. Add a redirect URI that matches the add-on's redirect URI shown in the Options page.
   - The options page shows a value like `https://<extension-id>.extensions.mozilla.org/`.
4. Enable the following **Microsoft Graph** delegated permissions:
   - `OnlineMeetings.ReadWrite`
   - `Calendars.ReadWrite`
   - `openid`, `profile`, `offline_access`
5. Grant admin consent for the tenant (recommended for corporate environments).

> Note: Microsoft Graph online meetings created with `/me/onlineMeetings` are not automatically stored as Exchange calendar events. This add-on inserts the join URL into the local Thunderbird event instead.

## Thunderbird Setup

1. Open Thunderbird.
2. Go to **Tools -> Add-ons and Themes -> Extensions**.
3. Click the gear icon and choose **Debug Add-ons**.
4. Click **Load Temporary Add-on** and select `manifest.json` from this repo.
5. Open **Add-on Options** and fill in:
   - Application ID
   - Tenant (use your tenant GUID or domain; default `organizations`)
   - Authority host (default `https://login.microsoftonline.com`)
6. Use **Test connection** to verify sign-in.
7. Create a new calendar event and click the **Teams** button in the toolbar.

## Notes

- This add-on uses delegated OAuth; each user must sign in the first time they click the button.
- If the user is already signed in to Microsoft in the auth window, SSO will reduce friction.
- Tokens are stored in Thunderbird extension storage and refreshed automatically.
- Direct meetings use `/me/onlineMeetings` and require a work or school account with Teams/Exchange licensing.
- Calendar scheduling uses `/me/events` with `isOnlineMeeting` and works for personal accounts only if the app allows consumer sign-in.
- App registration (Application ID) is required for all account types.

## Meeting Modes

- **Direct meeting**: Creates a Teams meeting via `/me/onlineMeetings` and inserts the join link into the Thunderbird event.
- **Calendar schedule**: Creates a calendar event via `/me/events` with `isOnlineMeeting=true`. The join link may not be returned for some personal accounts.

## Enterprise Deployment

You can preconfigure default values by editing `shared/constants.js` before packaging the add-on. If you replace `DEFAULT_APPLICATION_ID`, `DEFAULT_TENANT`, or `DEFAULT_AUTHORITY_HOST` with your organization defaults, the options page will be pre-filled and the add-on will only prompt users when the Application ID is still the placeholder value. Rebuild the XPI after making these changes.


## Security and Data Handling

This add-on requests the minimum permissions needed for OAuth (`identity`) and local settings storage (`storage`). It shows an “unrestricted access” warning because it uses a Thunderbird Experiment API to add the Teams button to the event dialog; experiments are treated as privileged even if the declared permissions are minimal.

What it can access:
- Calendar event data visible in the event dialog (title, start/end time, location, description) to create the meeting and insert the join URL.

What is stored locally:
- OAuth access and refresh tokens and their expiration time in Thunderbird’s extension storage.

Where data is sent:
- Microsoft login endpoints for OAuth.
- Microsoft Graph (`/me/onlineMeetings` and `/me/events`) to create the meeting.

The add-on does not transmit data to any other servers.

## Development

Key files:
- `manifest.json`
- `background.js`
- `experiments/teamsDialog/parent.js`
- `experiments/teamsDialog/schema.json`
- `options/options.html`
- `options/options.js`
- `options/options.css`
- `shared/constants.js`
- `shared/validation.js`
- `scripts/build-xpi.sh`
- `scripts/test.js`

Build:
- `scripts/build-xpi.sh` (outputs `dist/tb-teams-integration.xpi`)

Tests:
- `node scripts/test.js`
