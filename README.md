# TB Teams Integration

Adds a **Create Teams meeting** button to the Thunderbird event dialog. When clicked, the add-on authenticates the user with Microsoft (OAuth 2.0 + PKCE), creates an online meeting via Microsoft Graph, and injects the join link into the event description (and location if empty).

## Setup (Microsoft Entra ID)

1. Create a new App Registration in Microsoft Entra ID.
2. Note the **Application (client) ID**.
3. Add a redirect URI that matches the add-on's redirect URI shown in the Options page.
   - The options page shows a value like `https://<extension-id>.extensions.mozilla.org/`.
4. Enable the following **Microsoft Graph** delegated permissions:
   - `OnlineMeetings.ReadWrite`
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
- Personal Microsoft accounts can sign in if your app allows them, but Teams meeting creation may not be supported.

## Development

Key files:
- `manifest.json`
- `background.js`
- `experiments/teamsDialog/parent.js`
- `experiments/teamsDialog/schema.json`
- `options/options.html`
