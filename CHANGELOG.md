# Changelog

## 0.1.0

- Initial Microsoft 365 / Microsoft Graph (Entra ID) MCP server.
- Added stdio and Streamable HTTP transports.
- Added app-only (client-credentials) Graph client with in-memory token
  caching, safe base URL checks, and redacted errors.
- Added read tools for users, groups, group members, subscribed SKUs, and
  license details.
- Added write tools for user creation, license assignment/removal, group
  membership, and usage location, disabled by default.
- Added friendly license-name resolution against subscribed SKUs with
  availability checks.
- Added the `./gateway` export consumed by the Borgels gateway.
