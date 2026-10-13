# Changelog

## Unreleased

- Password reset and Temporary Access Pass require their own opt-in flags
  (`MS_ENABLE_PASSWORD_RESET`, `MS_ENABLE_TAP`) in addition to `MS_ENABLE_WRITES`.
- Password reset, TAP and group-membership operations refuse privileged targets
  (directory role holders, members of role-assignable groups, role-assignable
  groups, `MS_PROTECTED_USERS` / `MS_PROTECTED_GROUPS`) and fail closed when the
  check cannot be performed.
- Identifiers used in Graph URL paths are validated (no empty values, dot
  segments, separators or percent-encoded separators).
- Absolute `@odata.nextLink` URLs are only followed on the configured Graph origin.
- `MCP_HTTP_TOKEN` is compared in constant time.
- Replaced two backtracking regular expressions with linear-time code.
- Updated vulnerable production dependencies.

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
