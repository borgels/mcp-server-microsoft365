# Changelog

## 0.2.0

### Upgrade notes

- The HTTP server refuses to start on a non-loopback bind (`MCP_HTTP_HOST`/`HOST`,
  e.g. `0.0.0.0`) or with `MCP_ALLOW_ANY_ORIGIN=true` unless `MCP_HTTP_TOKEN` is set.
  `MCP_HTTP_ALLOW_UNAUTHENTICATED=true` restores the old behaviour explicitly.
- `MCP_HTTP_TOKEN` must be at least 32 printable ASCII characters.
- `PORT`, `MCP_MAX_BODY_BYTES`, `MS_TIMEOUT_MS`, `MCP_ALLOW_ANY_ORIGIN` and the other
  numeric/boolean settings are validated at startup; invalid values stop the server
  instead of being read as `NaN`.
- The `Bearer` scheme in `Authorization` is now matched case-insensitively.

### Changes

- HTTP transport: failed-authentication limiting per client address (`429` with
  `Retry-After`; new `MCP_AUTH_FAILURE_LIMIT`, `MCP_AUTH_FAILURE_WINDOW_MS`), `Host`
  header check on loopback binds (new `MCP_ALLOWED_HOSTS`), `401` with
  `WWW-Authenticate`, `400` for invalid JSON, `413` before reading an oversized body,
  `GET /healthz`, graceful shutdown on `SIGTERM`/`SIGINT`.
- Microsoft Graph and token requests no longer follow redirects, so the bearer token
  and client secret cannot be forwarded to another host.
- Upstream error bodies are read up to 64 KB; error text is redacted before it is
  truncated.
- The default PIM activation justification names this server.
- Docker image is published only after CI passes, with SBOM and provenance; GitHub
  Actions are pinned to commit SHAs; dependency updates.

### Earlier unreleased changes

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
- Added the `./gateway` export for embedding the tools in another host.
