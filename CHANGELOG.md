# Changelog

## 0.6.1

- Every change to an existing user addresses the object the protection guard
  checked, by the object id the directory gave, also when the tool was given a
  login name: profile, usage location, manager, licences, group membership,
  password reset and Temporary Access Passes. A login name in a group's member
  reference or in a removal by member is not accepted by Graph.
- `remove_group_member` removes by object id, and reads a 404 as "already
  removed" only when the group's member list shows that the user is not a
  direct member; if the user is still in the list, or the list cannot be
  read, the error stands. Before, a login name gave 404 and the tool reported
  success while the user kept the membership.
- `add_group_member` references the user's object id, not the login name.
- `microsoft365ToolCatalog`: `remove_group_member` lists the member list read.
- Tests: the host-mode test of a failed protection lookup turns password
  reset and Temporary Access Passes on, so it is the lookup that stops every
  tool, and it checks that the lookup was tried; the test that an ordinary
  user goes through no longer swallows errors.
- `RegisterToolsOptions.protection` describes the guard as it is from 0.6.0:
  before every tool that changes an existing user.

## 0.6.0

- The privileged-target guard runs before every change to an existing user,
  not only before password reset, Temporary Access Passes and group
  membership: `update_user`, `set_usage_location`, `set_manager` (the user
  whose manager changes), `assign_license`, `remove_license` and
  `delete_temporary_access_pass` now refuse a user who holds a directory role,
  has a role assignment, is in a role-assignable group or is listed in
  `MS_PROTECTED_USERS`, and refuse when that cannot be verified (the check
  fails closed). Nothing is sent to Graph in either case. These tools now need
  `RoleManagement.Read.Directory` as well. In host mode the refusal is the
  fixed `Refused: the target is protected.`
- `microsoft365ToolCatalog` lists the guard's lookups for those tools.
- Tests run every tool that changes an existing user against an administrator,
  a scoped role assignment, a refused lookup and a lookup that does not get
  through.

## 0.5.1

- Licensor and copyright holder corrected to Borgels Services ApS in `LICENSE`,
  `NOTICE` and `CONTRIBUTING.md`. The license is unchanged (FSL-1.1-ALv2).

## 0.5.0

- License changed from Apache-2.0 to the Functional Source License, Version 1.1,
  ALv2 Future License (FSL-1.1-ALv2). You may read, run, modify and use the code
  internally and contribute changes; offering it as a competing commercial
  product or service is not allowed. Each version becomes Apache-2.0 two years
  after its release. Versions released before 0.5.0 remain under Apache-2.0
  (see `NOTICE`).

## 0.4.0

- Host-mode gateway error results carry failure metadata: when Microsoft
  Graph answered with an HTTP error or a redirect (never followed), the
  result has `failure: { httpStatus, requestId? }`, so a host can count
  failures by status without reading their text. `requestId` is Graph's own
  id from a response header (`REQUEST_ID_HEADERS`: `request-id`,
  `client-request-id`, `x-ms-request-id` first), kept only when it is a hex
  id of 8 to 64 characters (one run of hex digits, or groups of four or more
  joined by single hyphens, as in a UUID), does not hold the shape of a
  personal number and shares no piece of six characters with the token. The
  text is unchanged. `GraphHttpError.requestId` now holds only an id in that
  form (the message keeps the header's value as before), errors from the
  token endpoint keep its request id too, and `failureMetadata` is exported
  from `./gateway` for an environment-mode instance, whose errors are thrown.
- A refusal because a protection lookup failed at Graph (`Refused: the
  target is protected.`) carries that failure's metadata, and an HTTP error
  whose body cannot be read is still raised with its status and request id
  (the body is left out) instead of being lost to the failed read.
- `microsoft365ToolCatalog`: `create_user` lists no PATCH, and password
  reset and temporary access passes list no group lookup (they do neither).
  The catalog test runs every tool to success against realistic Graph
  answers, so a failure can no longer hide a request.
- `package-lock.json`: the root package's version is 0.3.0.
- Development: tsup 9, vitest 5 and @types/node 26 (build and tests only;
  the published package is unchanged).
- `engines.node` is `>=22.18`: building from source (`prepare`) needs
  it for tsup 9, and Node 20 is no longer supported.

## 0.3.0

### Changes

- `./gateway`: host mode. `createMicrosoft365Gateway({ accessToken, policy })`
  with `policy: { enableWrites?, disabledTools? }` reads nothing from the
  environment: Graph at its own address, the host's access token (the host
  fetches it for the tenant, for example with a certificate), and the
  protection policy only from `protection`, otherwise everything refused.
  Write tools are refused unless `enableWrites` is true; disabled tools are
  left out of `tools`. Input is checked against the tool's schema (known
  fields, required fields, types) before any request. Errors are fixed texts
  with only the Graph HTTP status, and a result that holds the token is
  withheld. Without `policy` the gateway behaves as before.
- `./gateway`: `microsoft365ToolCatalog`, a deeply frozen catalog of every
  gateway tool with the Graph requests it may make (`upstream`) and the
  `broad` mark (false for all), held by a test against the requests each tool
  makes.
- Host mode sends the token to `https://graph.microsoft.com/v1.0` only;
  another `baseUrl` is refused. Validation errors are fixed texts (no field
  names or values) and enforce the schema's minimum and maximum.
  `get_user_license_details` is not offered in host mode: Graph does not
  support application permissions for it.

## 0.2.0

### Upgrade notes

- The HTTP server refuses to start on a non-loopback bind (`MCP_HTTP_HOST`/`HOST`,
  e.g. `0.0.0.0`) or with `MCP_ALLOW_ANY_ORIGIN=true` unless `MCP_HTTP_TOKEN` is set.
  `MCP_HTTP_ALLOW_UNAUTHENTICATED=true` restores the old behaviour explicitly.
- `MCP_HTTP_TOKEN` must be at least 32 printable ASCII characters.
- These settings are validated at startup and an invalid value stops the server
  instead of being read as `NaN` or ignored: integers `PORT` (1-65535),
  `MCP_MAX_BODY_BYTES` (>= 1), `MS_TIMEOUT_MS` (1-600000),
  `MCP_AUTH_FAILURE_LIMIT` (>= 0), `MCP_AUTH_FAILURE_WINDOW_MS` (>= 1); booleans
  `MCP_ALLOW_ANY_ORIGIN` and `MCP_HTTP_ALLOW_UNAUTHENTICATED` (`true/false`, `1/0`,
  `yes/no`, `on/off`); host names in `MCP_ALLOWED_HOSTS`. `MS_ENABLE_*` flags are
  unchanged and still need exactly `true`.
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
