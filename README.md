# mcp-server-microsoft365

TypeScript MCP server for Microsoft 365 / Entra ID via Microsoft Graph. It runs
app-only (client-credentials) and is intentionally boring good: typed,
documented, read-first, least-privilege, credential-sane, and idempotent where
Graph allows it. Its focus is identity provisioning — create a user, assign a
license, manage group membership — plus the supporting reads you need to do that
safely.

> **Disclaimer:** This is an independent, unofficial project by Borgels. Borgels
> is not affiliated with, endorsed by, or supported by Microsoft. "Microsoft
> 365", "Microsoft Graph", and "Entra ID" are referenced only to describe what
> this server talks to. You need your own Entra ID app registration, and use of
> Microsoft Graph is subject to Microsoft's own terms.

## Scope

Read tools (enabled by default):

- `get_user`, `list_users`
- `list_subscribed_skus` (license lookup and seat availability)
- `list_groups`, `get_group`, `list_group_members`
- `get_user_license_details`

Write tools (disabled by default):

- `create_user`
- `assign_license`, `remove_license`
- `add_group_member`, `remove_group_member`
- `set_usage_location`

The server runs application permissions only; it never signs a user in and never
accepts credentials as tool arguments.

## Least-privilege application permissions

Grant the app registration only what it needs, then admin-consent:

| Task | Application permission |
| --- | --- |
| Read users and groups | `User.Read.All` |
| Create users | `User.Create` (or `User.ReadWrite.All`) |
| Read license/SKU state | `LicenseAssignment.Read.All` |
| Assign / remove licenses | `LicenseAssignment.ReadWrite.All` |
| Manage group membership | `GroupMember.ReadWrite.All` |

`User.Create` is narrower than `User.ReadWrite.All`; use it when you only need to
provision new users. Reading SKUs (`/subscribedSkus`) is covered by
`LicenseAssignment.Read.All`.

## App registration and admin consent

1. In the Entra admin center, register an application (single tenant).
2. Under **Certificates & secrets**, create a client secret and copy its value.
3. Under **API permissions**, add the Microsoft Graph *Application* permissions
   above, then click **Grant admin consent**.
4. Copy the **Directory (tenant) ID** and **Application (client) ID**.

## Setup

Install dependencies and build the CLI:

```sh
npm install
npm run build
```

Provide credentials through the environment. The server never accepts them as
tool arguments.

```sh
export MS_TENANT_ID="contoso.onmicrosoft.com"   # or the tenant GUID
export MS_CLIENT_ID="your-application-client-id"
export MS_CLIENT_SECRET="your-client-secret"
```

Optional settings:

```sh
export MS_GRAPH_BASE_URL="https://graph.microsoft.com/v1.0"
export MS_TIMEOUT_MS=30000
# export MS_AUTHORITY_HOST="https://login.microsoftonline.com"
```

If another process manages tokens, provide a bearer token directly:

```sh
export MS_ACCESS_TOKEN="your-access-token"
```

## Claude or Cursor config

Use the stdio server for local MCP clients:

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "node",
      "args": ["/absolute/path/to/mcp-server-microsoft365/dist/transports/stdio.js"],
      "env": {
        "MS_TENANT_ID": "contoso.onmicrosoft.com",
        "MS_CLIENT_ID": "your-application-client-id",
        "MS_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

During development:

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "npm",
      "args": ["run", "dev", "--prefix", "/absolute/path/to/mcp-server-microsoft365"],
      "env": {
        "MS_TENANT_ID": "contoso.onmicrosoft.com",
        "MS_CLIENT_ID": "your-application-client-id",
        "MS_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

## Streamable HTTP transport

```sh
npm run dev:http
```

The HTTP transport binds to `127.0.0.1` by default and serves `POST /mcp`. Set
`MCP_HTTP_TOKEN` to require bearer auth, and `MCP_ALLOWED_ORIGINS` for
browser-based local clients. `MCP_MAX_BODY_BYTES` defaults to `10485760`
(10 MiB).

## Provisioning workflow

1. `list_subscribed_skus` to see which licenses exist and how many seats are
   free.
2. `create_user` with a `displayName` and `userPrincipalName`. A strong
   temporary password is generated when you do not supply one, and returned in
   the result with `forceChangePasswordNextSignIn` set. Store it securely and
   share it out-of-band. `create_user` fails clearly if the UPN already exists.
3. `assign_license` with a friendly name (`E3`), a `skuPartNumber`
   (`ENTERPRISEPACK`), or a SKU GUID. Pass `usageLocation` on the same call for a
   brand-new user.
4. `add_group_member` to place the user in the right groups. It is idempotent:
   an already-a-member response is treated as success.

### The usageLocation gotcha

Microsoft Graph rejects a license assignment for a user with no `usageLocation`
(the classic "License assignment failed because of an invalid usage location"
error). This server sets `usageLocation` **first**:

- If you pass `usageLocation` to `assign_license`, it is `PATCH`ed onto the user
  before the license is assigned.
- If you do not, the server reads the user's current `usageLocation` and returns
  a clear, actionable error when it is missing — rather than letting Graph fail
  cryptically.

`usageLocation` is an ISO 3166-1 alpha-2 country code, e.g. `DK`.

### License resolution

`assign_license` / `remove_license` accept a friendly name, a `skuPartNumber`,
or a SKU GUID. The server fetches `/subscribedSkus` (no `$filter`; it matches
client-side) and:

- matches a GUID directly against `skuId`;
- otherwise matches `skuPartNumber` case-insensitively;
- otherwise maps a friendly name (e.g. `E3` → `ENTERPRISEPACK` / `SPE_E3`,
  `E5` → `ENTERPRISEPREMIUM` / `SPE_E5`) to the first candidate the tenant owns;
- checks availability (`prepaidUnits.enabled - consumedUnits`) and refuses to
  assign when no seats are free.

## Rate limiting

On HTTP 429, the raised `GraphHttpError` carries the `Retry-After` value and the
Graph `request-id`. The Borgels connector layer uses these to retry.

## Borgels gateway

The `./gateway` subpath exports the Borgels gateway contract:

```ts
import { createMicrosoft365Gateway, microsoft365GatewayTools } from 'mcp-server-microsoft365/gateway';

const gateway = createMicrosoft365Gateway({ tenantId, clientId, clientSecret });
const result = await gateway.callTool('list_subscribed_skus');
```

`microsoft365GatewayTools` lists the same tool surface as the MCP server, with
`riskLevel` and `enabledByDefault` so the gateway can gate writes. Read tools are
`enabledByDefault: true`; every write tool is `enabledByDefault: false`.

## Live smoke test

The normal test suite uses mocked `fetch`. With real credentials in the
environment you can run a read-only live smoke test that counts subscribed SKUs:

```sh
npm run smoke:live
```

## Security and audit

- Credentials are read from environment variables only.
- `MS_GRAPH_BASE_URL` must be `https://`; loopback `http://` is allowed for local
  mocks.
- Formatted errors redact Authorization headers, client secrets, access tokens,
  and token-like material.
- Write tools are disabled by default and surfaced with write/destructive
  annotations.
- The Streamable HTTP transport binds to `127.0.0.1` by default.

Security reports: security@borgels.com.
