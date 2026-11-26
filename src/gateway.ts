import {
  GRAPH_DEFAULT_AUTHORITY_HOST,
  GRAPH_DEFAULT_BASE_URL,
  GRAPH_DEFAULT_SCOPE,
  GRAPH_DEFAULT_TIMEOUT_MS,
  GraphClient,
  type GraphClientOptions,
  type GraphQueryValue,
} from './graph/client.js';
import { failureFrom, failureMetadata, GraphHttpError, type FailureMetadata } from './errors.js';

export { failureMetadata } from './errors.js';
import {
  addGroupMember,
  assignLicense,
  createTemporaryAccessPass,
  createUser,
  deleteTemporaryAccessPass,
  listSubscribedSkus,
  pimSelfActivateRole,
  removeGroupMember,
  removeLicense,
  resetPassword,
  setManager,
  setUsageLocation,
  updateUser,
  type UpdateUserInput,
} from './graph/operations.js';
import { pathSegment } from './graph/path.js';
import * as personal from './graph/personal.js';
import { RefusedError } from './graph/personal.js';
import { ProtectedTargetError, protectionPolicyFromEnv, type ProtectionPolicy } from './graph/protection.js';
import { PERSONAL_TOOL_NAMES, personalTools } from './personal-tools.js';

export {
  MICROSOFT365_DELEGATED_PERMISSIONS,
  READ_BY_POST,
  permissionCovers,
  type DelegatedPermission,
} from './personal-tools.js';
export { REFUSAL_TEXT, type RefusalCode } from './graph/personal.js';

export type GatewayRiskLevel = 'read' | 'write' | 'destructive';
export type GatewayJsonValue = string | number | boolean | null | GatewayJsonValue[] | { [key: string]: GatewayJsonValue };
export type GatewayJsonObject = { [key: string]: GatewayJsonValue };

export interface GatewayToolDefinition {
  name: string;
  title: string;
  description: string;
  riskLevel: GatewayRiskLevel;
  enabledByDefault: boolean;
  inputSchema: GatewayJsonObject;
}

/**
 * Metadata about a failed call, on an error result whose cause is known: the HTTP status Microsoft Graph answered
 * with, and its request id when a response header gives one as a UUID or hex id. For a host that counts failures
 * without reading their text. Never the response body, Graph's error code or message, or the input.
 */
export type GatewayFailure = FailureMetadata;

export interface GatewayToolResult {
  content: Array<{ type: 'text'; text: string }>;
  structuredContent?: GatewayJsonValue;
  isError?: boolean;
  /** On an error result from Microsoft Graph only (host mode): its HTTP status and request id (see GatewayFailure). */
  failure?: GatewayFailure;
}

export interface Microsoft365GatewayOptions extends GraphClientOptions {
  /**
   * Guardrails for reset_password, Temporary Access Pass and group membership.
   * Defaults to the environment (MS_ENABLE_PASSWORD_RESET, MS_ENABLE_TAP,
   * MS_PROTECTED_USERS, MS_PROTECTED_GROUPS); password reset and TAP are refused
   * unless explicitly enabled.
   */
  protection?: Partial<ProtectionPolicy>;
  /**
   * Instance policy from the embedding host. When given, the instance is in
   * host mode: it reads nothing from the environment (no MS_* variable, not
   * even for the protection policy, which then defaults to everything refused),
   * it needs `accessToken` (the host fetches the access for its tenant), write
   * tools are refused unless `enableWrites` is true, and `disabledTools` are
   * left out of `tools` and refused. Errors are fixed texts that carry only the
   * Graph HTTP status, never Graph's own text. A host serving several tenants
   * should always pass it.
   */
  policy?: Microsoft365InstancePolicy;
}

/**
 * Which access the host's token is:
 *   - "application": an app-only access to a tenant (client credentials). The directory tools (users, groups,
 *     licences) are offered.
 *   - "delegated": one person's own access (authorization code, then refresh), for their own mail, calendar and files.
 *     Only the personal tools are offered, and Microsoft Graph keeps them to what that person can reach.
 */
export type Microsoft365Access = 'application' | 'delegated';

/** Policy given by an embedding host for one gateway instance. */
export interface Microsoft365InstancePolicy {
  /** The kind of access the token is, and so which tools are offered. Default "application". */
  access?: Microsoft365Access;
  /** Allow write and destructive tools. Default false. */
  enableWrites?: boolean;
  /** Tools turned off on this instance. Must name gateway tools. */
  disabledTools?: readonly string[];
}

export class GatewayConfigError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'GatewayConfigError';
  }
}

/** One Microsoft Graph request a tool may make: the method and the path, with `{...}` for a value from the input. */
export interface GatewayUpstreamOperation {
  readonly method: 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE';
  readonly path: string;
}

/** A tool catalog entry: the gateway definition plus the Graph requests the tool may make. */
export interface GatewayToolCatalogEntry {
  readonly name: string;
  readonly title: string;
  readonly description: string;
  readonly riskLevel: GatewayRiskLevel;
  readonly enabledByDefault: boolean;
  readonly inputSchema: { readonly [key: string]: unknown };
  readonly upstream: readonly GatewayUpstreamOperation[];
  /**
   * True when the tool can read more than its name says (anything in the
   * directory, chosen by the input). No gateway tool is broad: every read
   * names what it reads (users, groups, licences, the person's own mail,
   * calendar or files).
   */
  readonly broad: boolean;
  /** The access the tool is made for: "application" (the directory tools) or "delegated" (the person's own data). */
  readonly access: Microsoft365Access;
  /** For a delegated tool: the delegated Graph permissions its requests need (the least Microsoft documents), main one first. */
  readonly permissions?: readonly string[];
}

const userRef = { type: 'string', description: 'User object id (GUID) or userPrincipalName.' } satisfies GatewayJsonObject;
const groupRef = { type: 'string', description: 'Group object id (GUID).' } satisfies GatewayJsonObject;
const licenseRef = {
  type: 'string',
  description: 'License as a friendly name (e.g. "E3"), skuPartNumber (e.g. "ENTERPRISEPACK"), or SKU GUID.',
} satisfies GatewayJsonObject;
const readListProps = {
  top: { type: 'number', minimum: 1, maximum: 999 },
  filter: { type: 'string', description: 'OData $filter expression.' },
  select: { type: 'array', items: { type: 'string' } },
  orderby: { type: 'string', description: 'OData $orderby expression.' },
} satisfies Record<string, GatewayJsonObject>;

/** The directory tools (users, groups, licences), made for an app-only access. */
const directoryTools: GatewayToolDefinition[] = [
  {
    name: 'get_user',
    title: 'Get Microsoft 365 user',
    description: 'Fetch one user by object id or userPrincipalName.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: {
      type: 'object',
      required: ['user'],
      properties: { user: userRef, select: { type: 'array', items: { type: 'string' } } },
      additionalProperties: false,
    },
  },
  {
    name: 'list_users',
    title: 'List Microsoft 365 users',
    description: 'List users with optional $top, $filter, $select, and $orderby.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: { type: 'object', properties: { ...readListProps }, additionalProperties: false },
  },
  {
    name: 'list_subscribed_skus',
    title: 'List subscribed SKUs',
    description: 'List tenant subscribed SKUs for license lookup and seat availability.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: { type: 'object', properties: {}, additionalProperties: false },
  },
  {
    name: 'list_groups',
    title: 'List Microsoft 365 groups',
    description: 'List groups with optional $top, $filter, $select, and $orderby.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: { type: 'object', properties: { ...readListProps }, additionalProperties: false },
  },
  {
    name: 'get_group',
    title: 'Get Microsoft 365 group',
    description: 'Fetch one group by object id.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: {
      type: 'object',
      required: ['groupId'],
      properties: { groupId: groupRef, select: { type: 'array', items: { type: 'string' } } },
      additionalProperties: false,
    },
  },
  {
    name: 'list_group_members',
    title: 'List group members',
    description: 'List the members of a group.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: {
      type: 'object',
      required: ['groupId'],
      properties: { groupId: groupRef, top: { type: 'number', minimum: 1, maximum: 999 } },
      additionalProperties: false,
    },
  },
  {
    name: 'get_user_license_details',
    title: 'Get user license details',
    description: 'List the licenses currently assigned to a user.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: { type: 'object', required: ['user'], properties: { user: userRef }, additionalProperties: false },
  },
  {
    name: 'create_user',
    title: 'Create Microsoft 365 user',
    description: 'Create a user with a generated temporary password. Returns the temporary password.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: {
      type: 'object',
      required: ['displayName', 'userPrincipalName'],
      properties: {
        displayName: { type: 'string' },
        userPrincipalName: { type: 'string' },
        mailNickname: { type: 'string' },
        password: { type: 'string', description: 'Optional. A strong temporary password is generated when omitted.' },
        usageLocation: { type: 'string', description: 'ISO 3166-1 alpha-2 country code, e.g. "DK".' },
        accountEnabled: { type: 'boolean' },
        forceChangePasswordNextSignIn: { type: 'boolean' },
        checkExisting: { type: 'boolean', description: 'Defaults to true: fail if the UPN already exists.' },
      },
      additionalProperties: false,
    },
  },
  {
    name: 'assign_license',
    title: 'Assign license to user',
    description: 'Assign a license, setting usageLocation first when required.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: {
      type: 'object',
      required: ['user', 'license'],
      properties: {
        user: userRef,
        license: licenseRef,
        usageLocation: { type: 'string', description: 'ISO 3166-1 alpha-2 country code, set before assignment.' },
        disabledPlans: { type: 'array', items: { type: 'string' } },
      },
      additionalProperties: false,
    },
  },
  {
    name: 'remove_license',
    title: 'Remove license from user',
    description: 'Remove a license from a user.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: {
      type: 'object',
      required: ['user', 'license'],
      properties: { user: userRef, license: licenseRef },
      additionalProperties: false,
    },
  },
  {
    name: 'add_group_member',
    title: 'Add group member',
    description: 'Add a user to a group. Idempotent when already a member.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: {
      type: 'object',
      required: ['groupId', 'userId'],
      properties: { groupId: groupRef, userId: userRef },
      additionalProperties: false,
    },
  },
  {
    name: 'remove_group_member',
    title: 'Remove group member',
    description: 'Remove a user from a group. Idempotent when not a member.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: {
      type: 'object',
      required: ['groupId', 'userId'],
      properties: { groupId: groupRef, userId: userRef },
      additionalProperties: false,
    },
  },
  {
    name: 'set_usage_location',
    title: 'Set user usage location',
    description: 'Set a user usageLocation (required before license assignment).',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: {
      type: 'object',
      required: ['user', 'usageLocation'],
      properties: { user: userRef, usageLocation: { type: 'string', description: 'ISO 3166-1 alpha-2 country code.' } },
      additionalProperties: false,
    },
  },
  {
    name: 'update_user',
    title: 'Update Microsoft 365 user',
    description:
      'Update attributes on an existing user (PATCH). Use accountEnabled to enable/disable an existing account. Only provided fields change.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: {
      type: 'object',
      required: ['user'],
      properties: {
        user: userRef,
        accountEnabled: { type: 'boolean' },
        displayName: { type: 'string' },
        givenName: { type: 'string' },
        surname: { type: 'string' },
        jobTitle: { type: 'string' },
        department: { type: 'string' },
        companyName: { type: 'string' },
        employeeType: { type: 'string' },
        mobilePhone: { type: 'string' },
        streetAddress: { type: 'string' },
        city: { type: 'string' },
        postalCode: { type: 'string' },
        state: { type: 'string' },
        country: { type: 'string' },
        officeLocation: { type: 'string' },
        usageLocation: { type: 'string', description: 'ISO 3166-1 alpha-2 country code.' },
        otherMails: { type: 'array', items: { type: 'string' } },
        businessPhones: { type: 'array', items: { type: 'string' } },
      },
      additionalProperties: false,
    },
  },
  {
    name: 'reset_password',
    title: 'Reset user password',
    description:
      'Give an existing user a new temporary password, which they must change at next sign-in. Returns it once.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: {
      type: 'object',
      required: ['user'],
      properties: { user: userRef },
      additionalProperties: false,
    },
  },
  {
    name: 'set_manager',
    title: 'Set user manager',
    description: "Set a user's manager by object id or userPrincipalName.",
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: {
      type: 'object',
      required: ['user', 'manager'],
      properties: { user: userRef, manager: userRef },
      additionalProperties: false,
    },
  },
  {
    name: 'create_temporary_access_pass',
    title: 'Create Temporary Access Pass',
    description:
      'Create a Temporary Access Pass (multi-use by default), regenerated until alphanumeric. lifetimeInMinutes is bounded by the tenant TAP policy. Returns the passcode; deliver out-of-band.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: {
      type: 'object',
      required: ['user'],
      properties: {
        user: userRef,
        lifetimeInMinutes: { type: 'number', minimum: 10, maximum: 43200 },
        isUsableOnce: { type: 'boolean' },
        startDateTime: { type: 'string', description: 'ISO 8601 activation time (future-dating supported).' },
        requireAlphanumeric: { type: 'boolean', description: 'Default true.' },
        maxAttempts: { type: 'number', minimum: 1, maximum: 20 },
      },
      additionalProperties: false,
    },
  },
  {
    name: 'delete_temporary_access_pass',
    title: 'Delete Temporary Access Pass',
    description: "Delete a user's Temporary Access Pass. With methodId deletes that pass; otherwise deletes all.",
    riskLevel: 'destructive',
    enabledByDefault: false,
    inputSchema: {
      type: 'object',
      required: ['user'],
      properties: { user: userRef, methodId: { type: 'string' } },
      additionalProperties: false,
    },
  },
  {
    name: 'activate_pim_role',
    title: 'Activate PIM role (self)',
    description:
      'Just-in-time PIM elevation: the delegated caller self-activates one of their own eligible directory roles for a bounded window (e.g. Authentication Administrator for TAP). Requires a delegated token; fails if not eligible.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: {
      type: 'object',
      required: ['roleDefinitionId'],
      properties: {
        roleDefinitionId: { type: 'string', description: 'Role definition/template GUID to activate.' },
        duration: { type: 'string', description: 'ISO 8601 duration of the activation window (default PT1H).' },
        justification: { type: 'string' },
        directoryScopeId: { type: 'string', description: "Directory scope; '/' (tenant) by default." },
      },
      additionalProperties: false,
    },
  },
];

/** Every gateway tool: the directory tools, then the personal tools (personal-tools.ts). */
export const microsoft365GatewayTools: GatewayToolDefinition[] = [
  ...directoryTools,
  ...personalTools.map(({ upstream: _upstream, permissions: _permissions, ...definition }) => structuredClone(definition) as GatewayToolDefinition),
];

/** The protection guard's lookups for a user (assertUserNotPrivileged). */
const USER_GUARD_LOOKUPS: readonly GatewayUpstreamOperation[] = [
  { method: 'GET', path: '/users/{user}' },
  { method: 'GET', path: '/users/{user}/transitiveMemberOf' },
  { method: 'GET', path: '/roleManagement/directory/roleAssignments' },
];

/** The guard's lookups for a group membership change: the group (assertGroupNotProtected) and the user. */
const GROUP_GUARD_LOOKUPS: readonly GatewayUpstreamOperation[] = [{ method: 'GET', path: '/groups/{group}' }, ...USER_GUARD_LOOKUPS];

/** The Graph requests each gateway tool may make (the guard's directory lookups included where it runs). */
const UPSTREAM: Readonly<Record<string, readonly GatewayUpstreamOperation[]>> = {
  get_user: [{ method: 'GET', path: '/users/{user}' }],
  list_users: [{ method: 'GET', path: '/users' }],
  list_subscribed_skus: [{ method: 'GET', path: '/subscribedSkus' }],
  list_groups: [{ method: 'GET', path: '/groups' }],
  get_group: [{ method: 'GET', path: '/groups/{group}' }],
  list_group_members: [{ method: 'GET', path: '/groups/{group}/members' }],
  get_user_license_details: [{ method: 'GET', path: '/users/{user}/licenseDetails' }],
  create_user: [
    { method: 'GET', path: '/users/{user}' },
    { method: 'POST', path: '/users' },
  ],
  assign_license: [
    ...USER_GUARD_LOOKUPS,
    { method: 'GET', path: '/subscribedSkus' },
    { method: 'GET', path: '/users/{user}' },
    { method: 'PATCH', path: '/users/{user}' },
    { method: 'POST', path: '/users/{user}/assignLicense' },
  ],
  remove_license: [
    ...USER_GUARD_LOOKUPS,
    { method: 'GET', path: '/subscribedSkus' },
    { method: 'POST', path: '/users/{user}/assignLicense' },
  ],
  add_group_member: [...GROUP_GUARD_LOOKUPS, { method: 'POST', path: '/groups/{group}/members/$ref' }],
  remove_group_member: [
    ...GROUP_GUARD_LOOKUPS,
    { method: 'DELETE', path: '/groups/{group}/members/{user}/$ref' },
    // Only after a 404: whether the user is in the group's member list, before it reads as already removed.
    { method: 'GET', path: '/groups/{group}/members' },
  ],
  set_usage_location: [...USER_GUARD_LOOKUPS, { method: 'PATCH', path: '/users/{user}' }],
  update_user: [...USER_GUARD_LOOKUPS, { method: 'PATCH', path: '/users/{user}' }],
  reset_password: [...USER_GUARD_LOOKUPS, { method: 'PATCH', path: '/users/{user}' }],
  set_manager: [
    ...USER_GUARD_LOOKUPS,
    { method: 'GET', path: '/users/{manager}' },
    { method: 'PUT', path: '/users/{user}/manager/$ref' },
  ],
  create_temporary_access_pass: [
    ...USER_GUARD_LOOKUPS,
    { method: 'POST', path: '/users/{user}/authentication/temporaryAccessPassMethods' },
    { method: 'DELETE', path: '/users/{user}/authentication/temporaryAccessPassMethods/{method}' },
  ],
  delete_temporary_access_pass: [
    ...USER_GUARD_LOOKUPS,
    { method: 'GET', path: '/users/{user}/authentication/temporaryAccessPassMethods' },
    { method: 'DELETE', path: '/users/{user}/authentication/temporaryAccessPassMethods/{method}' },
  ],
  activate_pim_role: [
    { method: 'GET', path: '/me' },
    { method: 'POST', path: '/roleManagement/directory/roleAssignmentScheduleRequests' },
  ],
};

function deepFreeze<T>(value: T): T {
  if (value && typeof value === 'object') {
    for (const entry of Object.values(value as Record<string, unknown>)) deepFreeze(entry);
    Object.freeze(value);
  }
  return value;
}

/**
 * The tool catalog: every gateway tool with its definition, the Graph
 * requests it may make and the broad mark, deeply frozen. Hosts can build
 * their own permission model from it without creating an instance.
 */
export const microsoft365ToolCatalog: readonly GatewayToolCatalogEntry[] = deepFreeze([
  ...directoryTools.map(tool => {
    const upstream = UPSTREAM[tool.name];
    if (!upstream) throw new Error(`No upstream list for the gateway tool ${tool.name}`);
    return {
      ...structuredClone(tool),
      upstream: upstream.map(operation => ({ ...operation })),
      broad: false,
      access: 'application' as const,
    };
  }),
  ...personalTools.map(({ upstream, permissions, ...tool }) => ({
    ...(structuredClone(tool) as GatewayToolDefinition),
    upstream: upstream.map(operation => ({ ...operation })),
    broad: false,
    access: 'delegated' as const,
    permissions: [...permissions],
  })),
]);

/** The kind of access each tool is offered for in host mode. */
const ACCESS_OF: ReadonlyMap<string, Microsoft365Access> = new Map(microsoft365ToolCatalog.map(entry => [entry.name, entry.access]));

const TOOL_NAMES: ReadonlySet<string> = new Set(microsoft365GatewayTools.map(tool => tool.name));

/**
 * Create one gateway instance. With `policy` it is in host mode (see
 * Microsoft365GatewayOptions.policy); without, it reads the environment as
 * the MCP server does.
 */
export function createMicrosoft365Gateway(options: Microsoft365GatewayOptions = {}) {
  if (options.policy !== undefined) return createHostGateway(options, options.policy);
  return createEnvGateway(options);
}

/** The only address a host-mode instance sends its token to. */
const HOST_MODE_BASE_URL = GRAPH_DEFAULT_BASE_URL;

/**
 * Tools a host-mode instance never offers: Microsoft Graph does not support application permissions for them, so an
 * access fetched with client credentials always fails
 * (https://learn.microsoft.com/en-us/graph/api/user-list-licensedetails, "Application: Not supported").
 */
export const HOST_MODE_UNSUPPORTED: readonly string[] = Object.freeze(['get_user_license_details']);

/** A bearer token as the host gives it: printable ASCII without spaces, of a sensible length. */
const BEARER = /^[\x21-\x7e]{16,16384}$/;

function createHostGateway(options: Microsoft365GatewayOptions, hostPolicy: Microsoft365InstancePolicy) {
  const accessToken: unknown = options.accessToken;
  if (typeof accessToken !== 'string' || !BEARER.test(accessToken)) {
    throw new GatewayConfigError(
      'A host-mode Microsoft 365 gateway requires options.accessToken. Host-mode instances read nothing from the environment.',
    );
  }
  // The token goes to Microsoft Graph and nowhere else: no other host, port, scheme or version.
  if (options.baseUrl !== undefined && options.baseUrl !== HOST_MODE_BASE_URL) {
    throw new GatewayConfigError(`A host-mode Microsoft 365 gateway only talks to ${HOST_MODE_BASE_URL}.`);
  }
  const access: unknown = hostPolicy.access ?? 'application';
  if (access !== 'application' && access !== 'delegated') throw new GatewayConfigError('policy.access must be "application" or "delegated".');
  const disabled = new Set<string>(access === 'application' ? HOST_MODE_UNSUPPORTED : []);
  for (const name of hostPolicy.disabledTools ?? []) {
    if (!TOOL_NAMES.has(name)) throw new GatewayConfigError(`policy.disabledTools names no gateway tool: ${String(name).slice(0, 100)}`);
    disabled.add(name);
  }
  const enableWrites = hostPolicy.enableWrites === true;
  // Every client option explicit: `??` in GraphClient falls back to the environment only on undefined.
  const inner = createEnvGateway(
    {
      tenantId: '',
      clientId: '',
      clientSecret: '',
      refreshToken: '',
      accessToken,
      baseUrl: HOST_MODE_BASE_URL,
      authorityHost: GRAPH_DEFAULT_AUTHORITY_HOST,
      scope: GRAPH_DEFAULT_SCOPE,
      timeoutMs: options.timeoutMs ?? GRAPH_DEFAULT_TIMEOUT_MS,
      ...(options.fetchImpl ? { fetchImpl: options.fetchImpl } : {}),
    },
    {
      allowPasswordReset: false,
      allowTap: false,
      protectedUsers: [],
      protectedGroups: [],
      ...definedOnly(options.protection ?? {}),
    },
  );
  const offered = (tool: GatewayToolDefinition) =>
    ACCESS_OF.get(tool.name) === access && !disabled.has(tool.name) && (tool.riskLevel === 'read' || enableWrites);
  const tools = microsoft365GatewayTools.filter(offered);
  const byName = new Map(tools.map(tool => [tool.name, tool]));

  return {
    tools,
    async callTool(toolName: string, input: GatewayJsonObject = {}): Promise<GatewayToolResult> {
      const tool = typeof toolName === 'string' ? byName.get(toolName) : undefined;
      if (!tool) return errorResult('Refused: the tool is not offered on this instance.');
      // A fixed code and text: nothing of the input (a field name, a value) is repeated.
      const invalid = invalidInput(tool, input);
      if (invalid) return withoutToken(errorResult(`Invalid input for ${tool.name}: ${INVALID_TEXT[invalid]}`), accessToken);
      try {
        return withoutToken(await inner.callTool(tool.name, input), accessToken);
      } catch (error) {
        if (error instanceof GraphHttpError) {
          // The status and request id as metadata, checked against the token like the rest of the result, then
          // checked again for their form as the last step.
          const failed = errorResult(`Microsoft Graph request failed with HTTP ${error.status}`, failureWithoutToken(error, accessToken));
          return withFailureInForm(withoutToken(failed, accessToken));
        }
        if (error instanceof ProtectedTargetError) {
          // A refusal because a directory lookup failed at Graph keeps that HTTP failure's metadata.
          const refused = errorResult('Refused: the target is protected.', failureWithoutToken(error, accessToken));
          return withFailureInForm(withoutToken(refused, accessToken));
        }
        // A personal tool refused the change before anything was sent: a fixed text chosen by its code.
        if (error instanceof RefusedError) return errorResult(personal.REFUSAL_TEXT[error.code]);
        return errorResult('The Microsoft 365 call could not be completed.');
      }
    },
  };
}

function definedOnly(value: Partial<ProtectionPolicy>): Partial<ProtectionPolicy> {
  return Object.fromEntries(Object.entries(value).filter(([, entry]) => entry !== undefined)) as Partial<ProtectionPolicy>;
}

type InvalidInput = 'not_object' | 'unknown_field' | 'wrong_type' | 'out_of_range' | 'missing_field';

/** The fixed text for each kind of invalid input. */
const INVALID_TEXT: Readonly<Record<InvalidInput, string>> = {
  not_object: 'the input must be an object.',
  unknown_field: 'a field is not in the schema.',
  wrong_type: 'a field has the wrong type.',
  out_of_range: 'a number is outside its range.',
  missing_field: 'a required field is missing.',
};

/** Checks the input against the tool's schema: an object, known fields only, types, ranges, required fields. */
function invalidInput(tool: GatewayToolDefinition, input: unknown): InvalidInput | undefined {
  if (typeof input !== 'object' || input === null || Array.isArray(input)) return 'not_object';
  const properties = (tool.inputSchema.properties ?? {}) as Record<string, { type?: string; minimum?: number; maximum?: number }>;
  for (const [key, value] of Object.entries(input)) {
    const schema = Object.prototype.hasOwnProperty.call(properties, key) ? properties[key] : undefined;
    if (!schema) return 'unknown_field';
    const type = schema.type;
    const ok =
      type === 'string'
        ? typeof value === 'string'
        : type === 'number'
          ? typeof value === 'number' && Number.isFinite(value)
          : type === 'boolean'
            ? typeof value === 'boolean'
            : type === 'array'
              ? Array.isArray(value) && value.every(item => typeof item === 'string')
              : true;
    if (!ok) return 'wrong_type';
    if (typeof value === 'number') {
      if (typeof schema.minimum === 'number' && value < schema.minimum) return 'out_of_range';
      if (typeof schema.maximum === 'number' && value > schema.maximum) return 'out_of_range';
    }
  }
  for (const key of (tool.inputSchema.required ?? []) as string[]) {
    if (!Object.prototype.hasOwnProperty.call(input, key)) return 'missing_field';
  }
  return undefined;
}

/** Pieces of this many characters of the token may not appear in a request id. */
const TOKEN_PIECE = 6;

/** The failure metadata of an error (a Graph HTTP error, or a refusal caused by one), without a request id that shares a piece of TOKEN_PIECE characters with the token. */
function failureWithoutToken(error: unknown, token: string): GatewayFailure | undefined {
  const failure = error instanceof ProtectedTargetError ? failureFrom(error.status, error.requestId) : failureMetadata(error);
  const id = failure?.requestId;
  if (failure === undefined || id === undefined) return failure;
  for (let i = 0; i + TOKEN_PIECE <= id.length; i += 1) {
    if (token.includes(id.slice(i, i + TOKEN_PIECE))) return { httpStatus: failure.httpStatus };
  }
  return failure;
}

/**
 * The failure metadata as it must be when the result leaves: on an error result only, a status of 100-599 and a
 * request id in the form of one. Anything else is dropped, never passed on.
 */
function withFailureInForm(result: GatewayToolResult): GatewayToolResult {
  const { failure, ...rest } = result;
  if (!result.isError || failure === undefined || failure === null || typeof failure !== 'object') return rest;
  const checked = failureFrom(failure.httpStatus, failure.requestId);
  return checked === undefined ? rest : { ...rest, failure: checked };
}

/** A result that holds the instance's token anywhere (it never should) is withheld whole. */
function withoutToken(result: GatewayToolResult, token: string): GatewayToolResult {
  return JSON.stringify(result).includes(token) ? errorResult('The answer was withheld.') : result;
}

function createEnvGateway(options: Microsoft365GatewayOptions, fixedProtection?: ProtectionPolicy) {
  const { protection, policy: _policy, ...clientOptions } = options;
  const client = new GraphClient(clientOptions);
  const policy: ProtectionPolicy = fixedProtection ?? { ...protectionPolicyFromEnv() };
  if (!fixedProtection) {
    for (const [key, value] of Object.entries(protection ?? {})) {
      if (value !== undefined) (policy as unknown as Record<string, unknown>)[key] = value;
    }
  }

  return {
    tools: microsoft365GatewayTools,
    async callTool(toolName: string, input: GatewayJsonObject = {}): Promise<GatewayToolResult> {
      switch (toolName) {
        case 'get_user':
          return jsonResult(
            'Fetched user.',
            await client.get(`/users/${pathSegment(requiredString(input.user, 'user'), 'user')}`, readQuery(input)),
          );

        case 'list_users':
          return jsonResult('Fetched users.', await client.get('/users', readQuery(input)));

        case 'list_subscribed_skus':
          return jsonResult('Fetched subscribed SKUs.', await listSubscribedSkus(client));

        case 'list_groups':
          return jsonResult('Fetched groups.', await client.get('/groups', readQuery(input)));

        case 'get_group':
          return jsonResult(
            'Fetched group.',
            await client.get(`/groups/${pathSegment(requiredString(input.groupId, 'groupId'), 'groupId')}`, readQuery(input)),
          );

        case 'list_group_members':
          return jsonResult(
            'Fetched group members.',
            await client.get(`/groups/${pathSegment(requiredString(input.groupId, 'groupId'), 'groupId')}/members`, readQuery(input)),
          );

        case 'get_user_license_details':
          return jsonResult(
            'Fetched user license details.',
            await client.get(`/users/${pathSegment(requiredString(input.user, 'user'), 'user')}/licenseDetails`),
          );

        case 'create_user':
          return jsonResult(
            'Created user.',
            await createUser(client, {
              displayName: requiredString(input.displayName, 'displayName'),
              userPrincipalName: requiredString(input.userPrincipalName, 'userPrincipalName'),
              mailNickname: stringValue(input.mailNickname),
              password: stringValue(input.password),
              usageLocation: stringValue(input.usageLocation),
              accountEnabled: booleanValue(input.accountEnabled),
              forceChangePasswordNextSignIn: booleanValue(input.forceChangePasswordNextSignIn),
              checkExisting: booleanValue(input.checkExisting),
            }),
          );

        case 'assign_license':
          return jsonResult(
            'Assigned license.',
            await assignLicense(client, {
              userId: requiredString(input.user, 'user'),
              license: requiredString(input.license, 'license'),
              usageLocation: stringValue(input.usageLocation),
              disabledPlans: stringArray(input.disabledPlans),
            }, policy),
          );

        case 'remove_license':
          return jsonResult(
            'Removed license.',
            await removeLicense(client, {
              userId: requiredString(input.user, 'user'),
              license: requiredString(input.license, 'license'),
            }, policy),
          );

        case 'add_group_member':
          return jsonResult(
            'Added group member.',
            await addGroupMember(client, requiredString(input.groupId, 'groupId'), requiredString(input.userId, 'userId'), policy),
          );

        case 'remove_group_member':
          return jsonResult(
            'Removed group member.',
            await removeGroupMember(client, requiredString(input.groupId, 'groupId'), requiredString(input.userId, 'userId'), policy),
          );

        case 'set_usage_location':
          return jsonResult(
            'Set usage location.',
            await setUsageLocation(client, requiredString(input.user, 'user'), requiredString(input.usageLocation, 'usageLocation'), policy),
          );

        case 'update_user':
          return jsonResult(
            'Updated user.',
            await updateUser(client, requiredString(input.user, 'user'), updateUserPatch(input), policy),
          );

        case 'reset_password':
          return jsonResult(
            'Reset password.',
            await resetPassword(client, requiredString(input.user, 'user'), policy),
          );

        case 'set_manager':
          return jsonResult(
            'Set manager.',
            await setManager(client, requiredString(input.user, 'user'), requiredString(input.manager, 'manager'), policy),
          );

        case 'create_temporary_access_pass':
          return jsonResult(
            'Created Temporary Access Pass.',
            await createTemporaryAccessPass(client, requiredString(input.user, 'user'), {
              lifetimeInMinutes: numberValue(input.lifetimeInMinutes),
              isUsableOnce: booleanValue(input.isUsableOnce),
              startDateTime: stringValue(input.startDateTime),
              requireAlphanumeric: booleanValue(input.requireAlphanumeric),
              maxAttempts: numberValue(input.maxAttempts),
            }, policy),
          );

        case 'delete_temporary_access_pass':
          return jsonResult(
            'Deleted Temporary Access Pass.',
            await deleteTemporaryAccessPass(client, requiredString(input.user, 'user'), stringValue(input.methodId), policy),
          );

        case 'activate_pim_role':
          return jsonResult(
            'Activated PIM role.',
            await pimSelfActivateRole(client, {
              roleDefinitionId: requiredString(input.roleDefinitionId, 'roleDefinitionId'),
              duration: stringValue(input.duration),
              justification: stringValue(input.justification),
              directoryScopeId: stringValue(input.directoryScopeId),
            }),
          );

        default:
          if (PERSONAL_TOOL_NAMES.has(toolName)) return callPersonalTool(client, toolName, input);
          return errorResult(`Unsupported Microsoft 365 gateway tool: ${toolName}`);
      }
    },
  };
}

/** A string as given, untrimmed (a body, a file's text), or undefined when it is not a string. */
function rawString(value: GatewayJsonValue | undefined): string | undefined {
  return typeof value === 'string' ? value : undefined;
}

/** Runs one personal tool on the person's own data. A value the tool does not take is a RefusedError. */
async function callPersonalTool(client: GraphClient, toolName: string, input: GatewayJsonObject): Promise<GatewayToolResult> {
  const s = (name: string) => stringValue(input[name]);
  const req = (name: string) => requiredString(input[name], name);
  const n = (name: string) => numberValue(input[name]);
  const b = (name: string) => booleanValue(input[name]);
  const list = (name: string) => stringArray(input[name]);
  const draft = (): personal.DraftInput => ({
    subject: rawString(input.subject),
    body: rawString(input.body),
    bodyType: s('bodyType'),
    to: list('to'),
    cc: list('cc'),
    bcc: list('bcc'),
    importance: s('importance'),
  });
  const event = (): personal.EventInput => ({
    subject: rawString(input.subject),
    start: s('start'),
    end: s('end'),
    timeZone: s('timeZone'),
    allDay: b('allDay'),
    body: rawString(input.body),
    bodyType: s('bodyType'),
    location: rawString(input.location),
    attendees: list('attendees'),
    optionalAttendees: list('optionalAttendees'),
    isOnlineMeeting: b('isOnlineMeeting'),
    showAs: s('showAs'),
  });
  switch (toolName) {
    case 'get_my_profile':
      return jsonResult('Fetched profile.', await personal.getMyProfile(client));
    case 'search_mail':
      return jsonResult(
        'Fetched messages.',
        await personal.searchMail(client, {
          query: s('query'),
          from: s('from'),
          folder: s('folder'),
          unreadOnly: b('unreadOnly'),
          receivedAfter: s('receivedAfter'),
          receivedBefore: s('receivedBefore'),
          top: n('top'),
          skip: n('skip'),
        }),
      );
    case 'get_message':
      return jsonResult('Fetched message.', await personal.getMessage(client, req('messageId'), s('bodyType')));
    case 'get_conversation':
      return jsonResult('Fetched conversation.', await personal.getConversation(client, req('conversationId'), n('top')));
    case 'list_mail_folders':
      return jsonResult('Fetched mail folders.', await personal.listMailFolders(client, s('parentFolderId')));
    case 'list_attachments':
      return jsonResult('Fetched attachments.', await personal.listAttachments(client, req('messageId')));
    case 'get_attachment_text':
      return jsonResult('Fetched attachment.', await personal.getAttachmentText(client, req('messageId'), req('attachmentId')));
    case 'get_automatic_replies':
      return jsonResult('Fetched automatic replies.', await personal.getAutomaticReplies(client));
    case 'create_draft':
      return jsonResult('Created draft.', await personal.createDraft(client, draft()));
    case 'update_draft':
      return jsonResult('Updated draft.', await personal.updateDraft(client, req('messageId'), draft()));
    case 'create_reply_draft':
      return jsonResult('Created reply draft.', await personal.createReplyDraft(client, req('messageId'), rawString(input.comment), b('replyAll')));
    case 'create_forward_draft':
      return jsonResult('Created forward draft.', await personal.createForwardDraft(client, req('messageId'), list('to'), rawString(input.comment)));
    case 'delete_draft':
      return jsonResult('Deleted draft.', await personal.deleteDraft(client, req('messageId')));
    case 'move_message':
      return jsonResult('Moved message.', await personal.moveMessage(client, req('messageId'), req('destination')));
    case 'mark_message':
      return jsonResult(
        'Marked message.',
        await personal.markMessage(client, req('messageId'), { isRead: b('isRead'), flag: s('flag'), categories: list('categories') }),
      );
    case 'send_draft':
      return jsonResult('Sent draft.', await personal.sendDraft(client, req('messageId')));
    case 'send_mail':
      return jsonResult('Sent message.', await personal.sendMail(client, draft()));
    case 'reply':
      return jsonResult('Sent reply.', await personal.reply(client, req('messageId'), rawString(input.comment) ?? '', b('replyAll')));
    case 'forward':
      return jsonResult('Forwarded message.', await personal.forward(client, req('messageId'), list('to') ?? [], rawString(input.comment)));
    case 'set_automatic_replies':
      return jsonResult(
        'Set automatic replies.',
        await personal.setAutomaticReplies(client, {
          status: req('status'),
          internalMessage: rawString(input.internalMessage),
          externalMessage: rawString(input.externalMessage),
          externalAudience: s('externalAudience'),
          start: s('start'),
          end: s('end'),
          timeZone: s('timeZone'),
        }),
      );
    case 'list_calendars':
      return jsonResult('Fetched calendars.', await personal.listCalendars(client));
    case 'list_events':
      return jsonResult(
        'Fetched events.',
        await personal.listEvents(client, { start: req('start'), end: req('end'), calendarId: s('calendarId'), top: n('top'), timeZone: s('timeZone') }),
      );
    case 'get_event':
      return jsonResult('Fetched event.', await personal.getEvent(client, req('eventId'), s('timeZone')));
    case 'get_schedule':
      return jsonResult(
        'Fetched schedule.',
        await personal.getSchedule(client, {
          people: list('people') ?? [],
          start: req('start'),
          end: req('end'),
          timeZone: s('timeZone'),
          intervalMinutes: n('intervalMinutes'),
        }),
      );
    case 'find_meeting_times':
      return jsonResult(
        'Fetched meeting time suggestions.',
        await personal.findMeetingTimes(client, {
          attendees: list('attendees') ?? [],
          durationMinutes: n('durationMinutes') ?? 30,
          start: s('start'),
          end: s('end'),
          timeZone: s('timeZone'),
          maxCandidates: n('maxCandidates'),
        }),
      );
    case 'create_event':
      return jsonResult('Created event.', await personal.createEvent(client, { ...event(), calendarId: s('calendarId') }));
    case 'update_event':
      return jsonResult('Updated event.', await personal.updateEvent(client, req('eventId'), event()));
    case 'respond_to_event':
      return jsonResult(
        'Answered invitation.',
        await personal.respondToEvent(client, req('eventId'), { response: req('response'), comment: rawString(input.comment), sendResponse: b('sendResponse') }),
      );
    case 'cancel_event':
      return jsonResult('Cancelled meeting.', await personal.cancelEvent(client, req('eventId'), rawString(input.comment)));
    case 'delete_event':
      return jsonResult('Deleted appointment.', await personal.deleteEvent(client, req('eventId')));
    case 'search_files':
      return jsonResult('Fetched files.', await personal.searchFiles(client, req('query'), n('top'), n('from')));
    case 'list_folder':
      return jsonResult('Fetched folder.', await personal.listFolder(client, s('driveId'), s('itemId'), n('top')));
    case 'get_file':
      return jsonResult('Fetched file.', await personal.getFile(client, s('driveId'), req('itemId')));
    case 'read_file_text':
      return jsonResult('Fetched file text.', await personal.readFileText(client, s('driveId'), req('itemId')));
    case 'search_sites':
      return jsonResult('Fetched sites.', await personal.searchSites(client, req('query')));
    case 'list_site_drives':
      return jsonResult('Fetched document libraries.', await personal.listSiteDrives(client, req('siteId')));
    case 'upload_file':
      return jsonResult(
        'Uploaded file.',
        await personal.uploadFile(client, {
          name: req('name'),
          driveId: s('driveId'),
          parentId: s('parentId'),
          text: rawString(input.text),
          contentBase64: rawString(input.contentBase64),
          contentType: s('contentType'),
        }),
      );
    case 'upload_file_version':
      return jsonResult(
        'Uploaded new version.',
        await personal.uploadFileVersion(client, {
          driveId: s('driveId'),
          itemId: req('itemId'),
          text: rawString(input.text),
          contentBase64: rawString(input.contentBase64),
          contentType: s('contentType'),
        }),
      );
    case 'create_folder':
      return jsonResult('Created folder.', await personal.createFolder(client, { name: req('name'), driveId: s('driveId'), parentId: s('parentId') }));
    case 'move_file':
      return jsonResult(
        'Moved file.',
        await personal.moveFile(client, { driveId: s('driveId'), itemId: req('itemId'), newName: s('newName'), newParentId: s('newParentId') }),
      );
    default:
      return errorResult(`Unsupported Microsoft 365 gateway tool: ${toolName}`);
  }
}

function readQuery(input: GatewayJsonObject): Record<string, GraphQueryValue> {
  const query: Record<string, GraphQueryValue> = {};
  const top = numberValue(input.top);
  if (top !== undefined) query.$top = top;
  const filter = stringValue(input.filter);
  if (filter) query.$filter = filter;
  const orderby = stringValue(input.orderby);
  if (orderby) query.$orderby = orderby;
  const select = stringArray(input.select);
  if (select && select.length) query.$select = select.join(',');
  return query;
}

function stringValue(value: GatewayJsonValue | undefined): string | undefined {
  return typeof value === 'string' && value.trim().length > 0 ? value.trim() : undefined;
}

function requiredString(value: GatewayJsonValue | undefined, name: string): string {
  const parsed = stringValue(value);
  if (!parsed) {
    throw new Error(`Missing required input: ${name}`);
  }
  return parsed;
}

function numberValue(value: GatewayJsonValue | undefined): number | undefined {
  return typeof value === 'number' ? value : undefined;
}

function booleanValue(value: GatewayJsonValue | undefined): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined;
}

function stringArray(value: GatewayJsonValue | undefined): string[] | undefined {
  if (!Array.isArray(value)) {
    return undefined;
  }
  return value.filter((item): item is string => typeof item === 'string');
}

const UPDATE_USER_STRING_FIELDS = [
  'displayName',
  'givenName',
  'surname',
  'jobTitle',
  'department',
  'companyName',
  'employeeType',
  'mobilePhone',
  'streetAddress',
  'city',
  'postalCode',
  'state',
  'country',
  'officeLocation',
  'usageLocation',
] as const;

function updateUserPatch(input: GatewayJsonObject): UpdateUserInput {
  const patch: UpdateUserInput = {};
  for (const field of UPDATE_USER_STRING_FIELDS) {
    const value = stringValue(input[field]);
    if (value !== undefined) patch[field] = value;
  }
  const accountEnabled = booleanValue(input.accountEnabled);
  if (accountEnabled !== undefined) patch.accountEnabled = accountEnabled;
  const otherMails = stringArray(input.otherMails);
  if (otherMails) patch.otherMails = otherMails;
  const businessPhones = stringArray(input.businessPhones);
  if (businessPhones) patch.businessPhones = businessPhones;
  return patch;
}

function jsonResult(text: string, structuredContent: unknown): GatewayToolResult {
  return {
    content: [{ type: 'text', text }],
    structuredContent: JSON.parse(JSON.stringify(structuredContent ?? null)) as GatewayJsonValue,
  };
}

function errorResult(text: string, failure?: GatewayFailure): GatewayToolResult {
  return {
    isError: true,
    content: [{ type: 'text', text }],
    ...(failure ? { failure } : {}),
  };
}
