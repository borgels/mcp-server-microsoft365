import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod/v4';
import { formatUnknownError } from '../errors.js';
import { type GraphClient, type GraphQueryValue } from '../graph/client.js';
import {
  addGroupMember,
  assignLicense,
  createTemporaryAccessPass,
  createUser,
  deleteTemporaryAccessPass,
  listSubscribedSkus,
  removeGroupMember,
  removeLicense,
  setManager,
  setUsageLocation,
  updateUser,
} from '../graph/operations.js';

export const READ_TOOL_ANNOTATIONS = {
  readOnlyHint: true,
  destructiveHint: false,
  idempotentHint: true,
  openWorldHint: true,
} as const;

export const WRITE_TOOL_ANNOTATIONS = {
  readOnlyHint: false,
  destructiveHint: false,
  idempotentHint: false,
  openWorldHint: true,
} as const;

export const DESTRUCTIVE_TOOL_ANNOTATIONS = {
  readOnlyHint: false,
  destructiveHint: true,
  idempotentHint: true,
  openWorldHint: true,
} as const;

const userRef = z.string().trim().min(1);
const groupId = z.string().trim().min(1);
const license = z.string().trim().min(1);
const usageLocation = z
  .string()
  .trim()
  .regex(/^[A-Za-z]{2}$/, 'Use an ISO 3166-1 alpha-2 country code, e.g. "DK".');
const selectShape = z.array(z.string().trim().min(1)).max(50).optional();
const listShape = {
  top: z.number().int().min(1).max(999).optional(),
  filter: z.string().trim().min(1).optional(),
  select: selectShape,
  orderby: z.string().trim().min(1).optional(),
};

export function registerMicrosoft365Tools(server: McpServer, client: GraphClient): void {
  server.registerTool(
    'get_user',
    {
      title: 'Get Microsoft 365 User',
      description: 'Fetch one user by object id or userPrincipalName.',
      inputSchema: { user: userRef, select: selectShape },
      annotations: READ_TOOL_ANNOTATIONS,
    },
    async input =>
      runTool(async () => jsonToolResult(await client.get(`/users/${encodeURIComponent(input.user)}`, readQuery(input)))),
  );

  server.registerTool(
    'list_users',
    {
      title: 'List Microsoft 365 Users',
      description: 'List users with optional $top, $filter, $select, and $orderby.',
      inputSchema: { ...listShape },
      annotations: READ_TOOL_ANNOTATIONS,
    },
    async input => runTool(async () => jsonToolResult(await client.get('/users', readQuery(input)))),
  );

  server.registerTool(
    'list_subscribed_skus',
    {
      title: 'List Subscribed SKUs',
      description: 'List tenant subscribed SKUs for license lookup and seat availability.',
      inputSchema: {},
      annotations: READ_TOOL_ANNOTATIONS,
    },
    async () => runTool(async () => jsonToolResult(await listSubscribedSkus(client))),
  );

  server.registerTool(
    'list_groups',
    {
      title: 'List Microsoft 365 Groups',
      description: 'List groups with optional $top, $filter, $select, and $orderby.',
      inputSchema: { ...listShape },
      annotations: READ_TOOL_ANNOTATIONS,
    },
    async input => runTool(async () => jsonToolResult(await client.get('/groups', readQuery(input)))),
  );

  server.registerTool(
    'get_group',
    {
      title: 'Get Microsoft 365 Group',
      description: 'Fetch one group by object id.',
      inputSchema: { groupId, select: selectShape },
      annotations: READ_TOOL_ANNOTATIONS,
    },
    async input =>
      runTool(async () => jsonToolResult(await client.get(`/groups/${encodeURIComponent(input.groupId)}`, readQuery(input)))),
  );

  server.registerTool(
    'list_group_members',
    {
      title: 'List Group Members',
      description: 'List the members of a group.',
      inputSchema: { groupId, top: z.number().int().min(1).max(999).optional() },
      annotations: READ_TOOL_ANNOTATIONS,
    },
    async input =>
      runTool(async () =>
        jsonToolResult(await client.get(`/groups/${encodeURIComponent(input.groupId)}/members`, readQuery(input))),
      ),
  );

  server.registerTool(
    'get_user_license_details',
    {
      title: 'Get User License Details',
      description: 'List the licenses currently assigned to a user.',
      inputSchema: { user: userRef },
      annotations: READ_TOOL_ANNOTATIONS,
    },
    async input =>
      runTool(async () => jsonToolResult(await client.get(`/users/${encodeURIComponent(input.user)}/licenseDetails`))),
  );

  server.registerTool(
    'create_user',
    {
      title: 'Create Microsoft 365 User',
      description:
        'Create a user with a generated temporary password. Returns the temporary password; store it securely and share out-of-band.',
      inputSchema: {
        displayName: z.string().trim().min(1),
        userPrincipalName: z.string().trim().min(3),
        mailNickname: z.string().trim().min(1).optional(),
        password: z.string().min(8).optional(),
        usageLocation: usageLocation.optional(),
        accountEnabled: z.boolean().optional(),
        forceChangePasswordNextSignIn: z.boolean().optional(),
        checkExisting: z.boolean().optional(),
      },
      annotations: WRITE_TOOL_ANNOTATIONS,
    },
    async input => runTool(async () => jsonToolResult(await createUser(client, input))),
  );

  server.registerTool(
    'assign_license',
    {
      title: 'Assign License To User',
      description:
        'Assign a license to a user. Sets usageLocation first when provided; fails clearly if the user has no usageLocation.',
      inputSchema: {
        user: userRef,
        license,
        usageLocation: usageLocation.optional(),
        disabledPlans: z.array(z.string().trim().min(1)).max(200).optional(),
      },
      annotations: WRITE_TOOL_ANNOTATIONS,
    },
    async input =>
      runTool(async () =>
        jsonToolResult(
          await assignLicense(client, {
            userId: input.user,
            license: input.license,
            usageLocation: input.usageLocation,
            disabledPlans: input.disabledPlans,
          }),
        ),
      ),
  );

  server.registerTool(
    'remove_license',
    {
      title: 'Remove License From User',
      description: 'Remove a license from a user.',
      inputSchema: { user: userRef, license },
      annotations: DESTRUCTIVE_TOOL_ANNOTATIONS,
    },
    async input =>
      runTool(async () => jsonToolResult(await removeLicense(client, { userId: input.user, license: input.license }))),
  );

  server.registerTool(
    'add_group_member',
    {
      title: 'Add Group Member',
      description: 'Add a user to a group. Idempotent when the user is already a member.',
      inputSchema: { groupId, userId: userRef },
      annotations: WRITE_TOOL_ANNOTATIONS,
    },
    async input => runTool(async () => jsonToolResult(await addGroupMember(client, input.groupId, input.userId))),
  );

  server.registerTool(
    'remove_group_member',
    {
      title: 'Remove Group Member',
      description: 'Remove a user from a group. Idempotent when the user is not a member.',
      inputSchema: { groupId, userId: userRef },
      annotations: DESTRUCTIVE_TOOL_ANNOTATIONS,
    },
    async input => runTool(async () => jsonToolResult(await removeGroupMember(client, input.groupId, input.userId))),
  );

  server.registerTool(
    'set_usage_location',
    {
      title: 'Set User Usage Location',
      description: 'Set a user usageLocation (required before license assignment).',
      inputSchema: { user: userRef, usageLocation },
      annotations: WRITE_TOOL_ANNOTATIONS,
    },
    async input => runTool(async () => jsonToolResult(await setUsageLocation(client, input.user, input.usageLocation))),
  );

  server.registerTool(
    'update_user',
    {
      title: 'Update Microsoft 365 User',
      description:
        'Update attributes on an existing user (PATCH). Use accountEnabled to enable/disable an existing account (e.g. activate a pre-created hire or disable a leaver). Only provided fields are changed.',
      inputSchema: {
        user: userRef,
        accountEnabled: z.boolean().optional(),
        displayName: z.string().trim().min(1).optional(),
        givenName: z.string().trim().min(1).optional(),
        surname: z.string().trim().min(1).optional(),
        jobTitle: z.string().trim().min(1).optional(),
        department: z.string().trim().min(1).optional(),
        companyName: z.string().trim().min(1).optional(),
        employeeType: z.string().trim().min(1).optional(),
        mobilePhone: z.string().trim().min(1).optional(),
        streetAddress: z.string().trim().min(1).optional(),
        city: z.string().trim().min(1).optional(),
        postalCode: z.string().trim().min(1).optional(),
        state: z.string().trim().min(1).optional(),
        country: z.string().trim().min(1).optional(),
        officeLocation: z.string().trim().min(1).optional(),
        usageLocation: usageLocation.optional(),
        otherMails: z.array(z.string().trim().min(1)).max(10).optional(),
        businessPhones: z.array(z.string().trim().min(1)).max(10).optional(),
      },
      annotations: WRITE_TOOL_ANNOTATIONS,
    },
    async input => {
      const { user, ...patch } = input;
      return runTool(async () => jsonToolResult(await updateUser(client, user, patch)));
    },
  );

  server.registerTool(
    'set_manager',
    {
      title: 'Set User Manager',
      description: "Set a user's manager. Provide the manager's object id or userPrincipalName.",
      inputSchema: { user: userRef, manager: userRef },
      annotations: WRITE_TOOL_ANNOTATIONS,
    },
    async input => runTool(async () => jsonToolResult(await setManager(client, input.user, input.manager))),
  );

  server.registerTool(
    'create_temporary_access_pass',
    {
      title: 'Create Temporary Access Pass',
      description:
        'Create a Temporary Access Pass (TAP) for passwordless first sign-in / MFA setup. Multi-use by default; regenerated until the passcode is alphanumeric so it is easy to relay. lifetimeInMinutes is bounded by the tenant TAP policy. Returns the passcode; deliver it out-of-band.',
      inputSchema: {
        user: userRef,
        lifetimeInMinutes: z.number().int().min(10).max(43200).optional(),
        isUsableOnce: z.boolean().optional(),
        startDateTime: z.string().trim().min(1).optional(),
        requireAlphanumeric: z.boolean().optional(),
        maxAttempts: z.number().int().min(1).max(20).optional(),
      },
      annotations: WRITE_TOOL_ANNOTATIONS,
    },
    async input => {
      const { user, ...tap } = input;
      return runTool(async () => jsonToolResult(await createTemporaryAccessPass(client, user, tap)));
    },
  );

  server.registerTool(
    'delete_temporary_access_pass',
    {
      title: 'Delete Temporary Access Pass',
      description: "Delete a user's Temporary Access Pass. With methodId deletes that pass; otherwise deletes every TAP on the user.",
      inputSchema: { user: userRef, methodId: z.string().trim().min(1).optional() },
      annotations: DESTRUCTIVE_TOOL_ANNOTATIONS,
    },
    async input => runTool(async () => jsonToolResult(await deleteTemporaryAccessPass(client, input.user, input.methodId))),
  );
}

async function runTool(call: () => Promise<ReturnType<typeof jsonToolResult>>) {
  try {
    return await call();
  } catch (error) {
    return {
      isError: true,
      content: [{ type: 'text' as const, text: formatUnknownError(error) }],
    };
  }
}

function readQuery(input: {
  top?: number;
  filter?: string;
  select?: string[];
  orderby?: string;
}): Record<string, GraphQueryValue> {
  const query: Record<string, GraphQueryValue> = {};
  if (input.top !== undefined) query.$top = input.top;
  if (input.filter) query.$filter = input.filter;
  if (input.orderby) query.$orderby = input.orderby;
  if (input.select && input.select.length) query.$select = input.select.join(',');
  return query;
}

function jsonToolResult(data: unknown) {
  return {
    content: [
      {
        type: 'text' as const,
        text: JSON.stringify(data, null, 2),
      },
    ],
  };
}
