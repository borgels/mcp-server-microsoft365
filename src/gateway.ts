import { GraphClient, type GraphClientOptions, type GraphQueryValue } from './graph/client.js';
import {
  addGroupMember,
  assignLicense,
  createUser,
  listSubscribedSkus,
  removeGroupMember,
  removeLicense,
  setUsageLocation,
} from './graph/operations.js';

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

export interface GatewayToolResult {
  content: Array<{ type: 'text'; text: string }>;
  structuredContent?: GatewayJsonValue;
  isError?: boolean;
}

export interface Microsoft365GatewayOptions extends GraphClientOptions {}

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

export const microsoft365GatewayTools: GatewayToolDefinition[] = [
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
];

export function createMicrosoft365Gateway(options: Microsoft365GatewayOptions = {}) {
  const client = new GraphClient(options);

  return {
    tools: microsoft365GatewayTools,
    async callTool(toolName: string, input: GatewayJsonObject = {}): Promise<GatewayToolResult> {
      switch (toolName) {
        case 'get_user':
          return jsonResult(
            'Fetched user.',
            await client.get(`/users/${encodeURIComponent(requiredString(input.user, 'user'))}`, readQuery(input)),
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
            await client.get(`/groups/${encodeURIComponent(requiredString(input.groupId, 'groupId'))}`, readQuery(input)),
          );

        case 'list_group_members':
          return jsonResult(
            'Fetched group members.',
            await client.get(`/groups/${encodeURIComponent(requiredString(input.groupId, 'groupId'))}/members`, readQuery(input)),
          );

        case 'get_user_license_details':
          return jsonResult(
            'Fetched user license details.',
            await client.get(`/users/${encodeURIComponent(requiredString(input.user, 'user'))}/licenseDetails`),
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
            }),
          );

        case 'remove_license':
          return jsonResult(
            'Removed license.',
            await removeLicense(client, {
              userId: requiredString(input.user, 'user'),
              license: requiredString(input.license, 'license'),
            }),
          );

        case 'add_group_member':
          return jsonResult(
            'Added group member.',
            await addGroupMember(client, requiredString(input.groupId, 'groupId'), requiredString(input.userId, 'userId')),
          );

        case 'remove_group_member':
          return jsonResult(
            'Removed group member.',
            await removeGroupMember(client, requiredString(input.groupId, 'groupId'), requiredString(input.userId, 'userId')),
          );

        case 'set_usage_location':
          return jsonResult(
            'Set usage location.',
            await setUsageLocation(client, requiredString(input.user, 'user'), requiredString(input.usageLocation, 'usageLocation')),
          );

        default:
          return errorResult(`Unsupported Microsoft 365 gateway tool: ${toolName}`);
      }
    },
  };
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

function jsonResult(text: string, structuredContent: unknown): GatewayToolResult {
  return {
    content: [{ type: 'text', text }],
    structuredContent: JSON.parse(JSON.stringify(structuredContent ?? null)) as GatewayJsonValue,
  };
}

function errorResult(text: string): GatewayToolResult {
  return {
    isError: true,
    content: [{ type: 'text', text }],
  };
}
