import type { IncomingMessage } from 'node:http';
import { describe, expect, it } from 'vitest';
import { createMicrosoft365Gateway } from '../src/gateway.js';
import { GraphClient } from '../src/graph/client.js';
import { resolveSkuId } from '../src/graph/licenses.js';
import { assertSafePathParam } from '../src/graph/path.js';
import { protectionPolicyFromEnv, type ProtectionPolicy } from '../src/graph/protection.js';
import { assertAuthorized, constantTimeEquals } from '../src/transports/http-helpers.js';

const BASE = 'https://graph.example.test/v1.0';
const ENABLED: ProtectionPolicy = { allowPasswordReset: true, allowTap: true, protectedUsers: [], protectedGroups: [] };

interface Directory {
  memberOf?: unknown[] | ((page: number) => { value: unknown[]; next?: boolean });
  memberOfStatus?: number;
  roleAssignments?: unknown[];
  roleAssignmentsStatus?: number;
  group?: Record<string, unknown>;
  user?: Record<string, unknown>;
}

function setup(directory: Directory, policy: Partial<ProtectionPolicy> = ENABLED) {
  const requests: Request[] = [];
  const gateway = createMicrosoft365Gateway({
    accessToken: 'graph-token',
    baseUrl: BASE,
    protection: policy,
    fetchImpl: async (input, init) => {
      const req = new Request(input, init);
      requests.push(req);
      const url = new URL(req.url);
      const path = url.pathname;
      if (req.method === 'GET' && path.endsWith('/transitiveMemberOf')) {
        if (directory.memberOfStatus) return Response.json({ error: { message: 'Forbidden' } }, { status: directory.memberOfStatus });
        const m = directory.memberOf ?? [];
        if (typeof m === 'function') {
          const page = Number(url.searchParams.get('page') ?? '1');
          const { value, next } = m(page);
          return Response.json({
            value,
            ...(next ? { '@odata.nextLink': `${BASE}/users/u-1/transitiveMemberOf?page=${page + 1}` } : {}),
          });
        }
        return Response.json({ value: m });
      }
      if (req.method === 'GET' && path === '/v1.0/roleManagement/directory/roleAssignments') {
        if (directory.roleAssignmentsStatus) return Response.json({ error: { message: 'Forbidden' } }, { status: directory.roleAssignmentsStatus });
        return Response.json({ value: directory.roleAssignments ?? [] });
      }
      if (req.method === 'GET' && path.startsWith('/v1.0/groups/')) {
        return Response.json(directory.group ?? { id: 'g-1', displayName: 'Sales', isAssignableToRole: false });
      }
      if (req.method === 'GET' && path.startsWith('/v1.0/users/')) {
        return Response.json(directory.user ?? { id: 'u-1', userPrincipalName: 'someone@example.com', mail: 'someone@example.com' });
      }
      return new Response(null, { status: 204 });
    },
  });
  const writes = () => requests.filter(r => r.method !== 'GET');
  return { gateway, requests, writes };
}

describe('privileged target protection', () => {
  it('refuses password reset and TAP unless their own flags are enabled', async () => {
    const { gateway, requests } = setup({}, { allowPasswordReset: false, allowTap: false });
    await expect(gateway.callTool('reset_password', { user: 'u-1' })).rejects.toThrow(/MS_ENABLE_PASSWORD_RESET/);
    await expect(gateway.callTool('create_temporary_access_pass', { user: 'u-1' })).rejects.toThrow(/MS_ENABLE_TAP/);
    await expect(gateway.callTool('delete_temporary_access_pass', { user: 'u-1' })).rejects.toThrow(/MS_ENABLE_TAP/);
    expect(requests).toHaveLength(0);
  });

  it('defaults both flags to off from the environment', () => {
    expect(protectionPolicyFromEnv({})).toEqual({ allowPasswordReset: false, allowTap: false, protectedUsers: [], protectedGroups: [] });
    expect(
      protectionPolicyFromEnv({
        MS_ENABLE_PASSWORD_RESET: 'true',
        MS_ENABLE_TAP: '1',
        MS_PROTECTED_USERS: ' Admin@Example.com , break-glass@example.com,',
        MS_PROTECTED_GROUPS: 'GROUP-ID',
      }),
    ).toEqual({
      allowPasswordReset: true,
      allowTap: false,
      protectedUsers: ['admin@example.com', 'break-glass@example.com'],
      protectedGroups: ['group-id'],
    });
  });

  it('allows an ordinary user', async () => {
    const { gateway, writes } = setup({});
    await gateway.callTool('reset_password', { user: 'u-1' });
    expect(writes().map(r => r.method)).toEqual(['PATCH']);
  });

  it.each(['reset_password', 'create_temporary_access_pass'])('%s refuses a user holding a directory role', async tool => {
    const { gateway, writes } = setup({
      memberOf: [{ '@odata.type': '#microsoft.graph.directoryRole', id: 'r-1', displayName: 'Global Administrator' }],
    });
    await expect(gateway.callTool(tool, { user: 'admin@example.com' })).rejects.toThrow(/directory role/);
    expect(writes()).toHaveLength(0);
  });

  it('refuses a member of a role-assignable group', async () => {
    const { gateway, writes } = setup({
      memberOf: [{ '@odata.type': '#microsoft.graph.group', id: 'g-9', displayName: 'Tier0', isAssignableToRole: true }],
    });
    await expect(gateway.callTool('reset_password', { user: 'u-1' })).rejects.toThrow(/role-assignable group/);
    expect(writes()).toHaveLength(0);
  });

  it('finds a directory role on a later page of memberships', async () => {
    const { gateway, writes } = setup({
      memberOf: page =>
        page === 1
          ? { value: [{ '@odata.type': '#microsoft.graph.group', id: 'g-1' }], next: true }
          : { value: [{ '@odata.type': '#microsoft.graph.directoryRole', id: 'r-1' }] },
    });
    await expect(gateway.callTool('reset_password', { user: 'u-1' })).rejects.toThrow(/directory role/);
    expect(writes()).toHaveLength(0);
  });

  it('refuses a user with a scoped role assignment', async () => {
    const { gateway, writes } = setup({ roleAssignments: [{ id: 'ra-1', roleDefinitionId: 'x', directoryScopeId: '/administrativeUnits/a' }] });
    await expect(gateway.callTool('create_temporary_access_pass', { user: 'u-1' })).rejects.toThrow(/role assignment/);
    expect(writes()).toHaveLength(0);
  });

  it('refuses users listed in MS_PROTECTED_USERS (matched on resolved UPN)', async () => {
    const { gateway, writes } = setup({}, { ...ENABLED, protectedUsers: ['someone@example.com'] });
    await expect(gateway.callTool('reset_password', { user: 'u-1' })).rejects.toThrow(/MS_PROTECTED_USERS/);
    expect(writes()).toHaveLength(0);
  });

  it.each([
    ['memberships', { memberOfStatus: 403 }],
    ['role assignments', { roleAssignmentsStatus: 403 }],
  ])('fails closed when %s cannot be read', async (_label, directory) => {
    const { gateway, writes } = setup(directory as Directory);
    await expect(gateway.callTool('reset_password', { user: 'u-1' })).rejects.toThrow(/Failing closed/);
    expect(writes()).toHaveLength(0);
  });

  it.each(['add_group_member', 'remove_group_member'])('%s refuses a role-assignable group', async tool => {
    const { gateway, writes } = setup({ group: { id: 'g-1', displayName: 'Tier0', isAssignableToRole: true } });
    await expect(gateway.callTool(tool, { groupId: 'g-1', userId: 'u-1' })).rejects.toThrow(/role-assignable/);
    expect(writes()).toHaveLength(0);
  });

  it('refuses groups listed in MS_PROTECTED_GROUPS', async () => {
    const { gateway, writes } = setup({}, { ...ENABLED, protectedGroups: ['sales'] });
    await expect(gateway.callTool('add_group_member', { groupId: 'g-1', userId: 'u-1' })).rejects.toThrow(/MS_PROTECTED_GROUPS/);
    expect(writes()).toHaveLength(0);
  });

  it('refuses adding a privileged user to an ordinary group', async () => {
    const { gateway, writes } = setup({ memberOf: [{ '@odata.type': '#microsoft.graph.directoryRole', id: 'r-1' }] });
    await expect(gateway.callTool('add_group_member', { groupId: 'g-1', userId: 'u-1' })).rejects.toThrow(/directory role/);
    expect(writes()).toHaveLength(0);
  });

  it('adds an ordinary user to an ordinary group', async () => {
    const { gateway, writes } = setup({});
    await gateway.callTool('add_group_member', { groupId: 'g-1', userId: 'u-1' });
    expect(writes().map(r => new URL(r.url).pathname)).toEqual(['/v1.0/groups/g-1/members/$ref']);
  });
});

describe('path parameter guard', () => {
  it.each(['', ' ', '.', '..', 'a/b', 'a\\b', 'a?b', 'a#b', '%2e%2e', '..%2F', '%5c', 'a\nb'])('rejects %j', value => {
    expect(() => assertSafePathParam(value, 'user')).toThrow(/Invalid user/);
  });

  it.each(['user@example.com', '05e9a617-0261-4cee-bb44-138d3ef5d965', "o'brien@example.com", 'a.b'])('accepts %j', value => {
    expect(assertSafePathParam(value)).toBe(value);
  });

  it('stops a dot-segment user reference before any request is sent', async () => {
    const { gateway, requests } = setup({});
    await expect(gateway.callTool('get_user', { user: '..' })).rejects.toThrow(/Invalid user/);
    await expect(gateway.callTool('get_group', { groupId: '../users' })).rejects.toThrow(/Invalid groupId/);
    await expect(gateway.callTool('remove_license', { user: '%2e%2e', license: 'E3' })).rejects.toThrow(/Invalid/);
    expect(requests).toHaveLength(0);
  });
});

describe('Graph client hardening', () => {
  it('does not follow a nextLink to another origin with the bearer token', async () => {
    const seen: string[] = [];
    const client = new GraphClient({
      accessToken: 't',
      baseUrl: BASE,
      fetchImpl: async input => {
        seen.push(String(input));
        return Response.json({ value: [] });
      },
    });
    await expect(client.get('https://evil.example.test/v1.0/users')).rejects.toThrow(/different origin/);
    expect(seen).toHaveLength(0);
  });

  it('trims trailing slashes from the base URL in linear time', () => {
    const start = Date.now();
    const client = new GraphClient({ accessToken: 't', baseUrl: `https://graph.example.test/v1.0${'/'.repeat(100_000)}` });
    expect(client.baseUrl).toBe('https://graph.example.test/v1.0');
    new GraphClient({ accessToken: 't', baseUrl: `https://graph.example.test/${'/'.repeat(100_000)}x` });
    expect(Date.now() - start).toBeLessThan(1_000);
  });

  it('normalises license names in linear time', () => {
    const start = Date.now();
    expect(resolveSkuId([], `${'_'.repeat(100_000)}!`)).toBeUndefined();
    expect(Date.now() - start).toBeLessThan(1_000);
  });
});

describe('HTTP bearer token check', () => {
  const req = (authorization?: string) => ({ headers: { authorization } }) as unknown as IncomingMessage;
  const config = { host: '127.0.0.1', port: 0, maxBodyBytes: 1, httpToken: 'secret-token' };

  it('accepts the exact token and rejects anything else', () => {
    expect(() => assertAuthorized(req('Bearer secret-token'), config)).not.toThrow();
    for (const bad of [undefined, '', 'Bearer secret-toke', 'Bearer secret-token2', 'bearer secret-token', 'secret-token']) {
      expect(() => assertAuthorized(req(bad), config)).toThrow(/Unauthorized/);
    }
  });

  it('compares in constant time regardless of length', () => {
    expect(constantTimeEquals('a', 'a')).toBe(true);
    expect(constantTimeEquals('a', 'a'.repeat(1000))).toBe(false);
  });
});
