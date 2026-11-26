import { describe, expect, it } from 'vitest';
import { createMicrosoft365Gateway, type GatewayJsonObject } from '../src/gateway.js';

/**
 * Every change to an existing user addresses the object the protection guard checked, by the object id the directory
 * gave, also when the tool was given a login name. Graph accepts a login name in some user paths but not in a group's
 * member reference or in a removal by member, where it gives 404; a 404 on removal must never read as "already removed".
 */

const BASE = 'https://graph.example.test/v1.0';
const LOGIN = 'mette@example.com';
const USER_ID = 'aaaaaaaa-1111-4222-8333-444444444444';
const GROUP_ID = 'bbbbbbbb-1111-4222-8333-444444444444';
const MANAGER_ID = 'cccccccc-1111-4222-8333-444444444444';
const PERMISSIVE = { allowPasswordReset: true, allowTap: true, protectedUsers: [], protectedGroups: [] };

interface Options {
  /** The status of the removal (DELETE .../members/{id}/$ref). */
  removeStatus?: number;
  /** The group's direct members, as its member list answers; `'error'` when the list cannot be read. */
  members?: string[] | 'error';
  /** The answer to adding a member. */
  addStatus?: number;
}

function directory(options: Options = {}) {
  const requests: Request[] = [];
  const gateway = createMicrosoft365Gateway({
    accessToken: 'graph-token',
    baseUrl: BASE,
    protection: PERMISSIVE,
    fetchImpl: async (input, init) => {
      const req = new Request(input, init);
      requests.push(req);
      const url = new URL(req.url);
      const path = decodeURIComponent(url.pathname);
      if (req.method === 'GET') {
        if (path.endsWith('/transitiveMemberOf')) return Response.json({ value: [] });
        if (path === '/v1.0/roleManagement/directory/roleAssignments') return Response.json({ value: [] });
        if (path === '/v1.0/subscribedSkus') {
          return Response.json({ value: [{ skuId: 'sku-e3', skuPartNumber: 'SPE_E3', prepaidUnits: { enabled: 5 }, consumedUnits: 1 }] });
        }
        if (path === `/v1.0/groups/${GROUP_ID}/members`) {
          if (options.members === 'error') return Response.json({ error: { message: 'unavailable' } }, { status: 503 });
          return Response.json({ value: (options.members ?? []).map(id => ({ id })) });
        }
        if (path === `/v1.0/groups/${GROUP_ID}` || path === `/v1.0/groups/${GROUP_ID.toUpperCase()}`) {
          return Response.json({ id: GROUP_ID, isAssignableToRole: false });
        }
        if (path.endsWith('/authentication/temporaryAccessPassMethods')) return Response.json({ value: [{ id: 'tap-1' }] });
        // A user looked up by login name or by object id is the same object.
        if (path === `/v1.0/users/${LOGIN}` || path === `/v1.0/users/${USER_ID}`) {
          return Response.json({ id: USER_ID, userPrincipalName: LOGIN, usageLocation: 'DK' });
        }
        return Response.json({ error: { message: 'not found' } }, { status: 404 });
      }
      if (req.method === 'DELETE' && path.endsWith('/$ref') && path.startsWith('/v1.0/groups/')) {
        return new Response(null, { status: options.removeStatus ?? 204 });
      }
      if (req.method === 'POST' && path.endsWith('/members/$ref') && options.addStatus) {
        return Response.json({ error: { message: 'One or more added object references already exist' } }, { status: options.addStatus });
      }
      if (req.method === 'POST' && path.endsWith('/temporaryAccessPassMethods')) {
        return Response.json({ id: 'tap-2', temporaryAccessPass: 'example' }, { status: 201 });
      }
      return new Response(null, { status: 204 });
    },
  });
  const writes = () => requests.filter(r => r.method !== 'GET');
  return { gateway, writes };
}

const BY_LOGIN: ReadonlyArray<[string, GatewayJsonObject]> = [
  ['update_user', { user: LOGIN, jobTitle: 'Sales' }],
  ['set_usage_location', { user: LOGIN, usageLocation: 'DK' }],
  ['set_manager', { user: LOGIN, manager: MANAGER_ID }],
  ['assign_license', { user: LOGIN, license: 'E3', usageLocation: 'DK' }],
  ['assign_license', { user: LOGIN, license: 'E3' }],
  ['remove_license', { user: LOGIN, license: 'E3' }],
  ['add_group_member', { groupId: GROUP_ID, userId: LOGIN }],
  ['remove_group_member', { groupId: GROUP_ID, userId: LOGIN }],
  ['reset_password', { user: LOGIN }],
  ['create_temporary_access_pass', { user: LOGIN }],
  ['delete_temporary_access_pass', { user: LOGIN }],
  ['delete_temporary_access_pass', { user: LOGIN, methodId: 'tap-1' }],
];

describe('a change addresses the object the guard checked', () => {
  it.each(BY_LOGIN)('%s given a login name writes to the object id only', async (tool, input) => {
    const { gateway, writes } = directory();
    const result = await gateway.callTool(tool, input);
    expect(result.isError).toBeFalsy();
    expect(writes().length).toBeGreaterThan(0);
    for (const req of writes()) {
      const path = decodeURIComponent(new URL(req.url).pathname);
      const body = await req.clone().text();
      expect(`${path} ${body}`, `${req.method} ${path}`).not.toContain(LOGIN);
      if (path.startsWith('/v1.0/users/')) expect(path.startsWith(`/v1.0/users/${USER_ID}`), path).toBe(true);
    }
  });

  it('adds a member by a reference to the object id', async () => {
    const { gateway, writes } = directory();
    await gateway.callTool('add_group_member', { groupId: GROUP_ID, userId: LOGIN });
    const [post] = writes();
    expect(new URL(post!.url).pathname).toBe(`/v1.0/groups/${GROUP_ID}/members/$ref`);
    expect(JSON.parse(await post!.text())).toEqual({ '@odata.id': `${BASE}/directoryObjects/${USER_ID}` });
  });

  it('reports an existing membership as already a member, without a second write', async () => {
    const { gateway, writes } = directory({ addStatus: 400 });
    const result = await gateway.callTool('add_group_member', { groupId: GROUP_ID, userId: LOGIN });
    expect(result.structuredContent).toMatchObject({ added: false, alreadyMember: true });
    expect(writes()).toHaveLength(1);
  });

  it('removes a member by object id', async () => {
    const { gateway, writes } = directory({ members: [USER_ID] });
    const result = await gateway.callTool('remove_group_member', { groupId: GROUP_ID, userId: LOGIN });
    expect(result.structuredContent).toMatchObject({ removed: true, alreadyAbsent: false });
    expect(writes().map(r => decodeURIComponent(new URL(r.url).pathname))).toEqual([`/v1.0/groups/${GROUP_ID}/members/${USER_ID}/$ref`]);
  });

  it('reads a 404 on removal as already removed only when the member list does not have the user', async () => {
    const { gateway } = directory({ removeStatus: 404, members: ['dddddddd-1111-4222-8333-444444444444'] });
    const result = await gateway.callTool('remove_group_member', { groupId: GROUP_ID, userId: LOGIN });
    expect(result.structuredContent).toMatchObject({ removed: false, alreadyAbsent: true });
  });

  it('never reports success when the removal gave 404 but the user is still a member', async () => {
    const { gateway } = directory({ removeStatus: 404, members: [USER_ID] });
    await expect(gateway.callTool('remove_group_member', { groupId: GROUP_ID, userId: LOGIN })).rejects.toThrow(/404/);
  });

  it('never reports success when the removal gave 404 and the member list cannot be read', async () => {
    const { gateway } = directory({ removeStatus: 404, members: 'error' });
    await expect(gateway.callTool('remove_group_member', { groupId: GROUP_ID, userId: LOGIN })).rejects.toThrow();
  });

  it('finds a membership on a later page of the member list, and then a 404 on removal is an error', async () => {
    const pages: string[] = [];
    const gateway = createMicrosoft365Gateway({
      accessToken: 'graph-token',
      baseUrl: BASE,
      protection: PERMISSIVE,
      fetchImpl: async (input, init) => {
        const req = new Request(input, init);
        const url = new URL(req.url);
        const path = decodeURIComponent(url.pathname);
        if (req.method === 'DELETE') return Response.json({ error: { message: 'not found' } }, { status: 404 });
        if (path.endsWith('/transitiveMemberOf') || path.endsWith('/roleAssignments')) return Response.json({ value: [] });
        if (path === `/v1.0/groups/${GROUP_ID}/members`) {
          const page = url.searchParams.get('page') ?? '1';
          pages.push(page);
          if (page === '1') {
            return Response.json({
              value: [{ id: 'dddddddd-1111-4222-8333-444444444444' }],
              '@odata.nextLink': `${BASE}/groups/${GROUP_ID}/members?page=2`,
            });
          }
          return Response.json({ value: [{ id: USER_ID }] });
        }
        if (path === `/v1.0/groups/${GROUP_ID}`) return Response.json({ id: GROUP_ID, isAssignableToRole: false });
        return Response.json({ id: USER_ID, userPrincipalName: LOGIN });
      },
    });
    await expect(gateway.callTool('remove_group_member', { groupId: GROUP_ID, userId: LOGIN })).rejects.toThrow(/404/);
    expect(pages).toEqual(['1', '2']);
  });

  it.each(['add_group_member', 'remove_group_member'])(
    '%s writes to the group id the lookup returned, not to the reference it was given',
    async tool => {
      const given = GROUP_ID.toUpperCase();
      const { gateway, writes } = directory({ members: [USER_ID] });
      // The stand-in answers the group under the reference as given, with its own object id.
      const result = await gateway.callTool(tool, { groupId: given, userId: LOGIN });
      expect(result.isError).toBeFalsy();
      const paths = writes().map(r => decodeURIComponent(new URL(r.url).pathname));
      expect(paths.length).toBe(1);
      expect(paths[0]!.startsWith(`/v1.0/groups/${GROUP_ID}/members`), paths[0]).toBe(true);
      expect(paths[0]).not.toContain(given);
    },
  );

  it('in host mode, a removal that gave 404 for a user still in the group is an error, not a success', async () => {
    const gateway = createMicrosoft365Gateway({
      accessToken: 'graph-access-0123456789abcdef',
      policy: { enableWrites: true },
      fetchImpl: async (input, init) => {
        const req = new Request(input, init);
        const path = decodeURIComponent(new URL(req.url).pathname);
        if (req.method === 'DELETE') return Response.json({ error: { message: 'not found' } }, { status: 404 });
        if (path.endsWith('/transitiveMemberOf') || path.endsWith('/roleAssignments')) return Response.json({ value: [] });
        if (path === `/v1.0/groups/${GROUP_ID}/members`) return Response.json({ value: [{ id: USER_ID }] });
        if (path === `/v1.0/groups/${GROUP_ID}`) return Response.json({ id: GROUP_ID, isAssignableToRole: false });
        return Response.json({ id: USER_ID, userPrincipalName: LOGIN });
      },
    });
    const result = await gateway.callTool('remove_group_member', { groupId: GROUP_ID, userId: LOGIN });
    expect(result.isError).toBe(true);
    expect(result.content[0]!.text).toMatch(/HTTP 404/);
  });
});
