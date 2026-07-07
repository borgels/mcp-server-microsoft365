import { describe, expect, it } from 'vitest';
import { createMicrosoft365Gateway, microsoft365GatewayTools } from '../src/gateway.js';

const BASE = 'https://graph.example.test/v1.0';

function makeGateway(handler: (req: Request) => Promise<Response> | Response) {
  const requests: Request[] = [];
  const gateway = createMicrosoft365Gateway({
    accessToken: 'graph-token',
    baseUrl: BASE,
    fetchImpl: async (input, init) => {
      const req = new Request(input, init);
      requests.push(req);
      return handler(req);
    },
  });
  return { gateway, requests };
}

describe('Microsoft 365 gateway export', () => {
  it('exposes read tools enabled by default and write tools disabled by default', () => {
    const reads = microsoft365GatewayTools.filter(tool => tool.riskLevel === 'read');
    const writes = microsoft365GatewayTools.filter(tool => tool.riskLevel === 'write');

    expect(reads.map(tool => tool.name)).toEqual(
      expect.arrayContaining([
        'get_user',
        'list_users',
        'list_subscribed_skus',
        'list_groups',
        'get_group',
        'list_group_members',
        'get_user_license_details',
      ]),
    );
    expect(writes.map(tool => tool.name)).toEqual(
      expect.arrayContaining([
        'create_user',
        'assign_license',
        'remove_license',
        'add_group_member',
        'remove_group_member',
        'set_usage_location',
        'update_user',
        'set_manager',
        'create_temporary_access_pass',
      ]),
    );
    const destructive = microsoft365GatewayTools.filter(tool => tool.riskLevel === 'destructive');
    expect(destructive.map(tool => tool.name)).toEqual(expect.arrayContaining(['delete_temporary_access_pass']));
    expect(reads.every(tool => tool.enabledByDefault)).toBe(true);
    expect(writes.every(tool => !tool.enabledByDefault)).toBe(true);
    expect(destructive.every(tool => !tool.enabledByDefault)).toBe(true);
  });

  it('passes bearer credentials and $-prefixed query params to Graph reads', async () => {
    const { gateway, requests } = makeGateway(() => Response.json({ value: [] }));

    await gateway.callTool('list_users', { top: 5, select: ['id', 'displayName'] });

    expect(requests[0]?.headers.get('Authorization')).toBe('Bearer graph-token');
    const url = new URL(requests[0]!.url);
    expect(url.searchParams.get('$top')).toBe('5');
    expect(url.searchParams.get('$select')).toBe('id,displayName');
  });

  it('creates a user with a generated temporary password after an existence check', async () => {
    const { gateway, requests } = makeGateway(req => {
      const url = new URL(req.url);
      if (req.method === 'GET' && url.pathname.startsWith('/v1.0/users/')) {
        return Response.json({ error: { code: 'Request_ResourceNotFound', message: 'not found' } }, { status: 404 });
      }
      if (req.method === 'POST' && url.pathname === '/v1.0/users') {
        return Response.json(
          { id: 'new-id', userPrincipalName: 'new@contoso.com', displayName: 'New User' },
          { status: 201 },
        );
      }
      throw new Error(`unexpected ${req.method} ${url.pathname}`);
    });

    const result = await gateway.callTool('create_user', {
      displayName: 'New User',
      userPrincipalName: 'new@contoso.com',
    });

    expect(result.structuredContent).toMatchObject({ id: 'new-id', generatedPassword: true });
    expect(typeof (result.structuredContent as { temporaryPassword: string }).temporaryPassword).toBe('string');

    const postBody = JSON.parse((await requests[1]!.text()) || '{}');
    expect(postBody.accountEnabled).toBe(true);
    expect(postBody.passwordProfile.forceChangePasswordNextSignIn).toBe(true);
    expect(postBody.mailNickname).toBe('new');
  });

  it('fails clearly when creating a user whose UPN already exists', async () => {
    const { gateway } = makeGateway(() => Response.json({ id: 'existing' }, { status: 200 }));

    await expect(
      gateway.callTool('create_user', { displayName: 'Dup', userPrincipalName: 'dup@contoso.com' }),
    ).rejects.toThrow(/already exists/);
  });

  it('assigns a license, setting usage location first', async () => {
    const calls: string[] = [];
    const { gateway } = makeGateway(req => {
      const url = new URL(req.url);
      calls.push(`${req.method} ${url.pathname}`);
      if (req.method === 'PATCH' && url.pathname === '/v1.0/users/u1') {
        return new Response(null, { status: 204 });
      }
      if (req.method === 'GET' && url.pathname === '/v1.0/subscribedSkus') {
        return Response.json({
          value: [{ skuId: 'sku-e3', skuPartNumber: 'SPE_E3', prepaidUnits: { enabled: 5 }, consumedUnits: 1 }],
        });
      }
      if (req.method === 'POST' && url.pathname === '/v1.0/users/u1/assignLicense') {
        return Response.json({ id: 'u1' });
      }
      throw new Error(`unexpected ${req.method} ${url.pathname}`);
    });

    const result = await gateway.callTool('assign_license', { user: 'u1', license: 'E3', usageLocation: 'DK' });

    expect(result.structuredContent).toMatchObject({
      userId: 'u1',
      usageLocation: 'DK',
      assigned: { skuId: 'sku-e3', skuPartNumber: 'SPE_E3' },
    });
    // usageLocation PATCH happens before the SKU lookup/assignment.
    expect(calls[0]).toBe('PATCH /v1.0/users/u1');
    expect(calls).toContain('POST /v1.0/users/u1/assignLicense');
  });

  it('refuses license assignment when the user has no usage location', async () => {
    const { gateway } = makeGateway(req => {
      const url = new URL(req.url);
      if (req.method === 'GET' && url.pathname === '/v1.0/users/u2') {
        return Response.json({ usageLocation: null });
      }
      throw new Error(`unexpected ${req.method} ${url.pathname}`);
    });

    await expect(gateway.callTool('assign_license', { user: 'u2', license: 'E3' })).rejects.toThrow(/usageLocation/);
  });

  it('treats an already-existing group membership as idempotent success', async () => {
    const { gateway } = makeGateway(() =>
      Response.json(
        { error: { code: 'Request_BadRequest', message: 'One or more added object references already exist' } },
        { status: 400 },
      ),
    );

    const result = await gateway.callTool('add_group_member', { groupId: 'g1', userId: 'u1' });
    expect(result.structuredContent).toMatchObject({ added: false, alreadyMember: true });
  });

  it('updates only the provided user fields via PATCH, including accountEnabled', async () => {
    const { gateway, requests } = makeGateway(req => {
      const url = new URL(req.url);
      if (req.method === 'PATCH' && url.pathname === '/v1.0/users/u1') {
        return new Response(null, { status: 204 });
      }
      throw new Error(`unexpected ${req.method} ${url.pathname}`);
    });

    const result = await gateway.callTool('update_user', {
      user: 'u1',
      accountEnabled: true,
      jobTitle: 'Sales Executive',
      department: 'Sales',
      otherMails: ['private@example.com'],
    });

    expect(result.structuredContent).toMatchObject({ userId: 'u1' });
    expect((result.structuredContent as { updated: string[] }).updated).toEqual(
      expect.arrayContaining(['accountEnabled', 'jobTitle', 'department', 'otherMails']),
    );
    const body = JSON.parse((await requests[0]!.text()) || '{}');
    expect(body).toEqual({
      accountEnabled: true,
      jobTitle: 'Sales Executive',
      department: 'Sales',
      otherMails: ['private@example.com'],
    });
  });

  it('sets a manager via PUT manager/$ref with a directoryObjects reference', async () => {
    const { gateway, requests } = makeGateway(req => {
      const url = new URL(req.url);
      if (req.method === 'PUT' && url.pathname === '/v1.0/users/u1/manager/$ref') {
        return new Response(null, { status: 204 });
      }
      throw new Error(`unexpected ${req.method} ${url.pathname}`);
    });

    const result = await gateway.callTool('set_manager', { user: 'u1', manager: 'mgr-1' });
    expect(result.structuredContent).toMatchObject({ userId: 'u1', managerId: 'mgr-1' });
    const body = JSON.parse((await requests[0]!.text()) || '{}');
    expect(body['@odata.id']).toContain('/directoryObjects/mgr-1');
  });

  it('regenerates a Temporary Access Pass until the passcode is alphanumeric', async () => {
    let posts = 0;
    const deletes: string[] = [];
    const { gateway } = makeGateway(req => {
      const url = new URL(req.url);
      const path = url.pathname;
      const base = '/v1.0/users/u1/authentication/temporaryAccessPassMethods';
      if (req.method === 'POST' && path === base) {
        posts += 1;
        const pass = posts === 1 ? 'ab^d1234' : 'hn3WugkJ';
        return Response.json(
          { id: `tap-${posts}`, temporaryAccessPass: pass, isUsableOnce: false, lifetimeInMinutes: 60 },
          { status: 201 },
        );
      }
      if (req.method === 'DELETE' && path.startsWith(`${base}/`)) {
        deletes.push(path);
        return new Response(null, { status: 204 });
      }
      throw new Error(`unexpected ${req.method} ${path}`);
    });

    const result = await gateway.callTool('create_temporary_access_pass', { user: 'u1' });
    expect(result.structuredContent).toMatchObject({ temporaryAccessPass: 'hn3WugkJ', id: 'tap-2', attempts: 2 });
    // The rejected (non-alphanumeric) first pass is deleted before retrying.
    expect(deletes).toEqual(['/v1.0/users/u1/authentication/temporaryAccessPassMethods/tap-1']);
  });

  it('deletes all Temporary Access Passes when no methodId is given', async () => {
    const deletes: string[] = [];
    const { gateway } = makeGateway(req => {
      const url = new URL(req.url);
      const path = url.pathname;
      const base = '/v1.0/users/u1/authentication/temporaryAccessPassMethods';
      if (req.method === 'GET' && path === base) {
        return Response.json({ value: [{ id: 't1', temporaryAccessPass: 'x' }, { id: 't2', temporaryAccessPass: 'y' }] });
      }
      if (req.method === 'DELETE' && path.startsWith(`${base}/`)) {
        deletes.push(path);
        return new Response(null, { status: 204 });
      }
      throw new Error(`unexpected ${req.method} ${path}`);
    });

    const result = await gateway.callTool('delete_temporary_access_pass', { user: 'u1' });
    expect(result.structuredContent).toMatchObject({ deleted: ['t1', 't2'] });
    expect(deletes.length).toBe(2);
  });

  it('returns an error result for unsupported tools', async () => {
    const { gateway } = makeGateway(() => Response.json({}));
    const result = await gateway.callTool('does_not_exist');
    expect(result.isError).toBe(true);
  });
});
