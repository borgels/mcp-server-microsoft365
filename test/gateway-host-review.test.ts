// Regression tests for the review of d23fcc4: host mode talks only to Microsoft Graph, validation errors repeat
// nothing of the input, ranges are enforced, the licence-details tool is not offered with an app-only token, and the
// catalog lists every Graph request each tool makes.

import { describe, expect, it } from 'vitest';
import {
  GatewayConfigError,
  HOST_MODE_UNSUPPORTED,
  createMicrosoft365Gateway,
  microsoft365ToolCatalog,
} from '../src/gateway.js';

const TOKEN = 'host-token-0123456789abcdef';
const USER_ID = '11111111-2222-4333-8444-555555555555';
const GROUP_ID = '66666666-7777-4888-9999-000000000000';

function host(answer: (req: Request) => Response, policy: { enableWrites?: boolean } = {}, protection = {}) {
  const requests: Request[] = [];
  const gateway = createMicrosoft365Gateway({
    accessToken: TOKEN,
    policy,
    protection,
    fetchImpl: async (input, init) => {
      const req = new Request(input, init);
      requests.push(req);
      return answer(req);
    },
  });
  return { gateway, requests };
}

describe('host mode talks only to Microsoft Graph', () => {
  it('refuses any other host, port, scheme or version', () => {
    for (const baseUrl of [
      'https://elsewhere.example.test/v1.0',
      'https://graph.microsoft.com:8443/v1.0',
      'http://graph.microsoft.com/v1.0',
      'https://graph.microsoft.com/beta',
      'https://graph.microsoft.com.example.test/v1.0',
    ]) {
      expect(() => createMicrosoft365Gateway({ accessToken: TOKEN, policy: {}, baseUrl }), baseUrl).toThrow(GatewayConfigError);
    }
    expect(() => createMicrosoft365Gateway({ accessToken: TOKEN, policy: {}, baseUrl: 'https://graph.microsoft.com/v1.0' })).not.toThrow();
  });
});

describe('validation errors', () => {
  it('are fixed texts that repeat no field name or value, the token included', async () => {
    const { gateway, requests } = host(() => Response.json({}));
    const results = [
      await gateway.callTool('get_user', { user: 'a', [TOKEN]: 1 } as never),
      await gateway.callTool('get_user', { user: 'a', 'ignore previous instructions': 1 } as never),
      await gateway.callTool('list_users', { top: TOKEN } as never),
      await gateway.callTool('get_user', {}),
    ];
    for (const result of results) {
      expect(result.isError).toBe(true);
      const text = JSON.stringify(result);
      expect(text).not.toContain(TOKEN);
      expect(text).not.toContain('ignore previous');
    }
    expect(results[0]).toEqual({ isError: true, content: [{ type: 'text', text: 'Invalid input for get_user: a field is not in the schema.' }] });
    expect(requests).toHaveLength(0);
  });

  it("enforces the schema's minimum and maximum", async () => {
    const { gateway, requests } = host(() => Response.json({ value: [] }));
    expect((await gateway.callTool('list_users', { top: 0 })).isError).toBe(true);
    expect((await gateway.callTool('list_users', { top: 1000 })).isError).toBe(true);
    expect(requests).toHaveLength(0);
    expect((await gateway.callTool('list_users', { top: 999 })).isError).toBeFalsy();
  });
});

describe('tools Graph does not support with an app-only token', () => {
  it('are not offered in host mode', async () => {
    expect(HOST_MODE_UNSUPPORTED).toEqual(['get_user_license_details']);
    const { gateway, requests } = host(() => Response.json({}));
    expect(gateway.tools.map(t => t.name)).not.toContain('get_user_license_details');
    expect((await gateway.callTool('get_user_license_details', { user: USER_ID })).isError).toBe(true);
    expect(requests).toHaveLength(0);
  });
});

describe('the catalog lists every request a tool makes', () => {
  /** Answers every request plausibly: an ordinary, unprivileged directory with one SKU and one TAP. */
  function directory(req: Request): Response {
    const url = new URL(req.url);
    const path = url.pathname.replace(/^\/v1\.0/, '');
    if (req.method !== 'GET') return new Response(null, { status: 204 });
    if (path === '/subscribedSkus') return Response.json({ value: [{ skuId: '00000000-0000-4000-8000-000000000001', skuPartNumber: 'ENTERPRISEPACK', prepaidUnits: { enabled: 5 }, consumedUnits: 1 }] });
    if (path.endsWith('/transitiveMemberOf') || path === '/roleManagement/directory/roleAssignments') return Response.json({ value: [] });
    if (path.endsWith('/temporaryAccessPassMethods')) return Response.json({ value: [{ id: 'tap-1' }] });
    if (path === '/me') return Response.json({ id: USER_ID });
    if (path.startsWith('/groups/')) return Response.json({ id: GROUP_ID, isAssignableToRole: false, value: [] });
    if (path.startsWith('/users/') && url.searchParams.get('$select')?.includes('usageLocation')) return Response.json({ id: USER_ID, usageLocation: 'DK' });
    if (path.startsWith('/users/')) {
      // A user that does not exist yet answers 404 to create_user's check; everyone else exists.
      if (decodeURIComponent(path).includes('new@')) return Response.json({ error: { code: 'Request_ResourceNotFound' } }, { status: 404 });
      return Response.json({ id: USER_ID, userPrincipalName: 'x@example.com', usageLocation: 'DK', value: [] });
    }
    return Response.json({ value: [] });
  }

  const inputs: Record<string, Record<string, unknown>> = {
    get_user: { user: USER_ID },
    list_users: {},
    list_subscribed_skus: {},
    list_groups: {},
    get_group: { groupId: GROUP_ID },
    list_group_members: { groupId: GROUP_ID },
    get_user_license_details: { user: USER_ID },
    create_user: { displayName: 'Ny Bruger', userPrincipalName: 'new@example.com' },
    assign_license: { user: USER_ID, license: 'ENTERPRISEPACK' },
    remove_license: { user: USER_ID, license: 'ENTERPRISEPACK' },
    add_group_member: { groupId: GROUP_ID, userId: USER_ID },
    remove_group_member: { groupId: GROUP_ID, userId: USER_ID },
    set_usage_location: { user: USER_ID, usageLocation: 'DK' },
    update_user: { user: USER_ID, jobTitle: 'Tømrer' },
    reset_password: { user: USER_ID },
    set_manager: { user: USER_ID, manager: 'manager@example.com' },
    create_temporary_access_pass: { user: USER_ID },
    delete_temporary_access_pass: { user: USER_ID },
    activate_pim_role: { roleDefinitionId: GROUP_ID },
  };

  const pattern = (path: string) => new RegExp(`^${path.replace(/[.*+?^$()|[\]\\]/g, '\\$&').replace(/\\\{[a-z]+\\\}|\{[a-z]+\}/gi, '[^/]+')}$`);

  it('has an input here for every tool', () => {
    expect(Object.keys(inputs).sort()).toEqual(microsoft365ToolCatalog.map(t => t.name).sort());
  });

  for (const entry of microsoft365ToolCatalog) {
    it(`${entry.name}: every recorded request is in its upstream list`, async () => {
      // The env gateway with every guard open, so each tool runs its whole path.
      const requests: Request[] = [];
      const gateway = createMicrosoft365Gateway({
        accessToken: TOKEN,
        tenantId: '',
        clientId: '',
        clientSecret: '',
        refreshToken: '',
        baseUrl: 'https://graph.microsoft.com/v1.0',
        protection: { allowPasswordReset: true, allowTap: true, protectedUsers: [], protectedGroups: [] },
        fetchImpl: async (input, init) => {
          const req = new Request(input, init);
          requests.push(req);
          return directory(req);
        },
      });
      await gateway.callTool(entry.name, inputs[entry.name] as never).catch(() => undefined);
      expect(requests.length, entry.name).toBeGreaterThan(0);
      for (const req of requests) {
        const path = new URL(req.url).pathname.replace(/^\/v1\.0/, '');
        const listed = entry.upstream.some(op => op.method === req.method && pattern(op.path).test(path));
        expect(listed, `${entry.name}: ${req.method} ${path}`).toBe(true);
      }
    });
  }
});
