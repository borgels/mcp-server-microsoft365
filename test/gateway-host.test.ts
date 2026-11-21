import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  GatewayConfigError,
  createMicrosoft365Gateway,
  microsoft365GatewayTools,
  microsoft365ToolCatalog,
} from '../src/gateway.js';

const TOKEN = 'host-token-0123456789abcdef';

function host(handler: (req: Request) => Response | Promise<Response>, policy: { enableWrites?: boolean; disabledTools?: string[] } = {}) {
  const requests: Request[] = [];
  const gateway = createMicrosoft365Gateway({
    accessToken: TOKEN,
    policy,
    fetchImpl: async (input, init) => {
      const req = new Request(input, init);
      requests.push(req);
      return handler(req);
    },
  });
  return { gateway, requests };
}

afterEach(() => {
  vi.unstubAllEnvs();
});

describe('microsoft365ToolCatalog', () => {
  it('has every gateway tool, with its Graph requests, none broad, deeply frozen', () => {
    expect(microsoft365ToolCatalog.map(entry => entry.name)).toEqual(microsoft365GatewayTools.map(tool => tool.name));
    for (const entry of microsoft365ToolCatalog) {
      expect(entry.upstream.length).toBeGreaterThan(0);
      expect(entry.broad).toBe(false);
    }
    expect(Object.isFrozen(microsoft365ToolCatalog[0]!.upstream[0])).toBe(true);
    const reads = microsoft365ToolCatalog.filter(entry => entry.riskLevel === 'read');
    for (const entry of reads) expect(entry.upstream.every(op => op.method === 'GET')).toBe(true);
  });
});

describe('host mode', () => {
  it('requires an access token and refuses unknown disabled tools', () => {
    expect(() => createMicrosoft365Gateway({ policy: {} })).toThrow(GatewayConfigError);
    expect(() => createMicrosoft365Gateway({ accessToken: TOKEN, policy: { disabledTools: ['nope'] } })).toThrow(GatewayConfigError);
  });

  it('reads nothing from the environment: Graph at its own address, the host token, protection all refused', async () => {
    vi.stubEnv('MS_GRAPH_BASE_URL', 'https://elsewhere.example.test/v1.0');
    vi.stubEnv('MS_ACCESS_TOKEN', 'env-token-should-not-be-used');
    vi.stubEnv('MS_ENABLE_PASSWORD_RESET', 'true');
    const { gateway, requests } = host(() => Response.json({ value: [] }), { enableWrites: true });
    await gateway.callTool('list_users', {});
    expect(new URL(requests[0]!.url).origin).toBe('https://graph.microsoft.com');
    expect(requests[0]!.headers.get('authorization')).toBe(`Bearer ${TOKEN}`);
    const reset = await gateway.callTool('reset_password', { user: 'someone@example.com' });
    expect(reset.isError).toBe(true);
  });

  it('offers only reads unless writes are enabled, and leaves out disabled tools', async () => {
    const { gateway, requests } = host(() => Response.json({}));
    expect(gateway.tools.every(tool => tool.riskLevel === 'read')).toBe(true);
    const refused = await gateway.callTool('update_user', { user: 'a@example.com', jobTitle: 'x' });
    expect(refused.isError).toBe(true);
    expect(requests).toHaveLength(0);
    const writes = host(() => Response.json({}), { enableWrites: true, disabledTools: ['reset_password'] });
    expect(writes.gateway.tools.some(tool => tool.name === 'update_user')).toBe(true);
    expect(writes.gateway.tools.some(tool => tool.name === 'reset_password')).toBe(false);
  });

  it('refuses unknown fields and missing required fields before any request', async () => {
    const { gateway, requests } = host(() => Response.json({}));
    expect((await gateway.callTool('get_user', {})).isError).toBe(true);
    expect((await gateway.callTool('get_user', { user: 'a', extra: 1 } as never)).isError).toBe(true);
    expect((await gateway.callTool('list_users', { top: '5' } as never)).isError).toBe(true);
    expect(requests).toHaveLength(0);
  });

  it('gives a Graph error as a fixed text with only the status', async () => {
    const { gateway } = host(() => Response.json({ error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges secret-ish' } }, { status: 403 }));
    const result = await gateway.callTool('list_groups', {});
    expect(result).toEqual({
      isError: true,
      content: [{ type: 'text', text: 'Microsoft Graph request failed with HTTP 403' }],
      failure: { httpStatus: 403 },
    });
  });

  it('withholds an answer that holds the token', async () => {
    const { gateway } = host(() => Response.json({ id: '1', note: TOKEN }));
    const result = await gateway.callTool('get_user', { user: 'a@example.com' });
    expect(result.isError).toBe(true);
    expect(JSON.stringify(result)).not.toContain(TOKEN);
  });
});
