import { afterEach, describe, expect, it, vi } from 'vitest';
import { formatUnknownError, GraphHttpError, redactSecrets } from '../src/errors.js';
import { ConfigError } from '../src/config-env.js';
import { GraphClient, graphEnvTimeoutMs } from '../src/graph/client.js';

const originalEnv = { ...process.env };

describe('GraphClient', () => {
  afterEach(() => {
    process.env = { ...originalEnv };
  });

  it('fetches and caches OAuth client credentials tokens', async () => {
    const fetchMock = vi.fn<typeof fetch>(async input => {
      const url = String(input);
      if (url.endsWith('/oauth2/v2.0/token')) {
        return Response.json({ access_token: 'token-1', token_type: 'Bearer', expires_in: 3600 });
      }
      return Response.json({ value: [] });
    });
    const client = new GraphClient({
      tenantId: 'contoso.onmicrosoft.com',
      clientId: 'client-id',
      clientSecret: 'client-secret',
      baseUrl: 'https://graph.example.test/v1.0',
      authorityHost: 'https://login.example.test',
      fetchImpl: fetchMock as unknown as typeof fetch,
    });

    await client.get('/users', { $top: 20, $select: 'id,displayName' });
    await client.get('/groups');

    // token endpoint (once, cached) + two data calls
    expect(fetchMock).toHaveBeenCalledTimes(3);
    const tokenUrl = String(fetchMock.mock.calls[0]?.[0]);
    expect(tokenUrl).toBe('https://login.example.test/contoso.onmicrosoft.com/oauth2/v2.0/token');
    expect(fetchMock.mock.calls[0]?.[1]?.body).toContain('grant_type=client_credentials');
    expect(fetchMock.mock.calls[0]?.[1]?.body).toContain('scope=https%3A%2F%2Fgraph.microsoft.com%2F.default');
    expect(fetchMock.mock.calls[1]?.[1]?.headers).toMatchObject({
      Accept: 'application/json',
      Authorization: 'Bearer token-1',
    });
    expect(String(fetchMock.mock.calls[1]?.[0])).toContain('%24top=20');
  });

  it('can use externally managed access tokens without requesting OAuth tokens', async () => {
    const fetchMock = vi.fn<typeof fetch>(async () => Response.json({ id: 'user-1' }));
    const client = new GraphClient({
      accessToken: 'static-token',
      baseUrl: 'https://graph.example.test/v1.0',
      fetchImpl: fetchMock as unknown as typeof fetch,
    });

    await client.get('/users/user-1');

    expect(fetchMock).toHaveBeenCalledOnce();
    expect(fetchMock.mock.calls[0]?.[1]?.headers).toMatchObject({
      Authorization: 'Bearer static-token',
    });
  });

  it('rejects unsafe credential transport URLs', () => {
    expect(() => new GraphClient({ accessToken: 'token', baseUrl: 'http://graph.example.test' })).toThrow(/Refusing/);
    expect(() => new GraphClient({ accessToken: 'token', baseUrl: 'http://127.0.0.1:4010' })).not.toThrow();
  });

  it('surfaces Graph errors with retry-after and request-id, redacted', async () => {
    const fetchMock = vi.fn<typeof fetch>(async () =>
      Response.json(
        { error: { code: 'tooManyRequests', message: 'Bearer super-secret throttled' } },
        { status: 429, headers: { 'Retry-After': '30', 'request-id': 'req-123' } },
      ),
    );
    const client = new GraphClient({
      accessToken: 'static-token',
      baseUrl: 'https://graph.example.test/v1.0',
      fetchImpl: fetchMock as unknown as typeof fetch,
    });

    const error = (await client.get('/users').catch(caught => caught)) as GraphHttpError;
    expect(error).toMatchObject({ status: 429, retryAfter: '30' });
    // The message keeps the header's value; the field holds a request id only in the form of one (a UUID or hex id).
    expect(error.message).toContain('request-id=req-123');
    expect(error.requestId).toBeUndefined();
  });

  it('keeps a request id in the form of one, from the token endpoint too', async () => {
    const id = '3f2b8c1e-9a4d-4e6f-8b7a-1c2d3e4f5a6b';
    const graph = new GraphClient({
      accessToken: 'static-token',
      baseUrl: 'https://graph.example.test/v1.0',
      fetchImpl: (async () => Response.json({}, { status: 503, headers: { 'request-id': id } })) as typeof fetch,
    });
    await expect(graph.get('/users')).rejects.toMatchObject({ status: 503, requestId: id });
    const app = new GraphClient({
      tenantId: 'tenant-1',
      clientId: 'client-1',
      clientSecret: 'secret-value-1',
      baseUrl: 'https://graph.example.test/v1.0',
      fetchImpl: (async () =>
        Response.json({ error: 'invalid_client', error_description: 'Jane Doe' }, { status: 401, headers: { 'x-ms-request-id': id } })) as typeof fetch,
    });
    await expect(app.get('/users')).rejects.toMatchObject({ status: 401, requestId: id });
  });

  it('formats and redacts API errors', () => {
    const error = new GraphHttpError({
      status: 401,
      method: 'GET',
      url: 'https://graph.example.test/v1.0/users',
      payload: { error: { code: 'InvalidAuthentication', message: 'Bearer super-secret failed' } },
    });

    expect(error.message).toContain('Microsoft Graph request failed with HTTP 401');
    expect(error.message).toContain('InvalidAuthentication');
    expect(formatUnknownError(error)).not.toContain('super-secret');
    expect(redactSecrets('MS_CLIENT_SECRET=abc Authorization: Bearer xyz')).toContain('[REDACTED]');
    expect(redactSecrets('MS_CLIENT_SECRET=abc')).not.toContain('abc');
  });

  it('never follows redirects with credentials', async () => {
    const fetchMock = vi.fn<typeof fetch>(async (input, init) => {
      if (String(input).endsWith('/oauth2/v2.0/token')) {
        expect(init?.redirect).toBe('manual');
        return new Response(null, { status: 307, headers: { location: 'https://elsewhere.example/token' } });
      }
      expect(init?.redirect).toBe('manual');
      return new Response(null, { status: 302, headers: { location: 'https://elsewhere.example/users' } });
    });

    const tokenClient = new GraphClient({
      accessToken: 'static-token',
      baseUrl: 'https://graph.example.test/v1.0',
      fetchImpl: fetchMock as unknown as typeof fetch,
    });
    await expect(tokenClient.get('/users')).rejects.toThrow(/HTTP 302.*redirects are not followed/);

    const secretClient = new GraphClient({
      tenantId: 'tenant',
      clientId: 'client-id',
      clientSecret: 'client-secret',
      baseUrl: 'https://graph.example.test/v1.0',
      authorityHost: 'https://login.example.test',
      fetchImpl: fetchMock as unknown as typeof fetch,
    });
    await expect(secretClient.get('/users')).rejects.toThrow(/HTTP 307/);
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it('caps error bodies and redacts before truncating', async () => {
    const body = `client_secret=${'s'.repeat(40)} ${'x'.repeat(200_000)}`;
    const client = new GraphClient({
      accessToken: 'static-token',
      baseUrl: 'https://graph.example.test/v1.0',
      fetchImpl: (async () => new Response(body, { status: 500, headers: { 'content-type': 'text/plain' } })) as typeof fetch,
    });
    const error = (await client.get('/users').then(
      () => new Error('expected a rejection'),
      (caught: unknown) => caught as Error,
    ));
    expect(error.message).toContain('[REDACTED]');
    expect(error.message).not.toContain('sss');
    expect(error.message).toContain('[truncated]');
    expect(error.message.length).toBeLessThan(2_000);
  });

  it('validates MS_TIMEOUT_MS instead of producing NaN', () => {
    expect(graphEnvTimeoutMs({})).toBe(30_000);
    expect(graphEnvTimeoutMs({ MS_TIMEOUT_MS: '5000' })).toBe(5_000);
    for (const value of ['abc', '30s', '0', '-1']) {
      expect(() => graphEnvTimeoutMs({ MS_TIMEOUT_MS: value }), value).toThrow(ConfigError);
    }
    process.env.MS_TIMEOUT_MS = 'abc';
    expect(() => new GraphClient({ accessToken: 'x' })).toThrow(/MS_TIMEOUT_MS/);
  });
});
