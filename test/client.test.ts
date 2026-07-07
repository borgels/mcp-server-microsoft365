import { afterEach, describe, expect, it, vi } from 'vitest';
import { formatUnknownError, GraphHttpError, redactSecrets } from '../src/errors.js';
import { GraphClient } from '../src/graph/client.js';

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

    await expect(client.get('/users')).rejects.toMatchObject({
      status: 429,
      retryAfter: '30',
      requestId: 'req-123',
    });
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
});
