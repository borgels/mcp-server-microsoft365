// Failure metadata on gateway error results (host mode): Microsoft Graph's HTTP status and its request id (a UUID or
// hex id only), for a host that counts failures without reading their text. Never the body, Graph's code or message.
import { afterEach, describe, expect, it, vi } from 'vitest';
import { failureMetadata, GraphHttpError, requestIdFrom, vendorRequestId } from '../src/errors.js';
import { createMicrosoft365Gateway } from '../src/gateway.js';

const TOKEN = 'host-token-0123456789abcdef';
const UUID = '3f2b8c1e-9a4d-4e6f-8b7a-1c2d3e4f5a6b';
const BODY = { error: { code: 'E_NOPE', message: 'Customer Jane Doe not found' } };

afterEach(() => {
  vi.unstubAllEnvs();
});

function hostAnswering(status: number, headers: Record<string, string>, body: unknown = BODY, accessToken = TOKEN) {
  return createMicrosoft365Gateway({
    accessToken,
    policy: {},
    fetchImpl: async () =>
      // A 304 has no body.
      new Response(status === 304 ? null : JSON.stringify(body), {
        status,
        headers: { 'content-type': 'application/json', ...headers },
      }),
  });
}

describe('vendorRequestId', () => {
  it('keeps a UUID and a hex id', () => {
    expect(vendorRequestId(UUID)).toBe(UUID);
    expect(vendorRequestId(` ${UUID} `)).toBe(UUID);
    expect(vendorRequestId('a1b2c3d4e5f60718')).toBe('a1b2c3d4e5f60718');
    expect(vendorRequestId('ABCDEF12-3456')).toBe('ABCDEF12-3456');
    expect(vendorRequestId('f'.repeat(64))).toBe('f'.repeat(64));
  });

  it('drops anything that could be text, and the shape of a personal number', () => {
    for (const value of [
      undefined,
      null,
      '',
      'abc123',
      'Jane Doe',
      'jane@example.com',
      'req_8f3a9c2d1e',
      'deadbeef deadbeef',
      '0101901234',
      'ab-010190-1234-cd',
      'f'.repeat(65),
      '-abcdef12',
      'abcdef12-',
      'abc--def',
      'abcdef12-ab',
      'abcdef1',
    ]) {
      expect(vendorRequestId(value)).toBeUndefined();
    }
  });

  it("reads Graph's request-id first, then the first other header with an id in the right form", () => {
    const other = 'deadbeefcafe0011';
    expect(requestIdFrom(new Headers({ 'x-ms-request-id': other, 'request-id': UUID }))).toBe(UUID);
    expect(requestIdFrom(new Headers({ 'request-id': 'not an id', 'client-request-id': other }))).toBe(other);
    expect(requestIdFrom(new Headers({ 'x-vendor-id': UUID }))).toBeUndefined();
  });

  it('gives metadata for an HTTP error only, keeping the message as it was', () => {
    const error = new GraphHttpError({ status: 404, method: 'GET', url: 'https://graph.microsoft.com/v1.0/users/x', requestId: UUID });
    expect(failureMetadata(error)).toEqual({ httpStatus: 404, requestId: UUID });
    expect(error.message).toContain(`request-id=${UUID}`);
    const odd = new GraphHttpError({ status: 500, method: 'GET', url: 'https://graph.microsoft.com/v1.0/users/x', requestId: 'Jane Doe' });
    expect(failureMetadata(odd)).toEqual({ httpStatus: 500 });
    expect(failureMetadata(new Error('boom'))).toBeUndefined();
  });
});

describe('host-mode error results', () => {
  it('carry the HTTP status and the request id, never the body', async () => {
    for (const status of [400, 403, 404, 429, 500]) {
      const result = await hostAnswering(status, { 'request-id': UUID }).callTool('get_user', { user: 'a@example.com' });
      expect(result.isError).toBe(true);
      expect(result.failure).toEqual({ httpStatus: status, requestId: UUID });
      const serialised = JSON.stringify(result);
      expect(serialised).not.toContain('Jane Doe');
      expect(serialised).not.toContain('E_NOPE');
    }
  });

  it('carry the status and request id of a redirect, which is not followed', async () => {
    for (const status of [301, 302, 304, 307, 308]) {
      const result = await hostAnswering(status, { location: 'https://elsewhere.example/', 'request-id': UUID }).callTool('get_user', {
        user: 'a@example.com',
      });
      expect(result.isError).toBe(true);
      expect(result.failure).toEqual({ httpStatus: status, requestId: UUID });
      expect(JSON.stringify(result)).not.toContain('elsewhere');
    }
  });

  it('carry the status alone when the request id is not in the form of one', async () => {
    const result = await hostAnswering(403, { 'request-id': 'Jane Doe <jane@example.com>' }).callTool('get_user', {
      user: 'a@example.com',
    });
    expect(result.failure).toEqual({ httpStatus: 403 });
    expect(JSON.stringify(result)).not.toContain('jane@example.com');
  });

  it('drop a request id that holds a piece of the token', async () => {
    const result = await hostAnswering(500, { 'request-id': 'deadbeef00112233' }, {}, 'host-token-deadbeef00112233-x').callTool('get_user', {
      user: 'a@example.com',
    });
    expect(result.isError).toBe(true);
    expect(result.failure).toEqual({ httpStatus: 500 });
    expect(JSON.stringify(result)).not.toContain('deadbeef0011');
  });

  it('keep the status and request id when the error body cannot be read', async () => {
    const broken = () =>
      new ReadableStream<Uint8Array>({
        pull(controller) {
          controller.error(new TypeError('terminated'));
        },
      });
    const gw = createMicrosoft365Gateway({
      accessToken: TOKEN,
      policy: {},
      fetchImpl: async () => new Response(broken(), { status: 503, headers: { 'content-type': 'application/json', 'request-id': UUID } }),
    });
    const result = await gw.callTool('get_user', { user: 'a@example.com' });
    expect(result.isError).toBe(true);
    expect(result.failure).toEqual({ httpStatus: 503, requestId: UUID });
  });

  it('carry the metadata of a failed directory lookup behind a refusal', async () => {
    const gw = createMicrosoft365Gateway({
      accessToken: TOKEN,
      policy: { enableWrites: true },
      fetchImpl: async () =>
        new Response(JSON.stringify(BODY), { status: 503, headers: { 'content-type': 'application/json', 'request-id': UUID } }),
    });
    for (const [tool, input] of [
      ['add_group_member', { groupId: 'g1', userId: 'u1' }],
      ['remove_group_member', { groupId: 'g1', userId: 'u1' }],
    ] as const) {
      const result = await gw.callTool(tool, input);
      expect(result.isError).toBe(true);
      expect(result.content[0]?.text).toBe('Refused: the target is protected.');
      expect(result.failure).toEqual({ httpStatus: 503, requestId: UUID });
      expect(JSON.stringify(result)).not.toContain('Jane Doe');
    }
    // A refusal by the policy itself has none.
    const refused = await gw.callTool('reset_password', { user: 'a@example.com' });
    expect(refused.isError).toBe(true);
    expect(refused.failure).toBeUndefined();
  });

  it('have no metadata when the cause is not Microsoft Graph', async () => {
    const gw = hostAnswering(200, {}, {});
    expect((await gw.callTool('get_user', {})).failure).toBeUndefined();
    expect((await gw.callTool('no_such_tool', {})).failure).toBeUndefined();
    const offline = createMicrosoft365Gateway({
      accessToken: TOKEN,
      policy: {},
      fetchImpl: async () => {
        throw new TypeError('fetch failed');
      },
    });
    const result = await offline.callTool('get_user', { user: 'a@example.com' });
    expect(result.isError).toBe(true);
    expect(result.failure).toBeUndefined();
  });

  it('are absent on a successful result', async () => {
    const result = await hostAnswering(200, { 'request-id': UUID }, { id: 'u1', displayName: 'Test' }).callTool('get_user', {
      user: 'a@example.com',
    });
    expect(result.isError).toBeUndefined();
    expect(result.failure).toBeUndefined();
  });
});
