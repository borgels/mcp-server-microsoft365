import { GraphHttpError } from '../errors.js';

export const GRAPH_DEFAULT_BASE_URL = 'https://graph.microsoft.com/v1.0';
export const GRAPH_DEFAULT_AUTHORITY_HOST = 'https://login.microsoftonline.com';
export const GRAPH_DEFAULT_SCOPE = 'https://graph.microsoft.com/.default';

export type GraphHttpMethod = 'GET' | 'POST' | 'PUT' | 'PATCH' | 'DELETE';
export type GraphQueryValue = string | number | boolean | null | undefined | Array<string | number | boolean>;

export interface GraphClientOptions {
  tenantId?: string;
  clientId?: string;
  clientSecret?: string;
  /** Pre-fetched bearer token; short-circuits all grants. */
  accessToken?: string;
  /**
   * Delegated OAuth refresh token. When set (and no static accessToken), the
   * client mints short-lived DELEGATED access tokens via the refresh_token grant
   * instead of client_credentials — so calls act on-behalf-of the user who
   * consented, bounded by that user's own roles.
   */
  refreshToken?: string;
  baseUrl?: string;
  authorityHost?: string;
  scope?: string;
  fetchImpl?: typeof fetch;
  timeoutMs?: number;
}

export interface GraphRequest {
  method: GraphHttpMethod;
  path: string;
  query?: Record<string, GraphQueryValue>;
  body?: unknown;
  headers?: Record<string, string>;
}

export interface GraphResponse<T = unknown> {
  data: T;
  nextLink?: string;
  status: number;
}

interface OAuthToken {
  access_token?: string;
  token_type?: string;
  expires_in?: number;
  ext_expires_in?: number;
  refresh_token?: string;
}

export class GraphClient {
  readonly baseUrl: string;
  private readonly tenantId?: string;
  private readonly clientId?: string;
  private readonly clientSecret?: string;
  private readonly staticAccessToken?: string;
  private readonly refreshToken?: string;
  private readonly authorityHost: string;
  private readonly scope: string;
  private readonly fetchImpl: typeof fetch;
  private readonly timeoutMs: number;
  private cachedToken?: { accessToken: string; expiresAt: number };
  /** Most recent (possibly rotated) delegated refresh token; callers persist this. */
  latestRefreshToken?: string;

  constructor(options: GraphClientOptions = {}) {
    this.tenantId = options.tenantId ?? process.env.MS_TENANT_ID;
    this.clientId = options.clientId ?? process.env.MS_CLIENT_ID;
    this.clientSecret = options.clientSecret ?? process.env.MS_CLIENT_SECRET;
    this.staticAccessToken = options.accessToken ?? process.env.MS_ACCESS_TOKEN;
    this.refreshToken = options.refreshToken ?? process.env.MS_REFRESH_TOKEN;
    this.latestRefreshToken = this.refreshToken;
    this.baseUrl = trimTrailingSlash(options.baseUrl ?? process.env.MS_GRAPH_BASE_URL ?? GRAPH_DEFAULT_BASE_URL);
    assertSafeBaseUrl(this.baseUrl);
    this.authorityHost = trimTrailingSlash(
      options.authorityHost ?? process.env.MS_AUTHORITY_HOST ?? GRAPH_DEFAULT_AUTHORITY_HOST,
    );
    this.scope = options.scope ?? process.env.MS_GRAPH_SCOPE ?? GRAPH_DEFAULT_SCOPE;
    this.fetchImpl = options.fetchImpl ?? fetch;
    this.timeoutMs = options.timeoutMs ?? Number(process.env.MS_TIMEOUT_MS ?? 30_000);
  }

  async get<T = unknown>(path: string, query?: Record<string, GraphQueryValue>): Promise<GraphResponse<T>> {
    return this.request<T>({ method: 'GET', path, query });
  }

  async post<T = unknown>(
    path: string,
    body?: unknown,
    query?: Record<string, GraphQueryValue>,
  ): Promise<GraphResponse<T>> {
    return this.request<T>({ method: 'POST', path, query, body });
  }

  async patch<T = unknown>(
    path: string,
    body?: unknown,
    query?: Record<string, GraphQueryValue>,
  ): Promise<GraphResponse<T>> {
    return this.request<T>({ method: 'PATCH', path, query, body });
  }

  async delete<T = unknown>(
    path: string,
    body?: unknown,
    query?: Record<string, GraphQueryValue>,
  ): Promise<GraphResponse<T>> {
    return this.request<T>({ method: 'DELETE', path, query, body });
  }

  async request<T = unknown>(request: GraphRequest): Promise<GraphResponse<T>> {
    const url = this.url(request.path, request.query);
    const headers: Record<string, string> = {
      Accept: 'application/json',
      Authorization: `Bearer ${await this.accessToken()}`,
      ...request.headers,
    };

    if (request.body !== undefined) {
      headers['Content-Type'] = 'application/json';
    }

    const response = await this.fetchImpl(url, {
      method: request.method,
      headers,
      body: request.body === undefined ? undefined : JSON.stringify(request.body),
      signal: AbortSignal.timeout(this.timeoutMs),
    });

    const responseBody = await readResponseBody(response);

    if (!response.ok) {
      throw new GraphHttpError({
        status: response.status,
        method: request.method,
        url,
        payload: responseBody,
        retryAfter: response.headers.get('retry-after') ?? undefined,
        requestId:
          response.headers.get('request-id') ??
          response.headers.get('x-ms-request-id') ??
          response.headers.get('client-request-id') ??
          undefined,
        fallbackMessage: typeof responseBody === 'string' ? responseBody : undefined,
      });
    }

    return {
      data: responseBody as T,
      nextLink: extractNextLink(responseBody),
      status: response.status,
    };
  }

  /** URL for a `directoryObjects/{id}` reference, used by `$ref` membership calls. */
  directoryObjectUrl(id: string): string {
    return `${this.baseUrl}/directoryObjects/${encodeURIComponent(id)}`;
  }

  async getTokenMetadata(): Promise<{ authMethod: string; tenantId?: string; expiresAt?: string }> {
    if (this.staticAccessToken) {
      return { authMethod: 'access_token', tenantId: this.tenantId };
    }

    const token = await this.fetchClientCredentialsToken();
    return {
      authMethod: 'client_credentials',
      tenantId: this.tenantId,
      expiresAt: new Date(token.expiresAt).toISOString(),
    };
  }

  private async accessToken(): Promise<string> {
    if (this.staticAccessToken) {
      return this.staticAccessToken;
    }

    if (this.refreshToken) {
      return (await this.fetchDelegatedToken()).accessToken;
    }

    return (await this.fetchClientCredentialsToken()).accessToken;
  }

  /**
   * Mint a short-lived DELEGATED access token via the refresh_token grant
   * (confidential client). Acts on-behalf-of the consenting user, bounded by
   * their roles. Entra rotates the refresh token on redemption; the newest one
   * is exposed via {@link latestRefreshToken} so the caller can persist it.
   */
  private async fetchDelegatedToken(): Promise<{ accessToken: string; expiresAt: number }> {
    const now = Date.now();
    if (this.cachedToken && this.cachedToken.expiresAt - now > 60_000) {
      return this.cachedToken;
    }

    if (!this.tenantId || !this.clientId || !this.clientSecret || !this.latestRefreshToken) {
      throw new Error(
        'Missing delegated credentials. Set MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET and a refresh token.',
      );
    }

    const url = `${this.authorityHost}/${encodeURIComponent(this.tenantId)}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      client_id: this.clientId,
      client_secret: this.clientSecret,
      grant_type: 'refresh_token',
      refresh_token: this.latestRefreshToken,
      scope: this.scope,
    });

    const response = await this.fetchImpl(url, {
      method: 'POST',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: body.toString(),
      signal: AbortSignal.timeout(this.timeoutMs),
    });

    const responseBody = await readResponseBody(response);
    if (!response.ok) {
      throw new GraphHttpError({
        status: response.status,
        method: 'POST',
        url,
        payload: responseBody,
        fallbackMessage: typeof responseBody === 'string' ? responseBody : undefined,
      });
    }

    const token = responseBody as OAuthToken;
    if (!token.access_token) {
      throw new Error('Microsoft delegated token response did not include access_token.');
    }
    // Persist the rotated refresh token for the next mint / caller persistence.
    if (token.refresh_token) {
      this.latestRefreshToken = token.refresh_token;
    }

    this.cachedToken = {
      accessToken: token.access_token,
      expiresAt: now + Math.max(1, Number(token.expires_in ?? 3600)) * 1000,
    };
    return this.cachedToken;
  }

  private async fetchClientCredentialsToken(): Promise<{ accessToken: string; expiresAt: number }> {
    const now = Date.now();
    if (this.cachedToken && this.cachedToken.expiresAt - now > 60_000) {
      return this.cachedToken;
    }

    if (!this.tenantId || !this.clientId || !this.clientSecret) {
      throw new Error(
        'Missing Microsoft Graph credentials. Set MS_TENANT_ID, MS_CLIENT_ID and MS_CLIENT_SECRET, or set MS_ACCESS_TOKEN.',
      );
    }

    const url = `${this.authorityHost}/${encodeURIComponent(this.tenantId)}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      client_id: this.clientId,
      client_secret: this.clientSecret,
      grant_type: 'client_credentials',
      scope: this.scope,
    });

    const response = await this.fetchImpl(url, {
      method: 'POST',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: body.toString(),
      signal: AbortSignal.timeout(this.timeoutMs),
    });

    const responseBody = await readResponseBody(response);
    if (!response.ok) {
      throw new GraphHttpError({
        status: response.status,
        method: 'POST',
        url,
        payload: responseBody,
        fallbackMessage: typeof responseBody === 'string' ? responseBody : undefined,
      });
    }

    const token = responseBody as OAuthToken;
    if (!token.access_token) {
      throw new Error('Microsoft Graph token response did not include access_token.');
    }

    this.cachedToken = {
      accessToken: token.access_token,
      expiresAt: now + Math.max(1, Number(token.expires_in ?? 3600)) * 1000,
    };
    return this.cachedToken;
  }

  private url(path: string, query: Record<string, GraphQueryValue> = {}): string {
    if (/^https?:\/\//i.test(path)) {
      // Support absolute @odata.nextLink follow-through.
      const absolute = new URL(path);
      for (const [key, value] of Object.entries(query)) {
        appendQueryValue(absolute.searchParams, key, value);
      }
      return absolute.toString();
    }

    const normalizedPath = path.startsWith('/') ? path : `/${path}`;
    const url = new URL(`${this.baseUrl}${normalizedPath}`);

    for (const [key, value] of Object.entries(query)) {
      appendQueryValue(url.searchParams, key, value);
    }

    return url.toString();
  }
}

async function readResponseBody(response: Response): Promise<unknown> {
  if (response.status === 204) {
    return null;
  }

  const text = await response.text();
  if (!text) {
    return null;
  }

  const contentType = response.headers.get('content-type') ?? '';
  if (!contentType.includes('application/json')) {
    return text;
  }

  try {
    return JSON.parse(text) as unknown;
  } catch {
    return text;
  }
}

function extractNextLink(body: unknown): string | undefined {
  if (!body || typeof body !== 'object') {
    return undefined;
  }
  const next = (body as Record<string, unknown>)['@odata.nextLink'];
  return typeof next === 'string' ? next : undefined;
}

function appendQueryValue(searchParams: URLSearchParams, key: string, value: GraphQueryValue): void {
  if (value === undefined || value === null || value === '') {
    return;
  }

  if (Array.isArray(value)) {
    for (const item of value) {
      appendQueryValue(searchParams, key, item);
    }
    return;
  }

  searchParams.set(key, String(value));
}

function trimTrailingSlash(value: string): string {
  return value.replace(/\/+$/, '');
}

function assertSafeBaseUrl(baseUrl: string): void {
  let parsed: URL;
  try {
    parsed = new URL(baseUrl);
  } catch {
    throw new Error(`MS_GRAPH_BASE_URL is not a valid URL: ${baseUrl}`);
  }

  if (parsed.protocol === 'https:') {
    return;
  }

  if (parsed.protocol === 'http:' && isLocalHost(parsed.hostname)) {
    return;
  }

  throw new Error(
    `Refusing to send Microsoft Graph credentials over ${parsed.protocol}//. Use https:// (loopback http:// is allowed for local mocks).`,
  );
}

function isLocalHost(hostname: string): boolean {
  return hostname === 'localhost' || hostname === '127.0.0.1' || hostname === '::1';
}
