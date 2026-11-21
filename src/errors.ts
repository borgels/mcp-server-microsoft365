export interface GraphHttpErrorInput {
  status: number;
  method: string;
  url: string;
  payload?: unknown;
  retryAfter?: string;
  requestId?: string;
  fallbackMessage?: string;
}

/** Upstream error text included in a message is cut to this length (after redaction). */
export const MAX_ERROR_TEXT = 500;

const SECRET_PATTERNS = [
  /Authorization:\s*(Bearer|Basic)\s+[^,\s}]+/gi,
  /(MS_TENANT_ID|MS_CLIENT_ID|MS_CLIENT_SECRET|MS_ACCESS_TOKEN|client_secret|clientSecret|access_token|refresh_token|apiKey|accessToken|token)["']?\s*[:=]\s*["']?[^"',\s}]+/gi,
  /\b(Bearer|Basic)\s+[A-Za-z0-9._~+/-]+=*/g,
];

export class GraphHttpError extends Error {
  readonly status: number;
  readonly method: string;
  readonly url: string;
  readonly payload?: unknown;
  readonly retryAfter?: string;
  /**
   * The vendor's own id for the request, only in the form such ids have (a UUID or hex id, see requestIdFrom), for
   * the failure metadata. The message keeps the header's value as before.
   */
  readonly requestId?: string;

  constructor(input: GraphHttpErrorInput) {
    super(formatGraphHttpError(input));
    this.name = 'GraphHttpError';
    this.status = input.status;
    this.method = input.method;
    this.url = redactSecrets(input.url);
    this.payload = input.payload;
    this.retryAfter = input.retryAfter;
    this.requestId = vendorRequestId(input.requestId);
  }
}

/**
 * Response headers that carry the vendor's own id for a request, most specific first. Add the vendor's documented
 * header in front when it has one.
 */
export const REQUEST_ID_HEADERS: readonly string[] = [
  // Microsoft Graph's own ids for a request, then Microsoft Entra's token endpoint's.
  'request-id',
  'client-request-id',
  'x-ms-request-id',
  'x-request-id',
  'x-correlation-id',
  'correlation-id',
  'x-amzn-requestid',
];

/**
 * A hex id: one run of hex digits, or groups of at least four hex digits joined by single hyphens (a UUID is one),
 * 8 to 64 characters in all. A name, a sentence or an e-mail address never fits.
 */
const REQUEST_ID_FORM = /^[0-9A-Fa-f]{4,}(?:-[0-9A-Fa-f]{4,})*$/;
/** Six digits and four more: the shape of a Danish personal number. An id that holds it is dropped. */
const PERSONAL_NUMBER_SHAPE = /\d{6}-?\d{4}/;

/**
 * The vendor's request id, only in the form such ids have (a UUID or a hex id of 8 to 64 characters), or
 * undefined. Anything else, and an id that holds the shape of a personal number, is dropped: the value comes from
 * the vendor and is kept only when it cannot be text.
 */
export function vendorRequestId(value: string | null | undefined): string | undefined {
  if (typeof value !== 'string') return undefined;
  const id = value.trim();
  return id.length >= 8 && id.length <= 64 && REQUEST_ID_FORM.test(id) && !PERSONAL_NUMBER_SHAPE.test(id) ? id : undefined;
}

/** The vendor's request id from the first of REQUEST_ID_HEADERS a response has in the right form, or undefined. */
export function requestIdFrom(headers: Headers, names: readonly string[] = REQUEST_ID_HEADERS): string | undefined {
  for (const name of names) {
    const id = vendorRequestId(headers.get(name));
    if (id !== undefined) return id;
  }
  return undefined;
}

/**
 * Metadata about a failed call, for a host that counts failures: the HTTP status the upstream API answered with
 * and its request id. Never the response body, the vendor's error code or message, or the input.
 */
export interface FailureMetadata {
  /** The HTTP status of the upstream API's answer (100-599). */
  httpStatus: number;
  /** The vendor's own id for the request, as a UUID or hex id only. */
  requestId?: string;
}

/** Failure metadata from a status and a request id: undefined unless the status is 100-599; an id out of form is left out. */
export function failureFrom(status: unknown, requestId: unknown): FailureMetadata | undefined {
  if (typeof status !== 'number' || !Number.isInteger(status) || status < 100 || status > 599) return undefined;
  const id = typeof requestId === 'string' ? vendorRequestId(requestId) : undefined;
  return id === undefined ? { httpStatus: status } : { httpStatus: status, requestId: id };
}

/**
 * The failure metadata of an error: an HTTP error from Microsoft Graph or its token endpoint gives its status and
 * request id. Never the response body, Graph's error code or message.
 */
export function failureMetadata(error: unknown): FailureMetadata | undefined {
  if (!(error instanceof GraphHttpError)) return undefined;
  return failureFrom(error.status, error.requestId);
}

export function formatUnknownError(error: unknown): string {
  if (error instanceof Error) {
    return redactSecrets(error.message);
  }

  return redactSecrets(String(error));
}

export function redactSecrets(value: string): string {
  return SECRET_PATTERNS.reduce(
    (current, pattern) =>
      current.replace(pattern, match => {
        if (/^(Bearer|Basic)\s/i.test(match)) {
          const scheme = match.split(/\s+/)[0] ?? 'Authorization';
          return `${scheme} [REDACTED]`;
        }

        if (/^Authorization:/i.test(match)) {
          return 'Authorization: [REDACTED]';
        }

        const separator = match.includes(':') ? ':' : '=';
        const key = match.split(separator)[0]?.trim() ?? 'secret';
        return `${key}${separator} [REDACTED]`;
      }),
    value,
  );
}

function formatGraphHttpError(input: GraphHttpErrorInput): string {
  const parts = [
    `Microsoft Graph request failed with HTTP ${input.status}`,
    `${input.method.toUpperCase()} ${redactSecrets(input.url)}`,
    input.requestId ? `request-id=${input.requestId}` : undefined,
    input.retryAfter ? `retry-after=${input.retryAfter}s` : undefined,
    truncateRedacted(graphErrorText(input.payload)),
    truncateRedacted(input.fallbackMessage),
  ].filter(Boolean);

  return redactSecrets(parts.join(' | '));
}

/** Redacts first, then truncates, so a cut can never expose part of a secret. */
function truncateRedacted(value: string | undefined): string | undefined {
  if (!value) {
    return value;
  }
  const redacted = redactSecrets(value);
  return redacted.length > MAX_ERROR_TEXT ? `${redacted.slice(0, MAX_ERROR_TEXT)}... [truncated]` : redacted;
}

function graphErrorText(payload: unknown): string | undefined {
  if (!payload || typeof payload !== 'object') {
    return undefined;
  }

  const value = payload as Record<string, unknown>;

  // Microsoft Graph shape: { error: { code, message, innerError } }
  const graphError = value.error;
  if (graphError && typeof graphError === 'object') {
    const inner = graphError as Record<string, unknown>;
    const code = typeof inner.code === 'string' ? inner.code : undefined;
    const message = typeof inner.message === 'string' ? inner.message : undefined;
    const combined = [code, message].filter(Boolean).join(': ');
    if (combined) {
      return combined;
    }
  }

  const message =
    value.message ??
    value.error ??
    value.error_description ??
    value.errorMessage ??
    value.title ??
    (Array.isArray(value.errors) ? value.errors.map(String).join(', ') : undefined);

  return typeof message === 'string' ? message : undefined;
}
