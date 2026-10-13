/**
 * Guard for caller-supplied values that end up inside a Graph URL path.
 *
 * `encodeURIComponent` already escapes `/`, `?` and `#`, but it leaves `.` and
 * `..` untouched, and URL normalisation resolves those as path segments. Values
 * that are already percent-encoded (`%2e`, `%2f`, `%5c`) could also be decoded
 * by an intermediary. None of these are valid object ids or userPrincipalNames
 * we need to support, so they are rejected outright instead of being escaped.
 *
 * Note: guest UPNs contain `#EXT#`; address guests by object id instead.
 */
const FORBIDDEN_CHARS = /[/\\?#]/;
const FORBIDDEN_ENCODED = /%(2e|2f|5c)/i;

export function assertSafePathParam(value: string, name = 'value'): string {
  if (typeof value !== 'string') {
    throw new Error(`Invalid ${name}: expected a string.`);
  }
  const trimmed = value.trim();
  if (
    trimmed === '' ||
    trimmed === '.' ||
    trimmed === '..' ||
    FORBIDDEN_CHARS.test(trimmed) ||
    FORBIDDEN_ENCODED.test(trimmed) ||
    // Control characters have no place in an identifier.
    hasControlChars(trimmed)
  ) {
    throw new Error(
      `Invalid ${name}: must be a non-empty identifier without path separators, '?', '#', dot segments or percent-encoded separators.`,
    );
  }
  return trimmed;
}

/** Validate then percent-encode one URL path segment. */
export function pathSegment(value: string, name = 'value'): string {
  return encodeURIComponent(assertSafePathParam(value, name));
}

function hasControlChars(value: string): boolean {
  for (let i = 0; i < value.length; i += 1) {
    const code = value.charCodeAt(i);
    if (code < 0x20 || code === 0x7f) return true;
  }
  return false;
}
