import { randomBytes } from 'node:crypto';
import { GraphHttpError } from '../errors.js';
import type { GraphClient } from './client.js';
import { resolveSkuId, type SkuResolution, type SubscribedSku } from './licenses.js';

export interface CreateUserInput {
  displayName: string;
  userPrincipalName: string;
  mailNickname?: string;
  password?: string;
  usageLocation?: string;
  accountEnabled?: boolean;
  forceChangePasswordNextSignIn?: boolean;
  /** Defaults to true: GET the UPN first and fail clearly if it already exists. */
  checkExisting?: boolean;
}

export interface CreateUserResult {
  id: string;
  userPrincipalName: string;
  displayName: string;
  usageLocation?: string;
  temporaryPassword: string;
  generatedPassword: boolean;
}

export interface AssignLicenseInput {
  userId: string;
  license: string;
  usageLocation?: string;
  disabledPlans?: string[];
}

export interface RemoveLicenseInput {
  userId: string;
  license: string;
}

/**
 * Generate a strong temporary password: base64url random material plus a
 * complexity suffix guaranteeing upper, lower, digit, and symbol classes.
 */
export function generateTempPassword(): string {
  const material = randomBytes(18).toString('base64url').replace(/[-_]/g, '');
  return `${material}Aa1!`;
}

async function upnExists(client: GraphClient, userPrincipalName: string): Promise<boolean> {
  try {
    await client.get(`/users/${encodeURIComponent(userPrincipalName)}`, { $select: 'id' });
    return true;
  } catch (error) {
    if (error instanceof GraphHttpError && error.status === 404) {
      return false;
    }
    throw error;
  }
}

export async function createUser(client: GraphClient, input: CreateUserInput): Promise<CreateUserResult> {
  const userPrincipalName = input.userPrincipalName.trim();
  const displayName = input.displayName.trim();
  const mailNickname = (input.mailNickname?.trim() || defaultMailNickname(userPrincipalName));

  if (input.checkExisting !== false && (await upnExists(client, userPrincipalName))) {
    throw new Error(
      `A user with userPrincipalName "${userPrincipalName}" already exists. Choose a different UPN or update the existing user.`,
    );
  }

  const generatedPassword = !input.password;
  const password = input.password ?? generateTempPassword();

  const body: Record<string, unknown> = {
    accountEnabled: input.accountEnabled ?? true,
    displayName,
    mailNickname,
    userPrincipalName,
    passwordProfile: {
      password,
      forceChangePasswordNextSignIn: input.forceChangePasswordNextSignIn ?? true,
    },
  };

  if (input.usageLocation) {
    body.usageLocation = input.usageLocation;
  }

  const response = await client.post<{ id: string; userPrincipalName: string; displayName: string; usageLocation?: string }>(
    '/users',
    body,
  );

  const created = response.data;
  return {
    id: created.id,
    userPrincipalName: created.userPrincipalName ?? userPrincipalName,
    displayName: created.displayName ?? displayName,
    usageLocation: created.usageLocation ?? input.usageLocation,
    temporaryPassword: password,
    generatedPassword,
  };
}

export async function setUsageLocation(
  client: GraphClient,
  userId: string,
  usageLocation: string,
): Promise<{ userId: string; usageLocation: string }> {
  await client.patch(`/users/${encodeURIComponent(userId)}`, { usageLocation });
  return { userId, usageLocation };
}

export async function listSubscribedSkus(client: GraphClient): Promise<SubscribedSku[]> {
  const response = await client.get<{ value?: SubscribedSku[] }>('/subscribedSkus');
  return response.data?.value ?? [];
}

async function getUserUsageLocation(client: GraphClient, userId: string): Promise<string | undefined> {
  const response = await client.get<{ usageLocation?: string | null }>(`/users/${encodeURIComponent(userId)}`, {
    $select: 'usageLocation',
  });
  const location = response.data?.usageLocation;
  return location ? location : undefined;
}

export async function assignLicense(
  client: GraphClient,
  input: AssignLicenseInput,
): Promise<{ userId: string; assigned: SkuResolution; usageLocation: string }> {
  // usageLocation MUST be set before assigning a license, otherwise Graph
  // rejects the assignment with an "invalid usage location" error.
  let usageLocation = input.usageLocation?.trim();
  if (usageLocation) {
    await setUsageLocation(client, input.userId, usageLocation);
  } else {
    usageLocation = await getUserUsageLocation(client, input.userId);
    if (!usageLocation) {
      throw new Error(
        `Cannot assign a license: user ${input.userId} has no usageLocation set. Provide a usageLocation (ISO 3166-1 alpha-2, e.g. "DK") so it can be set before the license is assigned.`,
      );
    }
  }

  const skus = await listSubscribedSkus(client);
  const resolved = resolveSkuId(skus, input.license);
  if (!resolved) {
    throw new Error(
      `Could not resolve license "${input.license}" against the tenant's subscribed SKUs. Available SKUs: ${describeSkus(skus)}.`,
    );
  }

  if (resolved.available <= 0) {
    throw new Error(
      `No available seats for license "${resolved.skuPartNumber}" (enabled=${resolved.enabled}, consumed=${resolved.consumed}). Free a seat or buy more before assigning.`,
    );
  }

  await client.post(`/users/${encodeURIComponent(input.userId)}/assignLicense`, {
    addLicenses: [{ skuId: resolved.skuId, disabledPlans: input.disabledPlans ?? [] }],
    removeLicenses: [],
  });

  return { userId: input.userId, assigned: resolved, usageLocation };
}

export async function removeLicense(
  client: GraphClient,
  input: RemoveLicenseInput,
): Promise<{ userId: string; removed: SkuResolution }> {
  const skus = await listSubscribedSkus(client);
  const resolved = resolveSkuId(skus, input.license);
  if (!resolved) {
    throw new Error(
      `Could not resolve license "${input.license}" against the tenant's subscribed SKUs. Available SKUs: ${describeSkus(skus)}.`,
    );
  }

  await client.post(`/users/${encodeURIComponent(input.userId)}/assignLicense`, {
    addLicenses: [],
    removeLicenses: [resolved.skuId],
  });

  return { userId: input.userId, removed: resolved };
}

export async function addGroupMember(
  client: GraphClient,
  groupId: string,
  userId: string,
): Promise<{ groupId: string; userId: string; added: boolean; alreadyMember: boolean }> {
  try {
    await client.post(`/groups/${encodeURIComponent(groupId)}/members/$ref`, {
      '@odata.id': client.directoryObjectUrl(userId),
    });
    return { groupId, userId, added: true, alreadyMember: false };
  } catch (error) {
    if (isAlreadyMemberError(error)) {
      return { groupId, userId, added: false, alreadyMember: true };
    }
    throw error;
  }
}

export async function removeGroupMember(
  client: GraphClient,
  groupId: string,
  userId: string,
): Promise<{ groupId: string; userId: string; removed: boolean; alreadyAbsent: boolean }> {
  try {
    await client.delete(`/groups/${encodeURIComponent(groupId)}/members/${encodeURIComponent(userId)}/$ref`);
    return { groupId, userId, removed: true, alreadyAbsent: false };
  } catch (error) {
    if (error instanceof GraphHttpError && error.status === 404) {
      return { groupId, userId, removed: false, alreadyAbsent: true };
    }
    throw error;
  }
}

/** Attributes that {@link updateUser} may PATCH onto an existing user. */
export interface UpdateUserInput {
  accountEnabled?: boolean;
  displayName?: string;
  givenName?: string;
  surname?: string;
  jobTitle?: string;
  department?: string;
  companyName?: string;
  employeeType?: string;
  mobilePhone?: string;
  streetAddress?: string;
  city?: string;
  postalCode?: string;
  state?: string;
  country?: string;
  officeLocation?: string;
  usageLocation?: string;
  otherMails?: string[];
  businessPhones?: string[];
}

const UPDATE_USER_FIELDS: readonly (keyof UpdateUserInput)[] = [
  'accountEnabled',
  'displayName',
  'givenName',
  'surname',
  'jobTitle',
  'department',
  'companyName',
  'employeeType',
  'mobilePhone',
  'streetAddress',
  'city',
  'postalCode',
  'state',
  'country',
  'officeLocation',
  'usageLocation',
  'otherMails',
  'businessPhones',
];

/**
 * PATCH a curated set of attributes onto an existing user. Enables/disables the
 * account (`accountEnabled`) and updates profile fields; only keys present in
 * the input are sent. Graph returns 204 No Content on success.
 */
export async function updateUser(
  client: GraphClient,
  userId: string,
  patch: UpdateUserInput,
): Promise<{ userId: string; updated: string[] }> {
  const body: Record<string, unknown> = {};
  for (const key of UPDATE_USER_FIELDS) {
    if (patch[key] !== undefined) {
      body[key] = patch[key];
    }
  }

  const updated = Object.keys(body);
  if (updated.length === 0) {
    throw new Error('updateUser requires at least one field to change.');
  }

  await client.patch(`/users/${encodeURIComponent(userId)}`, body);
  return { userId, updated };
}

/** Set a user's manager (`manager/$ref`). Pass an object id or userPrincipalName. */
export async function setManager(
  client: GraphClient,
  userId: string,
  managerId: string,
): Promise<{ userId: string; managerId: string }> {
  await client.request({
    method: 'PUT',
    path: `/users/${encodeURIComponent(userId)}/manager/$ref`,
    body: { '@odata.id': client.directoryObjectUrl(managerId) },
  });
  return { userId, managerId };
}

/** Remove a user's manager assignment. Idempotent when no manager is set. */
export async function removeManager(
  client: GraphClient,
  userId: string,
): Promise<{ userId: string; removed: boolean }> {
  try {
    await client.delete(`/users/${encodeURIComponent(userId)}/manager/$ref`);
    return { userId, removed: true };
  } catch (error) {
    if (error instanceof GraphHttpError && error.status === 404) {
      return { userId, removed: false };
    }
    throw error;
  }
}

export interface CreateTapInput {
  /** Minutes the pass is valid. Bounded by the tenant TAP policy (Graph errors if it exceeds the cap). */
  lifetimeInMinutes?: number;
  /** Multi-use when false (default), single-use when true. */
  isUsableOnce?: boolean;
  /** ISO 8601 activation time; the pass is valid from this moment (future-dating supported). */
  startDateTime?: string;
  /** When true (default), regenerate until the passcode is purely alphanumeric. */
  requireAlphanumeric?: boolean;
  /** Max regeneration attempts when requiring alphanumeric (default 8). */
  maxAttempts?: number;
}

interface GraphTapMethod {
  id: string;
  temporaryAccessPass: string;
  isUsableOnce?: boolean;
  lifetimeInMinutes?: number;
  startDateTime?: string;
  methodUsabilityReason?: string;
}

export interface CreateTapResult {
  userId: string;
  id: string;
  temporaryAccessPass: string;
  isUsableOnce: boolean;
  lifetimeInMinutes?: number;
  startDateTime?: string;
  methodUsabilityReason?: string;
  attempts: number;
}

const ALPHANUMERIC = /^[A-Za-z0-9]+$/;

function tapMethodsPath(userId: string): string {
  return `/users/${encodeURIComponent(userId)}/authentication/temporaryAccessPassMethods`;
}

/**
 * Create a Temporary Access Pass for a user. Some tenants issue passcodes
 * containing symbols (e.g. `^@#$`); when {@link CreateTapInput.requireAlphanumeric}
 * is set (the default) any non-alphanumeric pass is deleted and regenerated up
 * to `maxAttempts` times so the pass is easy to relay over SMS/print/email.
 */
export async function createTemporaryAccessPass(
  client: GraphClient,
  userId: string,
  input: CreateTapInput = {},
): Promise<CreateTapResult> {
  const path = tapMethodsPath(userId);
  const requireAlphanumeric = input.requireAlphanumeric !== false;
  const maxAttempts = Math.max(1, input.maxAttempts ?? 8);

  const body: Record<string, unknown> = { isUsableOnce: input.isUsableOnce ?? false };
  if (input.lifetimeInMinutes !== undefined) body.lifetimeInMinutes = input.lifetimeInMinutes;
  if (input.startDateTime) body.startDateTime = input.startDateTime;

  let attempts = 0;
  while (attempts < maxAttempts) {
    attempts += 1;
    const response = await client.post<GraphTapMethod>(path, body);
    const tap = response.data;

    if (!requireAlphanumeric || ALPHANUMERIC.test(tap.temporaryAccessPass)) {
      return {
        userId,
        id: tap.id,
        temporaryAccessPass: tap.temporaryAccessPass,
        isUsableOnce: tap.isUsableOnce ?? (input.isUsableOnce ?? false),
        lifetimeInMinutes: tap.lifetimeInMinutes,
        startDateTime: tap.startDateTime,
        methodUsabilityReason: tap.methodUsabilityReason,
        attempts,
      };
    }

    // Non-alphanumeric passcode: delete it and try again.
    await client.delete(`${path}/${encodeURIComponent(tap.id)}`);
  }

  throw new Error(
    `Could not generate an alphanumeric Temporary Access Pass after ${maxAttempts} attempts. Retry or set requireAlphanumeric=false.`,
  );
}

/**
 * Delete a user's Temporary Access Pass. With `methodId`, deletes that pass;
 * otherwise deletes every TAP currently on the user.
 */
export async function deleteTemporaryAccessPass(
  client: GraphClient,
  userId: string,
  methodId?: string,
): Promise<{ userId: string; deleted: string[] }> {
  const path = tapMethodsPath(userId);

  if (methodId) {
    await client.delete(`${path}/${encodeURIComponent(methodId)}`);
    return { userId, deleted: [methodId] };
  }

  const existing = await client.get<{ value?: GraphTapMethod[] }>(path);
  const ids = (existing.data?.value ?? []).map(method => method.id);
  for (const id of ids) {
    await client.delete(`${path}/${encodeURIComponent(id)}`);
  }
  return { userId, deleted: ids };
}

function isAlreadyMemberError(error: unknown): boolean {
  if (!(error instanceof GraphHttpError)) {
    return false;
  }
  if (error.status !== 400) {
    return false;
  }
  return /already exist|already a member|added object references already exist/i.test(error.message);
}

function defaultMailNickname(userPrincipalName: string): string {
  const local = userPrincipalName.split('@')[0] ?? userPrincipalName;
  const cleaned = local.replace(/[^A-Za-z0-9._-]/g, '');
  return cleaned || 'user';
}

function describeSkus(skus: readonly SubscribedSku[]): string {
  if (!skus.length) {
    return '(none)';
  }
  return skus
    .map(sku => `${sku.skuPartNumber} (${(sku.prepaidUnits?.enabled ?? 0) - (sku.consumedUnits ?? 0)} available)`)
    .join(', ');
}
