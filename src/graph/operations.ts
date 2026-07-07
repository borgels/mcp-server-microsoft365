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
