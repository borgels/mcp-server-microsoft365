import type { GraphClient } from './client.js';
import { pathSegment } from './path.js';

/**
 * Guardrails for the operations that can hand an account to someone else
 * (password reset, Temporary Access Pass) or change group membership.
 *
 * Privileged identities are refused: any user holding a directory role
 * (directly, through a role-assignable group, or scoped to an administrative
 * unit), any member of a role-assignable group, anyone listed in
 * MS_PROTECTED_USERS, and any group that is role-assignable or listed in
 * MS_PROTECTED_GROUPS. Every lookup fails closed: if the directory cannot be
 * read, the operation is refused rather than attempted.
 */
export interface ProtectionPolicy {
  /** Allow reset_password. Env: MS_ENABLE_PASSWORD_RESET=true. */
  allowPasswordReset: boolean;
  /** Allow creating/deleting Temporary Access Passes. Env: MS_ENABLE_TAP=true. */
  allowTap: boolean;
  /** Lower-cased object ids / UPNs / mail addresses that are always refused. */
  protectedUsers: string[];
  /** Lower-cased group object ids / display names / mailNicknames that are always refused. */
  protectedGroups: string[];
}

export function protectionPolicyFromEnv(env: NodeJS.ProcessEnv = process.env): ProtectionPolicy {
  return {
    allowPasswordReset: env.MS_ENABLE_PASSWORD_RESET === 'true',
    allowTap: env.MS_ENABLE_TAP === 'true',
    protectedUsers: parseList(env.MS_PROTECTED_USERS),
    protectedGroups: parseList(env.MS_PROTECTED_GROUPS),
  };
}

export class ProtectedTargetError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ProtectedTargetError';
  }
}

interface DirectoryObject {
  '@odata.type'?: string;
  id?: string;
  displayName?: string | null;
  mailNickname?: string | null;
  isAssignableToRole?: boolean | null;
}

const MAX_PAGES = 50;

/** Throws unless `userRef` is an ordinary, non-privileged user. */
export async function assertUserNotPrivileged(
  client: GraphClient,
  userRef: string,
  policy: ProtectionPolicy,
): Promise<void> {
  let user: { id?: string; userPrincipalName?: string; mail?: string | null } | undefined;
  try {
    user = (
      await client.get<{ id?: string; userPrincipalName?: string; mail?: string | null }>(
        `/users/${pathSegment(userRef, 'user')}`,
        { $select: 'id,userPrincipalName,mail' },
      )
    ).data;
  } catch (error) {
    throw lookupFailed(`user "${userRef}"`, error);
  }
  if (!user?.id) {
    throw new ProtectedTargetError(`Refusing: could not resolve user "${userRef}" to verify it is not privileged.`);
  }

  const identifiers = [userRef, user.id, user.userPrincipalName, user.mail]
    .filter((v): v is string => typeof v === 'string' && v.length > 0)
    .map(v => v.trim().toLowerCase());
  if (identifiers.some(id => policy.protectedUsers.includes(id))) {
    throw new ProtectedTargetError(`Refusing: user "${userRef}" is listed in MS_PROTECTED_USERS.`);
  }

  let memberships: DirectoryObject[];
  try {
    memberships = await listAll<DirectoryObject>(client, `/users/${pathSegment(user.id, 'user')}/transitiveMemberOf`, {
      $select: 'id,displayName,mailNickname,isAssignableToRole',
    });
  } catch (error) {
    throw lookupFailed(`group/role memberships of "${userRef}"`, error);
  }

  for (const entry of memberships) {
    const type = entry['@odata.type'];
    const label = entry.displayName ?? entry.id ?? 'unknown';
    if (type === '#microsoft.graph.directoryRole') {
      throw new ProtectedTargetError(
        `Refusing: user "${userRef}" holds a directory role (${label}). Privileged accounts must be handled manually.`,
      );
    }
    if (entry.isAssignableToRole === true) {
      throw new ProtectedTargetError(`Refusing: user "${userRef}" is a member of a role-assignable group (${label}).`);
    }
    if (groupIsListed(entry, policy)) {
      throw new ProtectedTargetError(
        `Refusing: user "${userRef}" is a member of a group listed in MS_PROTECTED_GROUPS (${label}).`,
      );
    }
  }

  // Role assignments that do not surface in memberOf (e.g. administrative-unit
  // scoped roles). Requires RoleManagement.Read.Directory.
  let assignments: unknown[];
  try {
    assignments = await listAll<unknown>(client, '/roleManagement/directory/roleAssignments', {
      $filter: `principalId eq '${user.id.replace(/'/g, "''")}'`,
      $select: 'id,roleDefinitionId,directoryScopeId',
    });
  } catch (error) {
    throw lookupFailed(`directory role assignments of "${userRef}"`, error);
  }
  if (assignments.length > 0) {
    throw new ProtectedTargetError(
      `Refusing: user "${userRef}" has ${assignments.length} directory role assignment(s). Privileged accounts must be handled manually.`,
    );
  }
}

/** Throws unless `groupId` is an ordinary group (not role-assignable, not protected). */
export async function assertGroupNotProtected(
  client: GraphClient,
  groupId: string,
  policy: ProtectionPolicy,
): Promise<void> {
  if (policy.protectedGroups.includes(groupId.trim().toLowerCase())) {
    throw new ProtectedTargetError(`Refusing: group "${groupId}" is listed in MS_PROTECTED_GROUPS.`);
  }
  let group: DirectoryObject | undefined;
  try {
    group = (
      await client.get<DirectoryObject>(`/groups/${pathSegment(groupId, 'groupId')}`, {
        $select: 'id,displayName,mailNickname,isAssignableToRole',
      })
    ).data;
  } catch (error) {
    throw lookupFailed(`group "${groupId}"`, error);
  }
  if (!group?.id) {
    throw new ProtectedTargetError(`Refusing: could not resolve group "${groupId}" to verify it is not protected.`);
  }
  if (group.isAssignableToRole === true) {
    throw new ProtectedTargetError(
      `Refusing: group "${group.displayName ?? groupId}" is role-assignable; its membership grants directory roles and must be managed manually.`,
    );
  }
  if (groupIsListed(group, policy)) {
    throw new ProtectedTargetError(`Refusing: group "${group.displayName ?? groupId}" is listed in MS_PROTECTED_GROUPS.`);
  }
}

function groupIsListed(group: DirectoryObject, policy: ProtectionPolicy): boolean {
  if (policy.protectedGroups.length === 0) return false;
  return [group.id, group.displayName, group.mailNickname]
    .filter((v): v is string => typeof v === 'string' && v.length > 0)
    .some(v => policy.protectedGroups.includes(v.trim().toLowerCase()));
}

async function listAll<T>(client: GraphClient, path: string, query: Record<string, string>): Promise<T[]> {
  const items: T[] = [];
  let response = await client.get<{ value?: T[] }>(path, query);
  for (let page = 1; ; page += 1) {
    const value = response.data?.value;
    if (!Array.isArray(value)) {
      throw new Error(`Unexpected response shape from ${path}`);
    }
    items.push(...value);
    if (!response.nextLink) return items;
    if (page >= MAX_PAGES) {
      throw new Error(`Too many pages from ${path}`);
    }
    response = await client.get<{ value?: T[] }>(response.nextLink);
  }
}

function lookupFailed(what: string, error: unknown): ProtectedTargetError {
  const detail = error instanceof Error ? error.message : String(error);
  return new ProtectedTargetError(`Refusing: could not verify ${what} is not privileged (${detail}). Failing closed.`);
}

function parseList(value: string | undefined): string[] {
  return (value ?? '')
    .split(',')
    .map(item => item.trim().toLowerCase())
    .filter(Boolean);
}
