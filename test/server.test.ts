import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'vitest';
import { createServer, writesEnabledFromEnv } from '../src/server.js';

async function listToolNames(
  enableWrites: boolean,
  flags: { allowPasswordReset?: boolean; allowTap?: boolean } = { allowPasswordReset: true, allowTap: true },
): Promise<string[]> {
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const server = createServer({
    enableWrites,
    clientOptions: { accessToken: 'test-token' },
    protection: {
      allowPasswordReset: flags.allowPasswordReset ?? false,
      allowTap: flags.allowTap ?? false,
      protectedUsers: [],
      protectedGroups: [],
    },
  });
  await server.connect(serverTransport);
  const client = new Client({ name: 'test', version: '0.0.0' });
  await client.connect(clientTransport);
  const { tools } = await client.listTools();
  await client.close();
  await server.close();
  return tools.map(t => t.name).sort();
}

const READ_TOOLS = [
  'get_group',
  'get_user',
  'get_user_license_details',
  'list_compliance_policies',
  'list_device_configurations',
  'list_group_members',
  'list_groups',
  'list_managed_devices',
  'list_subscribed_skus',
  'list_user_devices',
  'list_users',
];

describe('createServer write gating', () => {
  it('registers only read tools when writes are disabled', async () => {
    expect(await listToolNames(false)).toEqual(READ_TOOLS);
  });

  it('registers the write tools too when writes are enabled', async () => {
    const names = await listToolNames(true);
    for (const read of READ_TOOLS) expect(names).toContain(read);
    expect(names).toEqual(expect.arrayContaining(['create_user', 'assign_license', 'add_group_member', 'create_temporary_access_pass', 'delete_temporary_access_pass']));
    expect(names.length).toBeGreaterThan(READ_TOOLS.length);
  });

  it('keeps reset_password and the TAP tools off unless their own flags are set', async () => {
    const names = await listToolNames(true, {});
    expect(names).toContain('create_user');
    expect(names).not.toContain('reset_password');
    expect(names).not.toContain('create_temporary_access_pass');
    expect(names).not.toContain('delete_temporary_access_pass');

    const resetOnly = await listToolNames(true, { allowPasswordReset: true });
    expect(resetOnly).toContain('reset_password');
    expect(resetOnly).not.toContain('create_temporary_access_pass');

    // The flags never override MS_ENABLE_WRITES.
    expect(await listToolNames(false)).toEqual(READ_TOOLS);
  });

  it('reads the flag strictly from the environment', () => {
    expect(writesEnabledFromEnv({})).toBe(false);
    expect(writesEnabledFromEnv({ MS_ENABLE_WRITES: '1' })).toBe(false);
    expect(writesEnabledFromEnv({ MS_ENABLE_WRITES: 'true' })).toBe(true);
  });
});
