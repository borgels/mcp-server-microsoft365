import { describe, expect, it } from 'vitest';
import {
  MICROSOFT365_DELEGATED_PERMISSIONS,
  REFUSAL_TEXT,
  createMicrosoft365Gateway,
  microsoft365ToolCatalog,
  permissionCovers,
  type GatewayJsonObject,
} from '../src/gateway.js';

const TOKEN = 'delegated-token-0123456789abcdef';
const DOWNLOAD = 'https://example-my.sharepoint.com/personal/download.aspx?tempauth=abc';

interface Mailbox {
  /** The message the mock answers for /me/messages/{id}. */
  message?: { isDraft?: boolean; parentFolderId?: string; from?: unknown; sender?: unknown };
  /** Messages by id, kept up to date by a move (Graph gives a moved message a new id). */
  messages?: Record<string, Record<string, unknown>>;
  /** What GET /me answers instead of the person's profile. */
  profile?: Record<string, unknown>;
  /** A mail folder lookup that fails with this status. */
  failFolder?: { name: string; status: number };
  /** The event the mock answers for /me/events/{id}. */
  event?: { isOrganizer?: boolean; attendees?: unknown[] };
  /** The drive item the mock answers. */
  item?: Record<string, unknown>;
  /** Where the content address redirects. */
  download?: string;
  /** Folder ids by the name or id asked for; otherwise "folder-" and the name. */
  folders?: Record<string, string>;
}

/** Answers every request of the personal tools as Microsoft Graph documents it. */
function mailbox(state: Mailbox = {}) {
  return (req: Request): Response => {
    const url = new URL(req.url);
    if (url.origin !== 'https://graph.microsoft.com') return new Response('hello, text', { status: 200, headers: { 'content-type': 'text/plain' } });
    const path = url.pathname.replace(/^\/v1\.0/, '');
    if (req.method === 'GET' && path.endsWith('/content')) return new Response(null, { status: 302, headers: { location: state.download ?? DOWNLOAD } });
    if (req.method === 'GET' && /^\/me\/mailFolders\/[^/]+$/.test(path)) {
      const asked = path.split('/').pop()!;
      if (state.failFolder?.name === asked) return Response.json({ error: { code: 'ServiceUnavailable' } }, { status: state.failFolder.status });
      return Response.json({ id: state.folders?.[asked] ?? `folder-${asked}` });
    }
    if (req.method === 'GET' && /^\/me\/messages\/[^/]+$/.test(path)) {
      const asked = path.split('/').pop()!;
      if (state.messages) {
        const found = state.messages[asked];
        return found ? Response.json(found) : Response.json({ error: { code: 'ErrorItemNotFound' } }, { status: 404 });
      }
      return Response.json({ id: 'm-1', isDraft: true, parentFolderId: 'folder-inbox', from: { emailAddress: { address: 'kunde@example.net' } }, ...state.message });
    }
    if (req.method === 'POST' && /^\/me\/messages\/[^/]+\/move$/.test(path) && state.messages) {
      const old = path.split('/')[3]!;
      const moved = { ...state.messages[old]!, id: `${old}-moved`, parentFolderId: 'folder-archive' };
      delete state.messages[old];
      state.messages[moved.id as string] = moved;
      return Response.json(moved, { status: 201 });
    }
    if (req.method === 'GET' && path === '/me') {
      if (state.profile) return Response.json(state.profile);
      return Response.json({ id: 'me-1', mail: 'person@example.com', userPrincipalName: 'person@example.com', proxyAddresses: ['SMTP:person@example.com', 'smtp:alias@example.com'] });
    }
    if (req.method === 'GET' && /^\/me\/events\/[^/]+$/.test(path)) return Response.json({ id: 'e-1', isOrganizer: true, attendees: [], ...state.event });
    if (req.method === 'GET' && /\/attachments\/[^/]+$/.test(path)) {
      return Response.json({ id: 'a-1', name: 'notes.txt', contentType: 'text/plain', size: 5, contentBytes: Buffer.from('notes').toString('base64') });
    }
    if (req.method === 'GET' && /\/items\/[^/]+$/.test(path)) {
      return Response.json({ id: 'i-1', name: 'notes.txt', size: 11, file: { mimeType: 'text/plain' }, webUrl: 'https://example.sharepoint.com/x', ...state.item });
    }
    if (req.method === 'GET') return Response.json({ value: [] });
    if (req.method === 'DELETE') return new Response(null, { status: 204 });
    if (req.method === 'POST' && /\/(send|reply|replyAll|forward|accept|decline|tentativelyAccept|cancel)$/.test(path)) return new Response(null, { status: 202 });
    if (req.method === 'POST' && path === '/me/sendMail') return new Response(null, { status: 202 });
    return Response.json({ id: 'new-1' }, { status: req.method === 'POST' || req.method === 'PUT' ? 201 : 200 });
  };
}

function delegated(state: Mailbox = {}, policy: { enableWrites?: boolean; disabledTools?: string[] } = { enableWrites: true }) {
  const requests: Request[] = [];
  const bodies: string[] = [];
  const handler = mailbox(state);
  const gateway = createMicrosoft365Gateway({
    accessToken: TOKEN,
    policy: { access: 'delegated', ...policy },
    fetchImpl: async (input, init) => {
      const req = new Request(input, init);
      requests.push(req);
      bodies.push(init?.body === undefined || init.body === null ? '' : typeof init.body === 'string' ? init.body : Buffer.from(init.body as Uint8Array).toString('utf8'));
      return handler(req);
    },
  });
  const graph = () => requests.filter(req => new URL(req.url).origin === 'https://graph.microsoft.com');
  const changes = () => graph().filter(req => req.method !== 'GET');
  return { gateway, requests, bodies, graph, changes };
}

const personalCatalog = microsoft365ToolCatalog.filter(entry => entry.access === 'delegated');

const inputs: Record<string, GatewayJsonObject> = {
  get_my_profile: {},
  search_mail: { query: 'tilbud', from: 'kunde@example.com' },
  get_message: { messageId: 'm-1' },
  get_conversation: { conversationId: 'c-1' },
  list_mail_folders: { parentFolderId: 'f-1' },
  list_attachments: { messageId: 'm-1' },
  get_attachment_text: { messageId: 'm-1', attachmentId: 'a-1' },
  get_automatic_replies: {},
  create_draft: { subject: 'Hej', body: 'Tekst', to: ['a@example.com'] },
  update_draft: { messageId: 'm-1', body: 'Ny tekst' },
  create_reply_draft: { messageId: 'm-1', comment: 'Tak', replyAll: true },
  create_forward_draft: { messageId: 'm-1', to: ['b@example.com'] },
  delete_draft: { messageId: 'm-1' },
  move_message: { messageId: 'm-1', destination: 'archive' },
  mark_message: { messageId: 'm-1', isRead: true, flag: 'flagged' },
  send_draft: { messageId: 'm-1' },
  send_mail: { subject: 'Hej', body: 'Tekst', to: ['a@example.com'], cc: ['b@example.com'] },
  reply: { messageId: 'm-1', comment: 'Tak' },
  forward: { messageId: 'm-1', to: ['b@example.com'] },
  set_automatic_replies: { status: 'scheduled', internalMessage: 'På ferie', start: '2026-12-20T00:00:00', end: '2027-01-02T00:00:00', timeZone: 'Europe/Copenhagen' },
  list_calendars: {},
  list_events: { start: '2026-12-01', end: '2026-12-08', calendarId: 'cal-1', timeZone: 'Europe/Copenhagen' },
  get_event: { eventId: 'e-1' },
  get_schedule: { people: ['a@example.com'], start: '2026-12-01T08:00:00', end: '2026-12-01T16:00:00' },
  find_meeting_times: { attendees: ['a@example.com'], durationMinutes: 30 },
  create_event: { subject: 'Møde', start: '2026-12-01T09:00:00', end: '2026-12-01T10:00:00', timeZone: 'Europe/Copenhagen', attendees: ['a@example.com'] },
  update_event: { eventId: 'e-1', start: '2026-12-01T10:00:00', end: '2026-12-01T11:00:00' },
  respond_to_event: { eventId: 'e-1', response: 'tentative' },
  cancel_event: { eventId: 'e-1', comment: 'Aflyst' },
  delete_event: { eventId: 'e-1' },
  search_files: { query: 'budget' },
  list_folder: { driveId: 'd-1', itemId: 'i-1' },
  get_file: { itemId: 'i-1' },
  read_file_text: { driveId: 'd-1', itemId: 'i-1' },
  search_sites: { query: 'salg' },
  list_site_drives: { siteId: 'example.sharepoint.com,1111,2222' },
  upload_file: { name: 'Referat.txt', text: 'Referat', parentId: 'p-1' },
  upload_file_version: { itemId: 'i-1', contentBase64: Buffer.from('v2').toString('base64') },
  create_folder: { name: 'Tilbud', driveId: 'd-1' },
  move_file: { itemId: 'i-1', newName: 'Nyt navn.txt' },
};

const pattern = (path: string) => new RegExp(`^${path.replace(/[.*+?^$()|[\]\\]/g, '\\$&').replace(/\\\{[a-z]+\\\}|\{[a-z]+\}/gi, '[^/]+')}$`);

describe('the personal tools', () => {
  it('have an input here for every tool', () => {
    expect(Object.keys(inputs).sort()).toEqual(personalCatalog.map(entry => entry.name).sort());
  });

  it('each name their delegated permissions, covered by the list an app registration is given; every listed one is used', () => {
    for (const entry of personalCatalog) {
      expect(entry.permissions?.length, entry.name).toBeGreaterThan(0);
      for (const needed of entry.permissions!) {
        expect(MICROSOFT365_DELEGATED_PERMISSIONS.some(granted => permissionCovers(granted, needed)), `${entry.name}: ${needed}`).toBe(true);
      }
    }
    for (const granted of MICROSOFT365_DELEGATED_PERMISSIONS) {
      expect(personalCatalog.some(entry => entry.permissions!.some(needed => permissionCovers(granted, needed))), granted).toBe(true);
    }
    // move_message reads the person's own addresses as well as moving the message.
    expect(personalCatalog.find(entry => entry.name === 'move_message')!.permissions).toEqual(['Mail.ReadWrite', 'User.Read']);
  });

  it('take no user, mailbox or calendar owner, and address only the person, their drives, sites or search', () => {
    for (const entry of personalCatalog) {
      const fields = Object.keys((entry.inputSchema.properties ?? {}) as object);
      for (const field of fields) expect(/^(user|userId|mailbox|owner|upn|userPrincipalName)$/i.test(field), `${entry.name}.${field}`).toBe(false);
      for (const op of entry.upstream) {
        expect(/^\/(me(\/|$)|drives\/|sites(\/|$)|search\/query$)/.test(op.path), `${entry.name}: ${op.path}`).toBe(true);
        expect(op.path.startsWith('/users'), `${entry.name}: ${op.path}`).toBe(false);
      }
    }
  });

  it('have no tool for inbox rules, mailbox settings beyond the automatic reply, sharing, permissions or deleting files', () => {
    const paths = personalCatalog.flatMap(entry => entry.upstream.map(op => `${op.method} ${op.path}`));
    for (const path of paths) {
      expect(/messageRules|forwarding|permissions|invite|createLink|members|\/versions/i.test(path), path).toBe(false);
      if (path.startsWith('DELETE')) expect(['DELETE /me/messages/{message}', 'DELETE /me/events/{event}']).toContain(path);
      if (path.startsWith('PATCH /me/mailboxSettings')) expect(path).toBe('PATCH /me/mailboxSettings');
    }
  });

  for (const entry of personalCatalog) {
    it(`${entry.name}: runs, and every Graph request is in its upstream list`, async () => {
      const { gateway, graph } = delegated();
      const result = await gateway.callTool(entry.name, inputs[entry.name]!);
      expect(result.isError, `${entry.name}: ${JSON.stringify(result.content)}`).toBeFalsy();
      expect(graph().length).toBeGreaterThan(0);
      for (const req of graph()) {
        const path = decodeURIComponent(new URL(req.url).pathname.replace(/^\/v1\.0/, ''));
        const listed = entry.upstream.some(op => op.method === req.method && pattern(op.path).test(path));
        expect(listed, `${entry.name}: ${req.method} ${path}`).toBe(true);
        expect(req.headers.get('authorization')).toBe(`Bearer ${TOKEN}`);
      }
    });
  }
});

describe('host mode and the kind of access', () => {
  it('offers the directory tools for an app-only access and the personal tools for a delegated one, never both', () => {
    const app = createMicrosoft365Gateway({ accessToken: TOKEN, policy: { enableWrites: true } });
    const person = createMicrosoft365Gateway({ accessToken: TOKEN, policy: { access: 'delegated', enableWrites: true } });
    const personal = new Set(personalCatalog.map(entry => entry.name));
    expect(app.tools.some(tool => personal.has(tool.name))).toBe(false);
    expect(person.tools.map(tool => tool.name).sort()).toEqual([...personal].sort());
  });

  it('offers only reads without writes, and refuses a directory tool on a delegated access before any request', async () => {
    const { gateway, requests } = delegated({}, {});
    expect(gateway.tools.every(tool => tool.riskLevel === 'read')).toBe(true);
    expect((await gateway.callTool('send_mail', inputs.send_mail!)).isError).toBe(true);
    expect((await gateway.callTool('list_users', {})).isError).toBe(true);
    expect(requests).toHaveLength(0);
  });

  it('refuses an unknown kind of access', () => {
    expect(() => createMicrosoft365Gateway({ accessToken: TOKEN, policy: { access: 'everyone' as never } })).toThrow(/policy.access/);
  });

  it("refuses another person's mailbox as input before any request", async () => {
    const { gateway, requests } = delegated();
    const result = await gateway.callTool('search_mail', { query: 'x', user: 'boss@example.com' } as never);
    expect(result.isError).toBe(true);
    expect(requests).toHaveLength(0);
  });
});

describe('the fixed refusals', () => {
  const refusal = (code: keyof typeof REFUSAL_TEXT) => ({ isError: true, content: [{ type: 'text', text: REFUSAL_TEXT[code] }] });

  it('changes or deletes only a draft, read just before', async () => {
    for (const tool of ['update_draft', 'delete_draft'] as const) {
      const { gateway, changes } = delegated({ message: { isDraft: false } });
      expect(await gateway.callTool(tool, inputs[tool]!)).toEqual(refusal('not_a_draft'));
      expect(changes()).toHaveLength(0);
    }
  });

  it('moves no sent message to Deleted Items, and nothing into the recoverable items', async () => {
    const sent = delegated({ message: { parentFolderId: 'folder-sentitems' } });
    expect(await sent.gateway.callTool('move_message', { messageId: 'm-1', destination: 'deleteditems' })).toEqual(refusal('sent_message_to_deleted_items'));
    expect(sent.changes()).toHaveLength(0);
    const received = delegated();
    expect((await received.gateway.callTool('move_message', { messageId: 'm-1', destination: 'deleteditems' })).isError).toBeFalsy();
    const recoverable = delegated();
    expect(await recoverable.gateway.callTool('move_message', { messageId: 'm-1', destination: 'RecoverableItemsDeletions' })).toEqual(refusal('recoverable_items'));
    expect(recoverable.requests).toHaveLength(0);
    // Deleted Items given by its id is compared by id.
    const sentById = delegated({ message: { parentFolderId: 'folder-sentitems' }, folders: { 'AAMkDeleted': 'folder-deleteditems' } });
    expect(await sentById.gateway.callTool('move_message', { messageId: 'm-1', destination: 'AAMkDeleted' })).toEqual(refusal('sent_message_to_deleted_items'));
    expect(sentById.changes()).toHaveLength(0);
  });

  it('moves a sent message to Archive, and then still not to Deleted Items under the new id Graph gave it', async () => {
    const sentMessage = {
      id: 'm-sent',
      isDraft: false,
      parentFolderId: 'folder-sentitems',
      from: { emailAddress: { address: 'Person@Example.com' } },
      sender: { emailAddress: { address: 'person@example.com' } },
    };
    const box = delegated({ messages: { 'm-sent': sentMessage } });
    const archived = await box.gateway.callTool('move_message', { messageId: 'm-sent', destination: 'archive' });
    expect(archived.isError).toBeFalsy();
    const newId = (archived.structuredContent as { id?: string }).id!;
    expect(newId).toBe('m-sent-moved');
    const before = box.changes().length;
    expect(await box.gateway.callTool('move_message', { messageId: newId, destination: 'deleteditems' })).toEqual(refusal('sent_message_to_deleted_items'));
    expect(box.changes()).toHaveLength(before);
  });

  it('tells a sent message by an alias of the person, and lets a received one go to Deleted Items', async () => {
    const alias = delegated({ messages: { a: { id: 'a', isDraft: false, parentFolderId: 'folder-archive', from: { emailAddress: { address: 'alias@example.com' } } } } });
    expect(await alias.gateway.callTool('move_message', { messageId: 'a', destination: 'deleteditems' })).toEqual(refusal('sent_message_to_deleted_items'));
    const received = delegated({ messages: { r: { id: 'r', isDraft: false, parentFolderId: 'folder-inbox', from: { emailAddress: { address: 'kunde@example.net' } } } } });
    expect((await received.gateway.callTool('move_message', { messageId: 'r', destination: 'deleteditems' })).isError).toBeFalsy();
  });

  it('moves nothing to Deleted Items when it is not known who sent the message', async () => {
    const cases: Array<[string, Mailbox]> = [
      ['an empty profile', { profile: {}, messages: { x: { id: 'x', isDraft: false, from: { emailAddress: { address: 'kunde@example.net' } } } } }],
      ['no from or sender', { messages: { x: { id: 'x', isDraft: false, parentFolderId: 'folder-inbox' } } }],
      ['no isDraft', { messages: { x: { id: 'x', parentFolderId: 'folder-inbox', from: { emailAddress: { address: 'kunde@example.net' } } } } }],
    ];
    for (const [what, state] of cases) {
      const box = delegated(state);
      expect(await box.gateway.callTool('move_message', { messageId: 'x', destination: 'deleteditems' }), what).toEqual(refusal('sender_unverified'));
      expect(box.changes(), what).toHaveLength(0);
      // Another folder than Deleted Items needs none of it.
      expect((await delegated(state).gateway.callTool('move_message', { messageId: 'x', destination: 'archive' })).isError, what).toBeFalsy();
    }
  });

  it('moves nothing when a lookup the checks need fails', async () => {
    for (const name of ['recoverableitemsdeletions', 'deleteditems', 'sentitems']) {
      const box = delegated({ failFolder: { name, status: 503 } });
      const result = await box.gateway.callTool('move_message', { messageId: 'm-1', destination: 'archive' });
      expect(result.isError, name).toBe(true);
      expect(box.changes(), name).toHaveLength(0);
    }
    // A recoverable items folder Graph says is not there is no folder the destination can be.
    const missing = delegated({ failFolder: { name: 'recoverableitemsdeletions', status: 404 } });
    expect((await missing.gateway.callTool('move_message', { messageId: 'm-1', destination: 'archive' })).isError).toBeFalsy();
  });

  it('deletes only an own appointment without attendees', async () => {
    const meeting = delegated({ event: { attendees: [{ emailAddress: { address: 'a@example.com' } }] } });
    expect(await meeting.gateway.callTool('delete_event', { eventId: 'e-1' })).toEqual(refusal('not_own_appointment'));
    const invitation = delegated({ event: { isOrganizer: false } });
    expect(await invitation.gateway.callTool('delete_event', { eventId: 'e-1' })).toEqual(refusal('not_own_appointment'));
    expect([...meeting.changes(), ...invitation.changes()]).toHaveLength(0);
  });

  it('refuses a search with unreadOnly, a value a tool does not take, and an upload without content', async () => {
    const { gateway, requests } = delegated();
    expect(await gateway.callTool('search_mail', { query: 'x', unreadOnly: true })).toEqual(refusal('query_with_unread'));
    expect(await gateway.callTool('respond_to_event', { eventId: 'e-1', response: 'maybe' })).toEqual(refusal('invalid_value'));
    expect(await gateway.callTool('upload_file', { name: 'a.txt' })).toEqual(refusal('content_missing'));
    expect(await gateway.callTool('upload_file', { name: '../a.txt', text: 'x' })).toEqual(refusal('invalid_value'));
    expect(await gateway.callTool('mark_message', { messageId: 'm-1' })).toEqual(refusal('no_change'));
    expect(requests).toHaveLength(0);
  });
});

describe('what is sent', () => {
  it('sends a new message from the person, saved in Sent Items', async () => {
    const { gateway, changes, bodies, requests } = delegated();
    await gateway.callTool('send_mail', inputs.send_mail!);
    expect(changes().map(req => `${req.method} ${new URL(req.url).pathname}`)).toEqual(['POST /v1.0/me/sendMail']);
    const sent = JSON.parse(bodies[requests.indexOf(changes()[0]!)]!);
    expect(sent).toEqual({
      message: {
        subject: 'Hej',
        body: { contentType: 'text', content: 'Tekst' },
        toRecipients: [{ emailAddress: { address: 'a@example.com' } }],
        ccRecipients: [{ emailAddress: { address: 'b@example.com' } }],
      },
      saveToSentItems: true,
    });
  });

  it('uploads a new file without replacing one of the same name', async () => {
    const { gateway, changes, bodies, requests } = delegated();
    await gateway.callTool('upload_file', inputs.upload_file!);
    const put = changes()[0]!;
    expect(put.method).toBe('PUT');
    expect(decodeURIComponent(new URL(put.url).pathname)).toBe('/v1.0/me/drive/items/p-1:/Referat.txt:/content');
    expect(new URL(put.url).search).toBe('?@microsoft.graph.conflictBehavior=fail');
    expect(put.headers.get('content-type')).toBe('text/plain; charset=utf-8');
    expect(bodies[requests.indexOf(put)]).toBe('Referat');
  });

  it('changes only the automatic reply in the mailbox settings', async () => {
    const { gateway, changes, bodies, requests } = delegated();
    await gateway.callTool('set_automatic_replies', inputs.set_automatic_replies!);
    const body = JSON.parse(bodies[requests.indexOf(changes()[0]!)]!);
    expect(Object.keys(body)).toEqual(['automaticRepliesSetting']);
    expect(body.automaticRepliesSetting.scheduledStartDateTime).toEqual({ dateTime: '2026-12-20T00:00:00', timeZone: 'Europe/Copenhagen' });
  });

  it('reads a file at the download address SharePoint gives, without the token', async () => {
    const { gateway, requests } = delegated();
    const result = await gateway.callTool('read_file_text', inputs.read_file_text!);
    expect((result.structuredContent as { text?: string }).text).toBe('hello, text');
    const download = requests.find(req => new URL(req.url).origin !== 'https://graph.microsoft.com')!;
    expect(download.url).toBe(DOWNLOAD);
    expect(download.headers.get('authorization')).toBeNull();
  });

  it('fetches no download address outside SharePoint, and no file that is not text', async () => {
    const elsewhere = delegated({ download: 'https://attacker.example.com/file' });
    expect((await elsewhere.gateway.callTool('read_file_text', inputs.read_file_text!)).isError).toBe(true);
    expect(elsewhere.requests.every(req => new URL(req.url).origin === 'https://graph.microsoft.com')).toBe(true);
    const binary = delegated({ item: { name: 'tilbud.docx', file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' } } });
    expect((await binary.gateway.callTool('read_file_text', inputs.read_file_text!)).content[0]!.text).toBe(REFUSAL_TEXT.not_text);
  });

  it('asks for mail bodies as text and times in the zone given', async () => {
    const { gateway, requests } = delegated();
    await gateway.callTool('get_message', { messageId: 'm-1' });
    await gateway.callTool('list_events', inputs.list_events!);
    expect(requests[0]!.headers.get('prefer')).toBe('outlook.body-content-type="text"');
    expect(requests[1]!.headers.get('prefer')).toBe('outlook.timezone="Europe/Copenhagen"');
  });
});
