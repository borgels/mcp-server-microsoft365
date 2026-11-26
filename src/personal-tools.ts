/**
 * The gateway tools on the signed-in person's own mail, calendar and files: delegated access only.
 *
 * Each tool names the Graph requests it may make and the delegated permissions they need (the least ones Microsoft
 * documents for them). Every request is on `/me`, a drive or site the person reaches, or Microsoft Search: no tool takes
 * a user or a mailbox, so Microsoft Graph decides whose data it is.
 *
 * Not here, on purpose, and not to be added: inbox rules and forwarding settings, mailbox settings other than the
 * automatic reply, permanent deletion of mail and deletion of sent mail, deletion of files and versions, sharing and
 * permissions, and anything that changes members of sites, teams or groups.
 */
import type { GatewayJsonObject, GatewayToolDefinition, GatewayUpstreamOperation } from './gateway.js';

/** A delegated Microsoft Graph permission a personal tool needs. */
export type DelegatedPermission =
  | 'User.Read'
  | 'Mail.Read'
  | 'Mail.ReadWrite'
  | 'Mail.Send'
  | 'MailboxSettings.Read'
  | 'MailboxSettings.ReadWrite'
  | 'Calendars.Read'
  | 'Calendars.ReadWrite'
  | 'Calendars.Read.Shared'
  | 'Files.Read.All'
  | 'Files.ReadWrite.All'
  | 'Sites.Read.All';

/**
 * The delegated permissions an app registration needs for every personal tool, besides `openid`, `profile`, `email`
 * and `offline_access`. A ReadWrite permission covers the Read one.
 */
export const MICROSOFT365_DELEGATED_PERMISSIONS: readonly DelegatedPermission[] = Object.freeze([
  'User.Read',
  'Mail.ReadWrite',
  'Mail.Send',
  'MailboxSettings.ReadWrite',
  'Calendars.ReadWrite',
  'Calendars.Read.Shared',
  'Files.ReadWrite.All',
  'Sites.Read.All',
]);

/** The permission that covers another: a ReadWrite one covers its Read one. */
export function permissionCovers(granted: string, needed: string): boolean {
  return granted === needed || (granted.includes('.ReadWrite') && granted.replace('.ReadWrite', '.Read') === needed);
}

export interface PersonalTool extends GatewayToolDefinition {
  readonly upstream: readonly GatewayUpstreamOperation[];
  /** The delegated permissions the tool's requests need, the main one first. */
  readonly permissions: readonly DelegatedPermission[];
}

const str = (description: string): GatewayJsonObject => ({ type: 'string', description });
const addresses = (description: string): GatewayJsonObject => ({ type: 'array', items: { type: 'string' }, description });
const top = (maximum: number): GatewayJsonObject => ({ type: 'number', minimum: 1, maximum });

const messageId = str('The message id (from search_mail or get_conversation).');
const eventId = str('The event id (from list_events).');
const driveId = str("The drive id (from search_files or list_site_drives). Leave out for the person's own OneDrive.");
const itemId = str('The file or folder id (from search_files or list_folder).');
const timeZone = str('A time zone, e.g. "Europe/Copenhagen". Times are given and answered in it. Default UTC.');
const body = str('The text of the message.');
const bodyType = { type: 'string', enum: ['text', 'html'], description: 'Whether body is plain text (default) or HTML.' } satisfies GatewayJsonObject;
const dateTime = (what: string) => str(`${what}, ISO 8601 (e.g. "2026-12-01T09:00:00").`);

function object(properties: Record<string, GatewayJsonObject>, required: string[] = []): GatewayJsonObject {
  return { type: 'object', ...(required.length ? { required } : {}), properties, additionalProperties: false };
}

const draftFields = {
  subject: str('The subject.'),
  body,
  bodyType,
  to: addresses('Recipients, as e-mail addresses.'),
  cc: addresses('Copy recipients, as e-mail addresses.'),
  bcc: addresses('Blind copy recipients, as e-mail addresses.'),
  importance: { type: 'string', enum: ['low', 'normal', 'high'] },
} satisfies Record<string, GatewayJsonObject>;

const eventFields = {
  subject: str('The title.'),
  start: dateTime('Start'),
  end: dateTime('End'),
  timeZone,
  allDay: { type: 'boolean', description: 'A whole-day event: start and end at midnight.' },
  body: str('The description.'),
  bodyType,
  location: str('Where, as text.'),
  attendees: addresses('Required attendees, as e-mail addresses. They get an invitation at once.'),
  optionalAttendees: addresses('Optional attendees, as e-mail addresses. They get an invitation at once.'),
  isOnlineMeeting: { type: 'boolean', description: 'Add a Teams meeting link.' },
  showAs: { type: 'string', enum: ['free', 'tentative', 'busy', 'oof', 'workingElsewhere'] },
} satisfies Record<string, GatewayJsonObject>;

const content = {
  text: str('The content as text (UTF-8). Give text or contentBase64.'),
  contentBase64: str('The content as base64, for a file that is not text. At most 4 MB.'),
  contentType: str('The content type, e.g. "text/csv". Default text/plain for text, else application/octet-stream.'),
} satisfies Record<string, GatewayJsonObject>;

const GET = (path: string): GatewayUpstreamOperation => ({ method: 'GET', path });
const POST = (path: string): GatewayUpstreamOperation => ({ method: 'POST', path });
const PATCH = (path: string): GatewayUpstreamOperation => ({ method: 'PATCH', path });
const PUT = (path: string): GatewayUpstreamOperation => ({ method: 'PUT', path });
const DELETE = (path: string): GatewayUpstreamOperation => ({ method: 'DELETE', path });

/**
 * Reads that Microsoft Graph takes as a POST: a question with a body, which changes nothing (free/busy, meeting time
 * suggestions, Microsoft Search).
 */
export const READ_BY_POST: readonly string[] = Object.freeze(['get_schedule', 'find_meeting_times', 'search_files']);

export const personalTools: readonly PersonalTool[] = [
  // ---------------------------------------------------------------------------------------------------- profile
  {
    name: 'get_my_profile',
    title: 'Get my Microsoft 365 profile',
    description: "The signed-in person's own name, e-mail address and id in Microsoft 365.",
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({}),
    upstream: [GET('/me')],
    permissions: ['User.Read'],
  },
  // ------------------------------------------------------------------------------------------------- mail reads
  {
    name: 'search_mail',
    title: 'Search my mail',
    description:
      "Find messages in the person's own mailbox, newest first: by words (query), sender, dates or unread. Only the person's own mail and mailboxes they have been given access to.",
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({
      query: str('Words to search for in subject, body and sender (Microsoft Search). Cannot be combined with unreadOnly.'),
      from: str("The sender's e-mail address."),
      folder: str('A folder id or a well-known name such as "inbox", "sentitems", "drafts", "archive".'),
      unreadOnly: { type: 'boolean' },
      receivedAfter: dateTime('Received on or after'),
      receivedBefore: dateTime('Received on or before'),
      top: top(50),
      skip: { type: 'number', minimum: 0, maximum: 1000, description: 'Skip this many (not with query).' },
    }),
    upstream: [GET('/me/messages'), GET('/me/mailFolders/{folder}/messages')],
    permissions: ['Mail.Read'],
  },
  {
    name: 'get_message',
    title: 'Read a message',
    description: 'One message with its full text, recipients and whether it has attachments.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({ messageId, bodyType }, ['messageId']),
    upstream: [GET('/me/messages/{message}')],
    permissions: ['Mail.Read'],
  },
  {
    name: 'get_conversation',
    title: 'Read a mail thread',
    description: 'All messages in one conversation (conversationId from a message), oldest first, with their text.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({ conversationId: str('The conversationId of a message.'), top: top(50) }, ['conversationId']),
    upstream: [GET('/me/messages')],
    permissions: ['Mail.Read'],
  },
  {
    name: 'list_mail_folders',
    title: 'List my mail folders',
    description: 'The mail folders, or the folders inside one, with how many messages are unread.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({ parentFolderId: str('A folder id, to list the folders inside it.') }),
    upstream: [GET('/me/mailFolders'), GET('/me/mailFolders/{folder}/childFolders')],
    permissions: ['Mail.Read'],
  },
  {
    name: 'list_attachments',
    title: "List a message's attachments",
    description: 'The names, types and sizes of the attachments on a message, without their content.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({ messageId }, ['messageId']),
    upstream: [GET('/me/messages/{message}/attachments')],
    permissions: ['Mail.Read'],
  },
  {
    name: 'get_attachment_text',
    title: 'Read an attachment as text',
    description: 'The text of an attachment that is a text file (txt, csv, json, html and similar) of at most 1 MB.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({ messageId, attachmentId: str('The attachment id (from list_attachments).') }, ['messageId', 'attachmentId']),
    upstream: [GET('/me/messages/{message}/attachments/{attachment}')],
    permissions: ['Mail.Read'],
  },
  {
    name: 'get_automatic_replies',
    title: 'See my automatic reply',
    description: 'Whether the automatic reply (out of office) is on, for which period, and its texts.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({}),
    upstream: [GET('/me/mailboxSettings/automaticRepliesSetting')],
    permissions: ['MailboxSettings.Read'],
  },
  // ------------------------------------------------------------------------------------------------ mail drafts
  {
    name: 'create_draft',
    title: 'Write a draft',
    description: 'A new message saved in Drafts. Nothing is sent.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object(draftFields),
    upstream: [POST('/me/messages')],
    permissions: ['Mail.ReadWrite'],
  },
  {
    name: 'update_draft',
    title: 'Change a draft',
    description: "Change a draft's subject, text or recipients. Refused for a message that is not a draft. Nothing is sent.",
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ messageId, ...draftFields }, ['messageId']),
    upstream: [GET('/me/messages/{message}'), PATCH('/me/messages/{message}')],
    permissions: ['Mail.ReadWrite'],
  },
  {
    name: 'create_reply_draft',
    title: 'Write a reply as a draft',
    description: 'A reply to a message (to the sender, or to everyone with replyAll), saved in Drafts. Nothing is sent.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ messageId, comment: str('The text of the reply.'), replyAll: { type: 'boolean' } }, ['messageId']),
    upstream: [POST('/me/messages/{message}/createReply'), POST('/me/messages/{message}/createReplyAll')],
    permissions: ['Mail.ReadWrite'],
  },
  {
    name: 'create_forward_draft',
    title: 'Write a forward as a draft',
    description: 'A forward of a message, saved in Drafts. Nothing is sent.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ messageId, to: addresses('Recipients, as e-mail addresses.'), comment: str('A text above the forwarded message.') }, ['messageId']),
    upstream: [POST('/me/messages/{message}/createForward')],
    permissions: ['Mail.ReadWrite'],
  },
  {
    name: 'delete_draft',
    title: 'Delete a draft',
    description: 'Delete a draft that has not been sent; it goes to Deleted Items. Refused for a message that is not a draft.',
    riskLevel: 'destructive',
    enabledByDefault: false,
    inputSchema: object({ messageId }, ['messageId']),
    upstream: [GET('/me/messages/{message}'), DELETE('/me/messages/{message}')],
    permissions: ['Mail.ReadWrite'],
  },
  // ------------------------------------------------------------------------------------------- mail organising
  {
    name: 'move_message',
    title: 'Move a message',
    description:
      'Move a message to another folder, Deleted Items included (it can be moved back). A message the person has sent is not moved to Deleted Items, wherever it lies.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ messageId, destination: str('A folder id or a well-known name such as "inbox", "archive", "deleteditems".') }, [
      'messageId',
      'destination',
    ]),
    // GET /me reads the person's own addresses (User.Read, which every sign-in has) to tell a sent message.
    upstream: [GET('/me/messages/{message}'), GET('/me'), GET('/me/mailFolders/{folder}'), POST('/me/messages/{message}/move')],
    permissions: ['Mail.ReadWrite', 'User.Read'],
  },
  {
    name: 'mark_message',
    title: 'Mark a message',
    description: 'Mark a message read or unread, flag it or mark the flag complete, or set its categories.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object(
      {
        messageId,
        isRead: { type: 'boolean' },
        flag: { type: 'string', enum: ['notFlagged', 'flagged', 'complete'] },
        categories: addresses('Category names. Replaces the categories the message has.'),
      },
      ['messageId'],
    ),
    upstream: [PATCH('/me/messages/{message}')],
    permissions: ['Mail.ReadWrite'],
  },
  // ----------------------------------------------------------------------------------------------- mail sending
  {
    name: 'send_draft',
    title: 'Send a draft',
    description: 'Send a draft from the person. A sent message cannot be called back.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ messageId }, ['messageId']),
    upstream: [POST('/me/messages/{message}/send')],
    permissions: ['Mail.Send'],
  },
  {
    name: 'send_mail',
    title: 'Send a message',
    description: 'Send a new message from the person at once; it is saved in Sent Items. A sent message cannot be called back.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object(draftFields, ['subject', 'body', 'to']),
    upstream: [POST('/me/sendMail')],
    permissions: ['Mail.Send'],
  },
  {
    name: 'reply',
    title: 'Reply to a message',
    description: 'Reply at once, to the sender or to everyone with replyAll. A sent message cannot be called back.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ messageId, comment: str('The text of the reply.'), replyAll: { type: 'boolean' } }, ['messageId', 'comment']),
    upstream: [POST('/me/messages/{message}/reply'), POST('/me/messages/{message}/replyAll')],
    permissions: ['Mail.Send'],
  },
  {
    name: 'forward',
    title: 'Forward a message',
    description: 'Forward a message at once to the recipients given. A sent message cannot be called back.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ messageId, to: addresses('Recipients, as e-mail addresses.'), comment: str('A text above the forwarded message.') }, [
      'messageId',
      'to',
    ]),
    upstream: [POST('/me/messages/{message}/forward')],
    permissions: ['Mail.Send'],
  },
  {
    name: 'set_automatic_replies',
    title: 'Set my automatic reply',
    description:
      'Turn the automatic reply (out of office) on, off, or on for a period, with its texts. Nothing else in the mailbox settings is changed.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object(
      {
        status: { type: 'string', enum: ['disabled', 'alwaysEnabled', 'scheduled'] },
        internalMessage: str('The reply to colleagues.'),
        externalMessage: str('The reply to people outside the organisation.'),
        externalAudience: { type: 'string', enum: ['none', 'contactsOnly', 'all'] },
        start: dateTime('Start of the period (with scheduled)'),
        end: dateTime('End of the period (with scheduled)'),
        timeZone,
      },
      ['status'],
    ),
    upstream: [PATCH('/me/mailboxSettings')],
    permissions: ['MailboxSettings.ReadWrite'],
  },
  // --------------------------------------------------------------------------------------------- calendar reads
  {
    name: 'list_calendars',
    title: 'List my calendars',
    description: "The person's calendars, also ones shared with them.",
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({}),
    upstream: [GET('/me/calendars')],
    permissions: ['Calendars.Read'],
  },
  {
    name: 'list_events',
    title: 'See my calendar',
    description: 'The events in a period, recurring ones as single occurrences, in the default calendar or one named.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object(
      { start: dateTime('From'), end: dateTime('To'), calendarId: str('A calendar id (from list_calendars).'), timeZone, top: top(100) },
      ['start', 'end'],
    ),
    upstream: [GET('/me/calendarView'), GET('/me/calendars/{calendar}/calendarView')],
    permissions: ['Calendars.Read'],
  },
  {
    name: 'get_event',
    title: 'Read an event',
    description: 'One event with its description, attendees and their answers.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({ eventId, timeZone }, ['eventId']),
    upstream: [GET('/me/events/{event}')],
    permissions: ['Calendars.Read'],
  },
  {
    name: 'get_schedule',
    title: 'See when people are free',
    description: 'Free and busy times for colleagues (by e-mail address) in a period. Changes nothing.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object(
      {
        people: addresses('E-mail addresses of the people.'),
        start: dateTime('From'),
        end: dateTime('To'),
        timeZone,
        intervalMinutes: { type: 'number', minimum: 5, maximum: 1440 },
      },
      ['people', 'start', 'end'],
    ),
    upstream: [POST('/me/calendar/getSchedule')],
    permissions: ['Calendars.Read'],
  },
  {
    name: 'find_meeting_times',
    title: 'Find a time for a meeting',
    description: 'Suggested times when the person and the attendees are free. Changes nothing and sends nothing.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object(
      {
        attendees: addresses('E-mail addresses of the attendees.'),
        durationMinutes: { type: 'number', minimum: 5, maximum: 1440 },
        start: dateTime('Earliest'),
        end: dateTime('Latest'),
        timeZone,
        maxCandidates: { type: 'number', minimum: 1, maximum: 20 },
      },
      ['attendees', 'durationMinutes'],
    ),
    upstream: [POST('/me/findMeetingTimes')],
    permissions: ['Calendars.Read.Shared'],
  },
  // -------------------------------------------------------------------------------------------- calendar changes
  {
    name: 'create_event',
    title: 'Create an event',
    description:
      "A new event in the person's calendar. With attendees, Microsoft 365 sends them an invitation at once from the person.",
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ ...eventFields, calendarId: str('A calendar id (from list_calendars). Default: the main calendar.') }, [
      'subject',
      'start',
      'end',
    ]),
    upstream: [POST('/me/events'), POST('/me/calendars/{calendar}/events')],
    permissions: ['Calendars.ReadWrite'],
  },
  {
    name: 'update_event',
    title: 'Change an event',
    description: 'Move or change an event. With attendees, Microsoft 365 sends them the change at once.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ eventId, ...eventFields }, ['eventId']),
    upstream: [PATCH('/me/events/{event}')],
    permissions: ['Calendars.ReadWrite'],
  },
  {
    name: 'respond_to_event',
    title: 'Answer an invitation',
    description: 'Accept, decline or answer tentative to an invitation. The organiser gets the answer unless sendResponse is false.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object(
      {
        eventId,
        response: { type: 'string', enum: ['accept', 'decline', 'tentative'] },
        comment: str('A note to the organiser.'),
        sendResponse: { type: 'boolean' },
      },
      ['eventId', 'response'],
    ),
    upstream: [
      POST('/me/events/{event}/accept'),
      POST('/me/events/{event}/decline'),
      POST('/me/events/{event}/tentativelyAccept'),
    ],
    permissions: ['Calendars.ReadWrite'],
  },
  {
    name: 'cancel_event',
    title: 'Cancel a meeting',
    description: 'Cancel a meeting the person organises. The attendees get the cancellation at once; it cannot be undone.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ eventId, comment: str('A note to the attendees.') }, ['eventId']),
    upstream: [POST('/me/events/{event}/cancel')],
    permissions: ['Calendars.ReadWrite'],
  },
  {
    name: 'delete_event',
    title: 'Delete an appointment',
    description:
      'Delete an appointment the person organises that has no attendees. Refused for a meeting (cancel it) or an invitation (decline it).',
    riskLevel: 'destructive',
    enabledByDefault: false,
    inputSchema: object({ eventId }, ['eventId']),
    upstream: [GET('/me/events/{event}'), DELETE('/me/events/{event}')],
    permissions: ['Calendars.ReadWrite'],
  },
  // ------------------------------------------------------------------------------------------------ file reads
  {
    name: 'search_files',
    title: 'Search my files',
    description: 'Find files in OneDrive and SharePoint that the person can open, by words in the name or the content.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object(
      { query: str('Words to search for.'), top: top(25), from: { type: 'number', minimum: 0, maximum: 1000, description: 'Skip this many.' } },
      ['query'],
    ),
    upstream: [POST('/search/query')],
    permissions: ['Files.Read.All'],
  },
  {
    name: 'list_folder',
    title: 'List a folder',
    description: "The files and folders in a folder; the person's own OneDrive when nothing is named.",
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({ driveId, itemId: str('A folder id. Leave out for the top of the drive.'), top: top(200) }),
    upstream: [
      GET('/me/drive/root/children'),
      GET('/me/drive/items/{item}/children'),
      GET('/drives/{drive}/root/children'),
      GET('/drives/{drive}/items/{item}/children'),
    ],
    permissions: ['Files.Read.All'],
  },
  {
    name: 'get_file',
    title: 'See a file',
    description: 'Name, size, folder, link and who changed a file or folder last, without its content.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({ driveId, itemId }, ['itemId']),
    upstream: [GET('/me/drive/items/{item}'), GET('/drives/{drive}/items/{item}')],
    permissions: ['Files.Read.All'],
  },
  {
    name: 'read_file_text',
    title: 'Read a text file',
    description: 'The text of a text file (txt, csv, md, json, html and similar) of at most 1 MB.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({ driveId, itemId }, ['itemId']),
    upstream: [
      GET('/me/drive/items/{item}'),
      GET('/me/drive/items/{item}/content'),
      GET('/drives/{drive}/items/{item}'),
      GET('/drives/{drive}/items/{item}/content'),
    ],
    permissions: ['Files.Read.All'],
  },
  {
    name: 'search_sites',
    title: 'Find a SharePoint site',
    description: 'SharePoint sites the person can open, by name.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({ query: str('Words in the site name.') }, ['query']),
    upstream: [GET('/sites')],
    permissions: ['Sites.Read.All'],
  },
  {
    name: 'list_site_drives',
    title: "List a site's document libraries",
    description: 'The document libraries of a SharePoint site, with the drive id to list or upload files.',
    riskLevel: 'read',
    enabledByDefault: true,
    inputSchema: object({ siteId: str('The site id (from search_sites).') }, ['siteId']),
    upstream: [GET('/sites/{site}/drives')],
    permissions: ['Sites.Read.All'],
  },
  // ---------------------------------------------------------------------------------------------- file changes
  {
    name: 'upload_file',
    title: 'Upload a file',
    description:
      "A new file in a folder (the person's own OneDrive when none is named), at most 4 MB. An existing file of the same name is not replaced.",
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ name: str('The file name, e.g. "Referat.txt".'), driveId, parentId: str('A folder id. Leave out for the top of the drive.'), ...content }, [
      'name',
    ]),
    upstream: [
      PUT('/me/drive/root:/{name}:/content'),
      PUT('/me/drive/items/{parent}:/{name}:/content'),
      PUT('/drives/{drive}/root:/{name}:/content'),
      PUT('/drives/{drive}/items/{parent}:/{name}:/content'),
    ],
    permissions: ['Files.ReadWrite.All'],
  },
  {
    name: 'upload_file_version',
    title: 'Upload a new version of a file',
    description: 'New content for an existing file, at most 4 MB. The earlier versions stay in the version history.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ driveId, itemId, ...content }, ['itemId']),
    upstream: [PUT('/me/drive/items/{item}/content'), PUT('/drives/{drive}/items/{item}/content')],
    permissions: ['Files.ReadWrite.All'],
  },
  {
    name: 'create_folder',
    title: 'Create a folder',
    description: "A new folder (in the person's own OneDrive when no drive is named). An existing folder of the same name is not replaced.",
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ name: str('The folder name.'), driveId, parentId: str('A folder id. Leave out for the top of the drive.') }, ['name']),
    upstream: [
      POST('/me/drive/root/children'),
      POST('/me/drive/items/{parent}/children'),
      POST('/drives/{drive}/root/children'),
      POST('/drives/{drive}/items/{parent}/children'),
    ],
    permissions: ['Files.ReadWrite.All'],
  },
  {
    name: 'move_file',
    title: 'Move or rename a file',
    description: 'Rename a file or folder, or move it to another folder in the same drive.',
    riskLevel: 'write',
    enabledByDefault: false,
    inputSchema: object({ driveId, itemId, newName: str('The new name.'), newParentId: str('The id of the folder to move it to.') }, ['itemId']),
    upstream: [PATCH('/me/drive/items/{item}'), PATCH('/drives/{drive}/items/{item}')],
    permissions: ['Files.ReadWrite.All'],
  },
];

export const PERSONAL_TOOL_NAMES: ReadonlySet<string> = new Set(personalTools.map(tool => tool.name));
