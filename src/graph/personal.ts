/**
 * Operations on the signed-in person's own mail, calendar and files (delegated access, `/me`).
 *
 * Every path here starts at `/me`, `/drives/{drive}`, `/sites` or `/search/query`, so Microsoft Graph itself decides
 * whose data it is: only what the person can reach in Microsoft 365. No operation takes a user, a mailbox or a
 * calendar owner as input. A few things are refused before anything is changed (RefusedError, a fixed code):
 * changing or deleting something that is not a draft, deleting an event that has left the person's own calendar,
 * moving a sent message to Deleted Items, and moving anything into the recoverable items folders.
 */
import type { GraphClient, GraphQueryValue } from './client.js';
import { GraphHttpError } from '../errors.js';
import { pathSegment } from './path.js';

/** Why a change was refused before anything was sent. The code is the message, and both are fixed. */
export type RefusalCode =
  | 'not_a_draft'
  | 'not_own_appointment'
  | 'sent_message_to_deleted_items'
  | 'recoverable_items'
  | 'query_with_unread'
  | 'no_change'
  | 'content_missing'
  | 'content_too_large'
  | 'invalid_value'
  | 'not_text';

export class RefusedError extends Error {
  readonly code: RefusalCode;
  constructor(code: RefusalCode) {
    super(REFUSAL_TEXT[code]);
    this.name = 'RefusedError';
    this.code = code;
  }
}

/** The fixed text of each refusal: nothing of the input is repeated. */
export const REFUSAL_TEXT: Readonly<Record<RefusalCode, string>> = {
  not_a_draft: 'Refused: the message is not a draft.',
  not_own_appointment: 'Refused: only an appointment without attendees that the person organises can be deleted.',
  sent_message_to_deleted_items: 'Refused: a sent message is not moved to Deleted Items.',
  recoverable_items: 'Refused: nothing is moved into the recoverable items folders.',
  query_with_unread: 'Refused: unreadOnly cannot be combined with query.',
  no_change: 'Refused: give at least one field to change.',
  content_missing: 'Refused: give the content as text or as contentBase64, not both and not neither.',
  content_too_large: 'Refused: the content is larger than 4 MB.',
  invalid_value: 'Refused: a value is not one the tool takes.',
  not_text: 'Refused: the file is not a text file of at most 1 MB.',
};

/** Largest upload sent in one request (Graph's simple upload takes more; we keep it small). */
export const MAX_UPLOAD_BYTES = 4 * 1024 * 1024;
/** Largest text read from a file or an attachment. */
export const MAX_TEXT_BYTES = 1024 * 1024;

const MESSAGE_LIST_FIELDS = [
  'id',
  'subject',
  'from',
  'toRecipients',
  'ccRecipients',
  'receivedDateTime',
  'sentDateTime',
  'isRead',
  'isDraft',
  'hasAttachments',
  'importance',
  'flag',
  'bodyPreview',
  'conversationId',
  'parentFolderId',
  'webLink',
].join(',');

const MESSAGE_FIELDS = `${MESSAGE_LIST_FIELDS},bccRecipients,replyTo,body,categories,internetMessageId`;

const EVENT_FIELDS = [
  'id',
  'subject',
  'start',
  'end',
  'isAllDay',
  'location',
  'organizer',
  'attendees',
  'isOrganizer',
  'isCancelled',
  'isOnlineMeeting',
  'onlineMeeting',
  'responseStatus',
  'showAs',
  'sensitivity',
  'bodyPreview',
  'webLink',
  'seriesMasterId',
  'type',
].join(',');

const DRIVE_ITEM_FIELDS = 'id,name,size,file,folder,webUrl,parentReference,createdDateTime,lastModifiedDateTime,lastModifiedBy';

/** File types read as text: by extension, lower case. */
const TEXT_EXTENSIONS = new Set(['txt', 'csv', 'tsv', 'md', 'json', 'xml', 'html', 'htm', 'ics', 'vtt', 'log', 'yaml', 'yml']);
/** Content types read as text. */
const TEXT_TYPE = /^(text\/[a-z0-9.+-]+|application\/(json|xml|csv|x-yaml|yaml))(\s*;.*)?$/i;

/** One Graph id in a path segment. */
function id(value: string, name: string): string {
  return pathSegment(value, name);
}

/** Recipients as Graph takes them: `{ emailAddress: { address } }`. */
function recipients(addresses: readonly string[] | undefined): Array<{ emailAddress: { address: string } }> | undefined {
  if (addresses === undefined) return undefined;
  return addresses.map(address => ({ emailAddress: { address } }));
}

function bodyOf(content: string | undefined, bodyType: string | undefined): { contentType: string; content: string } | undefined {
  if (content === undefined) return undefined;
  return { contentType: contentTypeOf(bodyType), content };
}

function contentTypeOf(bodyType: string | undefined): 'text' | 'html' {
  if (bodyType === undefined || bodyType === 'text') return 'text';
  if (bodyType === 'html') return 'html';
  throw new RefusedError('invalid_value');
}

function oneOf<T extends string>(value: string | undefined, allowed: readonly T[]): T | undefined {
  if (value === undefined) return undefined;
  if ((allowed as readonly string[]).includes(value)) return value as T;
  throw new RefusedError('invalid_value');
}

function defined<T extends Record<string, unknown>>(value: T): Partial<T> {
  return Object.fromEntries(Object.entries(value).filter(([, entry]) => entry !== undefined)) as Partial<T>;
}

/** The Prefer header for a time zone and the body as text. */
function prefer(options: { timeZone?: string; textBody?: boolean }): Record<string, string> | undefined {
  const parts: string[] = [];
  if (options.timeZone) parts.push(`outlook.timezone="${safeHeaderValue(options.timeZone)}"`);
  if (options.textBody) parts.push('outlook.body-content-type="text"');
  return parts.length ? { Prefer: parts.join(', ') } : undefined;
}

/** A value inside a quoted header parameter: printable, no quote or backslash. */
function safeHeaderValue(value: string): string {
  if (!/^[\x20-\x7e]{1,100}$/.test(value) || /["\\]/.test(value)) throw new RefusedError('invalid_value');
  return value;
}

/** An OData string literal: single quotes doubled. */
function odataString(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

/** An ISO 8601 date or date and time, as a person or an AI gives it; refused otherwise. */
function isoDate(value: string): string {
  if (!/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d{1,7})?)?(Z|[+-]\d{2}:\d{2})?)?$/.test(value)) throw new RefusedError('invalid_value');
  return value;
}

/** A time for Graph's dateTimeTimeZone: the local time and its zone (UTC when none is given). */
function dateTimeTimeZone(value: string, timeZone: string | undefined): { dateTime: string; timeZone: string } {
  return { dateTime: isoDate(value), timeZone: timeZone ? safeHeaderValue(timeZone) : 'UTC' };
}

// ---------------------------------------------------------------------------------------------------------------- mail

export interface SearchMailInput {
  query?: string;
  from?: string;
  folder?: string;
  unreadOnly?: boolean;
  receivedAfter?: string;
  receivedBefore?: string;
  top?: number;
  skip?: number;
}

/**
 * Messages in the person's mailbox, newest first: by a search (KQL in `$search`, with sender and dates added to it)
 * or by a filter on sender, dates and unread. Graph does not take `$search` with `$filter`, so `unreadOnly` with a
 * query is refused.
 */
export async function searchMail(client: GraphClient, input: SearchMailInput) {
  const path = input.folder ? `/me/mailFolders/${id(input.folder, 'folder')}/messages` : '/me/messages';
  const query: Record<string, GraphQueryValue> = { $select: MESSAGE_LIST_FIELDS, $top: input.top ?? 25 };
  if (input.query !== undefined) {
    if (input.unreadOnly) throw new RefusedError('query_with_unread');
    const parts = [input.query.replace(/"/g, ' ')];
    if (input.from) parts.push(`from:${input.from.replace(/[\s"]/g, '')}`);
    if (input.receivedAfter) parts.push(`received>=${isoDate(input.receivedAfter).slice(0, 10)}`);
    if (input.receivedBefore) parts.push(`received<=${isoDate(input.receivedBefore).slice(0, 10)}`);
    query.$search = `"${parts.join(' ')}"`;
  } else {
    // A property in $orderby must also be first in $filter, or Graph refuses the filter as inefficient.
    const filter = [`receivedDateTime ge ${isoDate(input.receivedAfter ?? '1900-01-01T00:00:00Z')}`];
    if (input.receivedBefore) filter.push(`receivedDateTime le ${isoDate(input.receivedBefore)}`);
    if (input.from) filter.push(`from/emailAddress/address eq ${odataString(input.from)}`);
    if (input.unreadOnly) filter.push('isRead eq false');
    query.$filter = filter.join(' and ');
    query.$orderby = 'receivedDateTime desc';
    if (input.skip) query.$skip = input.skip;
  }
  return (await client.get(path, query)).data;
}

/** One message, the body as text unless html is asked for. */
export async function getMessage(client: GraphClient, messageId: string, bodyType: string | undefined) {
  const text = contentTypeOf(bodyType) === 'text';
  return (await client.request({
    method: 'GET',
    path: `/me/messages/${id(messageId, 'messageId')}`,
    query: { $select: MESSAGE_FIELDS },
    headers: prefer({ textBody: text }),
  })).data;
}

/** The messages of one conversation, oldest first, bodies as text. */
export async function getConversation(client: GraphClient, conversationId: string, top: number | undefined) {
  const response = await client.request<{ value?: Array<{ receivedDateTime?: string }> }>({
    method: 'GET',
    path: '/me/messages',
    query: { $filter: `conversationId eq ${odataString(conversationId)}`, $select: `${MESSAGE_LIST_FIELDS},body`, $top: top ?? 25 },
    headers: prefer({ textBody: true }),
  });
  const value = Array.isArray(response.data?.value) ? [...response.data.value] : [];
  value.sort((a, b) => String(a.receivedDateTime ?? '').localeCompare(String(b.receivedDateTime ?? '')));
  return { value };
}

/** The person's mail folders, or the folders inside one. */
export async function listMailFolders(client: GraphClient, parentFolderId: string | undefined) {
  const path = parentFolderId ? `/me/mailFolders/${id(parentFolderId, 'parentFolderId')}/childFolders` : '/me/mailFolders';
  return (await client.get(path, { $top: 100, $select: 'id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount' })).data;
}

/** A message's attachments, without their content. */
export async function listAttachments(client: GraphClient, messageId: string) {
  return (await client.get(`/me/messages/${id(messageId, 'messageId')}/attachments`, { $select: 'id,name,contentType,size,isInline' })).data;
}

/** An attachment's text, for a text file of at most 1 MB; refused for anything else. */
export async function getAttachmentText(client: GraphClient, messageId: string, attachmentId: string) {
  const attachment = (await client.get<Record<string, unknown>>(
    `/me/messages/${id(messageId, 'messageId')}/attachments/${id(attachmentId, 'attachmentId')}`,
  )).data;
  const name = typeof attachment?.name === 'string' ? attachment.name : '';
  const contentType = typeof attachment?.contentType === 'string' ? attachment.contentType : '';
  const size = typeof attachment?.size === 'number' ? attachment.size : Number.POSITIVE_INFINITY;
  const bytes = typeof attachment?.contentBytes === 'string' ? attachment.contentBytes : undefined;
  if (bytes === undefined || size > MAX_TEXT_BYTES || !isText(name, contentType)) throw new RefusedError('not_text');
  return { id: attachment.id ?? attachmentId, name, contentType, size, text: Buffer.from(bytes, 'base64').toString('utf8') };
}

function isText(name: string, contentType: string): boolean {
  const extension = name.includes('.') ? name.split('.').pop()!.toLowerCase() : '';
  return TEXT_TYPE.test(contentType) || TEXT_EXTENSIONS.has(extension);
}

export interface DraftInput {
  subject?: string;
  body?: string;
  bodyType?: string;
  to?: string[];
  cc?: string[];
  bcc?: string[];
  importance?: string;
}

function messageFields(input: DraftInput) {
  return defined({
    subject: input.subject,
    body: bodyOf(input.body, input.bodyType),
    toRecipients: recipients(input.to),
    ccRecipients: recipients(input.cc),
    bccRecipients: recipients(input.bcc),
    importance: oneOf(input.importance, ['low', 'normal', 'high'] as const),
  });
}

/** A new draft in Drafts. Nothing is sent. */
export async function createDraft(client: GraphClient, input: DraftInput) {
  return (await client.post('/me/messages', messageFields(input))).data;
}

/** Refuses unless the message is a draft (read just before the change). */
async function assertDraft(client: GraphClient, messageId: string): Promise<void> {
  const message = (await client.get<{ isDraft?: unknown }>(`/me/messages/${id(messageId, 'messageId')}`, { $select: 'id,isDraft' })).data;
  if (message?.isDraft !== true) throw new RefusedError('not_a_draft');
}

/** Changes a draft's subject, body or recipients. Refused for a message that is not a draft. */
export async function updateDraft(client: GraphClient, messageId: string, input: DraftInput) {
  const fields = messageFields(input);
  if (Object.keys(fields).length === 0) throw new RefusedError('no_change');
  await assertDraft(client, messageId);
  return (await client.patch(`/me/messages/${id(messageId, 'messageId')}`, fields)).data;
}

/** Deletes a draft that has not been sent (it goes to Deleted Items). Refused for a message that is not a draft. */
export async function deleteDraft(client: GraphClient, messageId: string) {
  await assertDraft(client, messageId);
  await client.delete(`/me/messages/${id(messageId, 'messageId')}`);
  return { deleted: true, id: messageId };
}

/** A reply as a draft (to the sender, or to everyone). Nothing is sent. */
export async function createReplyDraft(client: GraphClient, messageId: string, comment: string | undefined, replyAll: boolean | undefined) {
  const action = replyAll ? 'createReplyAll' : 'createReply';
  return (await client.post(`/me/messages/${id(messageId, 'messageId')}/${action}`, defined({ comment }))).data;
}

/** A forward as a draft. Nothing is sent. */
export async function createForwardDraft(client: GraphClient, messageId: string, to: string[] | undefined, comment: string | undefined) {
  return (await client.post(`/me/messages/${id(messageId, 'messageId')}/createForward`, defined({ toRecipients: recipients(to), comment }))).data;
}

const RECOVERABLE = /^recoverableitems/i;

/** An e-mail address as given in a message or a profile, lower case; empty when there is none. */
function addressOf(value: unknown): string {
  const address = (value as { emailAddress?: { address?: unknown } } | null | undefined)?.emailAddress?.address;
  return typeof address === 'string' ? address.trim().toLowerCase() : '';
}

/** The person's own addresses: the primary one, the sign-in name and every SMTP alias. */
function ownAddresses(me: { mail?: unknown; userPrincipalName?: unknown; proxyAddresses?: unknown } | null | undefined): Set<string> {
  const own = new Set<string>();
  for (const value of [me?.mail, me?.userPrincipalName]) if (typeof value === 'string' && value.trim()) own.add(value.trim().toLowerCase());
  if (Array.isArray(me?.proxyAddresses)) {
    for (const proxy of me.proxyAddresses) if (typeof proxy === 'string' && /^smtp:/i.test(proxy)) own.add(proxy.slice(5).trim().toLowerCase());
  }
  return own;
}

/**
 * Moves a message to another folder. Deleted Items is a folder like any other (it can be undone), except for a message
 * the person has sent: one that is not a draft and whose sender is the person, wherever it lies now (a sent message
 * moved to Archive first is still sent). Nothing goes into the recoverable items folders. The folders are compared by
 * id, read just before. When a lookup the checks need fails, nothing is moved; only a recoverable items folder that
 * Graph says does not exist (404) is taken as one the destination cannot be.
 */
export async function moveMessage(client: GraphClient, messageId: string, destination: string) {
  if (RECOVERABLE.test(destination.trim())) throw new RefusedError('recoverable_items');
  const messagePath = `/me/messages/${id(messageId, 'messageId')}`;
  const folderId = async (folder: string) =>
    (await client.get<{ id?: unknown }>(`/me/mailFolders/${id(folder, 'folder')}`, { $select: 'id' })).data?.id;
  const [message, me, target, deletedItems, sentItems, recoverable] = await Promise.all([
    client
      .get<{ parentFolderId?: unknown; isDraft?: unknown; from?: unknown; sender?: unknown }>(messagePath, { $select: 'id,parentFolderId,isDraft,from,sender' })
      .then(r => r.data),
    client.get<{ mail?: unknown; userPrincipalName?: unknown; proxyAddresses?: unknown }>('/me', { $select: 'mail,userPrincipalName,proxyAddresses' }).then(r => r.data),
    folderId(destination),
    folderId('deleteditems'),
    folderId('sentitems'),
    folderId('recoverableitemsdeletions').catch((error: unknown) => {
      if (error instanceof GraphHttpError && error.status === 404) return null;
      throw error;
    }),
  ]);
  if (typeof target !== 'string' || typeof deletedItems !== 'string' || typeof sentItems !== 'string') throw new RefusedError('invalid_value');
  if (recoverable !== null && typeof recoverable !== 'string') throw new RefusedError('invalid_value');
  if (recoverable !== null && target === recoverable) throw new RefusedError('recoverable_items');
  if (target === deletedItems) {
    const own = ownAddresses(me);
    // A message is the person's own sent one unless Graph says it is a draft: a missing isDraft counts as sent.
    const sentByPerson = message?.isDraft !== true && (own.has(addressOf(message?.from)) || own.has(addressOf(message?.sender)));
    if (sentByPerson || message?.parentFolderId === sentItems) throw new RefusedError('sent_message_to_deleted_items');
  }
  return (await client.post(`${messagePath}/move`, { destinationId: target })).data;
}

/** Marks a message read or unread, flags it, or sets its categories. */
export async function markMessage(
  client: GraphClient,
  messageId: string,
  input: { isRead?: boolean; flag?: string; categories?: string[] },
) {
  const flag = oneOf(input.flag, ['notFlagged', 'flagged', 'complete'] as const);
  const fields = defined({ isRead: input.isRead, flag: flag ? { flagStatus: flag } : undefined, categories: input.categories });
  if (Object.keys(fields).length === 0) throw new RefusedError('no_change');
  return (await client.patch(`/me/messages/${id(messageId, 'messageId')}`, fields)).data;
}

/** Sends a draft. A sent message cannot be called back. */
export async function sendDraft(client: GraphClient, messageId: string) {
  await client.post(`/me/messages/${id(messageId, 'messageId')}/send`);
  return { sent: true, id: messageId };
}

/** Sends a new message from the person, saved in Sent Items. A sent message cannot be called back. */
export async function sendMail(client: GraphClient, input: DraftInput) {
  await client.post('/me/sendMail', { message: messageFields(input), saveToSentItems: true });
  return { sent: true };
}

/** Replies (to the sender, or to everyone) at once. A sent message cannot be called back. */
export async function reply(client: GraphClient, messageId: string, comment: string, replyAll: boolean | undefined) {
  const action = replyAll ? 'replyAll' : 'reply';
  await client.post(`/me/messages/${id(messageId, 'messageId')}/${action}`, { comment });
  return { sent: true, id: messageId };
}

/** Forwards a message at once. A sent message cannot be called back. */
export async function forward(client: GraphClient, messageId: string, to: string[], comment: string | undefined) {
  await client.post(`/me/messages/${id(messageId, 'messageId')}/forward`, defined({ toRecipients: recipients(to), comment }));
  return { sent: true, id: messageId };
}

/** The automatic reply setting (out of office). */
export async function getAutomaticReplies(client: GraphClient) {
  return (await client.get('/me/mailboxSettings/automaticRepliesSetting')).data;
}

export interface AutomaticRepliesInput {
  status: string;
  internalMessage?: string;
  externalMessage?: string;
  externalAudience?: string;
  start?: string;
  end?: string;
  timeZone?: string;
}

/** Turns the automatic reply on, off or on for a period. Nothing else in the mailbox settings is sent. */
export async function setAutomaticReplies(client: GraphClient, input: AutomaticRepliesInput) {
  const status = oneOf(input.status, ['disabled', 'alwaysEnabled', 'scheduled'] as const);
  const setting = defined({
    status,
    internalReplyMessage: input.internalMessage,
    externalReplyMessage: input.externalMessage,
    externalAudience: oneOf(input.externalAudience, ['none', 'contactsOnly', 'all'] as const),
    scheduledStartDateTime: input.start ? dateTimeTimeZone(input.start, input.timeZone) : undefined,
    scheduledEndDateTime: input.end ? dateTimeTimeZone(input.end, input.timeZone) : undefined,
  });
  await client.patch('/me/mailboxSettings', { automaticRepliesSetting: setting });
  return { automaticRepliesSetting: setting };
}

// ------------------------------------------------------------------------------------------------------------ calendar

/** The person's calendars. */
export async function listCalendars(client: GraphClient) {
  return (await client.get('/me/calendars', { $top: 100, $select: 'id,name,color,isDefaultCalendar,canEdit,owner' })).data;
}

/** Events in a period (recurring ones as single occurrences), in one calendar or the default one. */
export async function listEvents(
  client: GraphClient,
  input: { start: string; end: string; calendarId?: string; top?: number; timeZone?: string },
) {
  const path = input.calendarId ? `/me/calendars/${id(input.calendarId, 'calendarId')}/calendarView` : '/me/calendarView';
  return (await client.request({
    method: 'GET',
    path,
    query: { startDateTime: isoDate(input.start), endDateTime: isoDate(input.end), $select: EVENT_FIELDS, $orderby: 'start/dateTime', $top: input.top ?? 50 },
    headers: prefer({ timeZone: input.timeZone }),
  })).data;
}

/** One event, the body as text. */
export async function getEvent(client: GraphClient, eventId: string, timeZone: string | undefined) {
  return (await client.request({
    method: 'GET',
    path: `/me/events/${id(eventId, 'eventId')}`,
    query: { $select: `${EVENT_FIELDS},body` },
    headers: prefer({ timeZone, textBody: true }),
  })).data;
}

/** Free and busy times for people (the person's colleagues or the person), in slots. */
export async function getSchedule(
  client: GraphClient,
  input: { people: string[]; start: string; end: string; timeZone?: string; intervalMinutes?: number },
) {
  const timeZone = input.timeZone;
  return (await client.request({
    method: 'POST',
    path: '/me/calendar/getSchedule',
    body: defined({
      schedules: input.people,
      startTime: dateTimeTimeZone(input.start, timeZone),
      endTime: dateTimeTimeZone(input.end, timeZone),
      availabilityViewInterval: input.intervalMinutes,
    }),
    headers: prefer({ timeZone }),
  })).data;
}

/** Suggested times when the attendees are free. */
export async function findMeetingTimes(
  client: GraphClient,
  input: { attendees: string[]; durationMinutes: number; start?: string; end?: string; timeZone?: string; maxCandidates?: number },
) {
  const timeZone = input.timeZone;
  const timeConstraint =
    input.start && input.end
      ? { timeslots: [{ start: dateTimeTimeZone(input.start, timeZone), end: dateTimeTimeZone(input.end, timeZone) }] }
      : undefined;
  return (await client.request({
    method: 'POST',
    path: '/me/findMeetingTimes',
    body: defined({
      attendees: input.attendees.map(address => ({ type: 'required', emailAddress: { address } })),
      meetingDuration: `PT${Math.round(input.durationMinutes)}M`,
      timeConstraint,
      maxCandidates: input.maxCandidates,
      returnSuggestionReasons: true,
    }),
    headers: prefer({ timeZone }),
  })).data;
}

export interface EventInput {
  subject?: string;
  start?: string;
  end?: string;
  timeZone?: string;
  allDay?: boolean;
  body?: string;
  bodyType?: string;
  location?: string;
  attendees?: string[];
  optionalAttendees?: string[];
  isOnlineMeeting?: boolean;
  showAs?: string;
}

function eventFields(input: EventInput) {
  const attendees =
    input.attendees === undefined && input.optionalAttendees === undefined
      ? undefined
      : [
          ...(input.attendees ?? []).map(address => ({ type: 'required', emailAddress: { address } })),
          ...(input.optionalAttendees ?? []).map(address => ({ type: 'optional', emailAddress: { address } })),
        ];
  return defined({
    subject: input.subject,
    start: input.start ? dateTimeTimeZone(input.start, input.timeZone) : undefined,
    end: input.end ? dateTimeTimeZone(input.end, input.timeZone) : undefined,
    isAllDay: input.allDay,
    body: bodyOf(input.body, input.bodyType),
    location: input.location === undefined ? undefined : { displayName: input.location },
    attendees,
    isOnlineMeeting: input.isOnlineMeeting,
    onlineMeetingProvider: input.isOnlineMeeting ? 'teamsForBusiness' : undefined,
    showAs: oneOf(input.showAs, ['free', 'tentative', 'busy', 'oof', 'workingElsewhere'] as const),
  });
}

/** A new event in the person's calendar. With attendees, Microsoft 365 sends them an invitation at once. */
export async function createEvent(client: GraphClient, input: EventInput & { calendarId?: string }) {
  if (!input.subject || !input.start || !input.end) throw new RefusedError('invalid_value');
  const path = input.calendarId ? `/me/calendars/${id(input.calendarId, 'calendarId')}/events` : '/me/events';
  return (await client.post(path, eventFields(input))).data;
}

/** Changes an event. With attendees, Microsoft 365 sends them the change at once. */
export async function updateEvent(client: GraphClient, eventId: string, input: EventInput) {
  const fields = eventFields(input);
  if (Object.keys(fields).length === 0) throw new RefusedError('no_change');
  return (await client.patch(`/me/events/${id(eventId, 'eventId')}`, fields)).data;
}

/** Answers an invitation: accept, decline or tentative. */
export async function respondToEvent(
  client: GraphClient,
  eventId: string,
  input: { response: string; comment?: string; sendResponse?: boolean },
) {
  const response = oneOf(input.response, ['accept', 'decline', 'tentative'] as const)!;
  const action = response === 'tentative' ? 'tentativelyAccept' : response;
  await client.post(`/me/events/${id(eventId, 'eventId')}/${action}`, defined({ comment: input.comment, sendResponse: input.sendResponse ?? true }));
  return { responded: response, id: eventId };
}

/** Cancels a meeting the person organises; Microsoft 365 sends the cancellation to the attendees at once. */
export async function cancelEvent(client: GraphClient, eventId: string, comment: string | undefined) {
  await client.post(`/me/events/${id(eventId, 'eventId')}/cancel`, defined({ comment }));
  return { cancelled: true, id: eventId };
}

/**
 * Deletes an appointment that never left the person's calendar: the person organises it and it has no attendees
 * (read just before). A meeting is cancelled instead (cancel_event), and an invitation is declined.
 */
export async function deleteEvent(client: GraphClient, eventId: string) {
  const path = `/me/events/${id(eventId, 'eventId')}`;
  const event = (await client.get<{ isOrganizer?: unknown; attendees?: unknown }>(path, { $select: 'id,isOrganizer,attendees' })).data;
  const attendees = Array.isArray(event?.attendees) ? event.attendees.length : -1;
  if (event?.isOrganizer !== true || attendees !== 0) throw new RefusedError('not_own_appointment');
  await client.delete(path);
  return { deleted: true, id: eventId };
}

// --------------------------------------------------------------------------------------------------------------- files

/** Files in OneDrive and SharePoint that the person can open, by Microsoft Search. */
export async function searchFiles(client: GraphClient, query: string, top: number | undefined, from: number | undefined) {
  return (await client.post('/search/query', {
    requests: [
      {
        entityTypes: ['driveItem'],
        query: { queryString: query },
        from: from ?? 0,
        size: top ?? 25,
        fields: ['id', 'name', 'webUrl', 'size', 'lastModifiedDateTime', 'parentReference', 'file', 'folder'],
      },
    ],
  })).data;
}

function driveBase(driveId: string | undefined): string {
  return driveId ? `/drives/${id(driveId, 'driveId')}` : '/me/drive';
}

/** The items in a folder: the person's own OneDrive root when nothing is named. */
export async function listFolder(client: GraphClient, driveId: string | undefined, itemId: string | undefined, top: number | undefined) {
  const folder = itemId ? `items/${id(itemId, 'itemId')}` : 'root';
  return (await client.get(`${driveBase(driveId)}/${folder}/children`, { $select: DRIVE_ITEM_FIELDS, $top: top ?? 100 })).data;
}

/** One file or folder, without its content. */
export async function getFile(client: GraphClient, driveId: string | undefined, itemId: string) {
  return (await client.get(`${driveBase(driveId)}/items/${id(itemId, 'itemId')}`, { $select: DRIVE_ITEM_FIELDS })).data;
}

/**
 * The text of a text file of at most 1 MB. Graph answers the content address with a redirect to a short-lived,
 * pre-authenticated download address at SharePoint; that address is fetched without the access token.
 */
export async function readFileText(client: GraphClient, driveId: string | undefined, itemId: string) {
  const itemPath = `${driveBase(driveId)}/items/${id(itemId, 'itemId')}`;
  const item = (await client.get<Record<string, unknown>>(itemPath, { $select: 'id,name,size,file,webUrl' })).data;
  const name = typeof item?.name === 'string' ? item.name : '';
  const file = item?.file as { mimeType?: unknown } | undefined;
  const mimeType = typeof file?.mimeType === 'string' ? file.mimeType : '';
  const size = typeof item?.size === 'number' ? item.size : Number.POSITIVE_INFINITY;
  if (!file || size > MAX_TEXT_BYTES || !isText(name, mimeType)) throw new RefusedError('not_text');
  const location = await client.redirectLocation(`${itemPath}/content`);
  const bytes = await client.downloadPreauthenticated(location, MAX_TEXT_BYTES);
  return { id: item.id ?? itemId, name, size, webUrl: item.webUrl, text: bytes.toString('utf8') };
}

/** SharePoint sites the person can reach, by name. */
export async function searchSites(client: GraphClient, query: string) {
  return (await client.get('/sites', { search: query, $select: 'id,name,displayName,webUrl' })).data;
}

/** The document libraries of a site. */
export async function listSiteDrives(client: GraphClient, siteId: string) {
  // A site id holds commas (host,site,web); they are allowed in a path segment.
  return (await client.get(`/sites/${id(siteId, 'siteId')}/drives`, { $select: 'id,name,webUrl,driveType' })).data;
}

/** Content as bytes: text as UTF-8, or base64. Exactly one, at most 4 MB. */
function contentOf(text: string | undefined, contentBase64: string | undefined): Buffer {
  if ((text === undefined) === (contentBase64 === undefined)) throw new RefusedError('content_missing');
  if (contentBase64 !== undefined && !/^[A-Za-z0-9+/\s]*={0,2}\s*$/.test(contentBase64)) throw new RefusedError('invalid_value');
  const bytes = text !== undefined ? Buffer.from(text, 'utf8') : Buffer.from(contentBase64!, 'base64');
  if (bytes.byteLength > MAX_UPLOAD_BYTES) throw new RefusedError('content_too_large');
  return bytes;
}

/** A file name for a path: no separators, no dot segments, no characters SharePoint refuses. */
function fileName(value: string): string {
  const name = value.trim();
  if (name === '' || name === '.' || name === '..' || name.length > 255 || /[\\/:*?"<>|#%\x00-\x1f]/.test(name)) {
    throw new RefusedError('invalid_value');
  }
  return encodeURIComponent(name);
}

/** A new file in a folder (the OneDrive root when none is named). An existing file of the same name is not replaced. */
export async function uploadFile(
  client: GraphClient,
  input: { name: string; driveId?: string; parentId?: string; text?: string; contentBase64?: string; contentType?: string },
) {
  const bytes = contentOf(input.text, input.contentBase64);
  const parent = input.parentId ? `items/${id(input.parentId, 'parentId')}` : 'root';
  return (await client.upload(
    `${driveBase(input.driveId)}/${parent}:/${fileName(input.name)}:/content`,
    bytes,
    uploadType(input.contentType, input.text !== undefined),
    true,
  )).data;
}

/** New content for an existing file: the file gets a new version, and the earlier ones stay in its version history. */
export async function uploadFileVersion(
  client: GraphClient,
  input: { driveId?: string; itemId: string; text?: string; contentBase64?: string; contentType?: string },
) {
  const bytes = contentOf(input.text, input.contentBase64);
  return (await client.upload(
    `${driveBase(input.driveId)}/items/${id(input.itemId, 'itemId')}/content`,
    bytes,
    uploadType(input.contentType, input.text !== undefined),
  )).data;
}

function uploadType(contentType: string | undefined, text: boolean): string {
  if (contentType === undefined) return text ? 'text/plain; charset=utf-8' : 'application/octet-stream';
  if (!/^[a-z0-9.+-]+\/[a-z0-9.+-]+$/i.test(contentType)) throw new RefusedError('invalid_value');
  return contentType;
}

/** A new folder (in the OneDrive root when no parent is named). An existing folder of the same name is not replaced. */
export async function createFolder(client: GraphClient, input: { name: string; driveId?: string; parentId?: string }) {
  fileName(input.name);
  const parent = input.parentId ? `items/${id(input.parentId, 'parentId')}` : 'root';
  return (await client.post(`${driveBase(input.driveId)}/${parent}/children`, {
    name: input.name.trim(),
    folder: {},
    '@microsoft.graph.conflictBehavior': 'fail',
  })).data;
}

/** Renames a file or folder, or moves it to another folder in the same drive. */
export async function moveFile(client: GraphClient, input: { driveId?: string; itemId: string; newName?: string; newParentId?: string }) {
  if (input.newName === undefined && input.newParentId === undefined) throw new RefusedError('no_change');
  if (input.newName !== undefined) fileName(input.newName);
  const parentReference = input.newParentId ? { id: assertIdLike(input.newParentId) } : undefined;
  return (await client.patch(`${driveBase(input.driveId)}/items/${id(input.itemId, 'itemId')}`, defined({ name: input.newName?.trim(), parentReference }))).data;
}

function assertIdLike(value: string): string {
  // The same rule as a path segment: an id, never a path.
  pathSegment(value, 'id');
  return value.trim();
}

// ------------------------------------------------------------------------------------------------------------- profile

/** Who the access belongs to: the person's id, name and addresses. */
export async function getMyProfile(client: GraphClient) {
  return (await client.get('/me', { $select: 'id,displayName,mail,userPrincipalName,preferredLanguage' })).data;
}
