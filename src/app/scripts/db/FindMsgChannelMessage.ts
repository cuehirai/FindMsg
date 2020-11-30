import Dexie from 'dexie';
import * as du from "../dateUtils";
import { ChatMessage, BodyType } from '@microsoft/microsoft-graph-types';
import { IFindMsgChannel } from './IFindMsgChannel';
import { IFindMsgChannelMessage } from './IFindMsgChannelMessage';
import { IFindMsgChannelMessageDb } from './IFindMsgChannelMessageDb';
import { checkFn, filterNull, fixMessageLink } from '../utils';
import { db, idx } from './Database';
import { traceAsync } from '../logger';
import * as log from '../logger';
import { FindMsgUserCache } from './FindMsgUserCache';
import { collapseWhitespace, sanitize, stripHtml } from '../purify';

const order2IdxMap = {
    author: idx.messages.channelId$author$subject,
    touched: idx.messages.channelId$touched$subject,
    subject: idx.messages.channelId$subject,
}

const nullUser = {
    id: "",
    displayName: "",
};

export enum MessageOrder {
    author = "author",
    touched = "touched",
    subject = "subject",
}

export enum Direction {
    ascending = "ascending",
    descending = "descending",
}


export class FindMsgChannelMessage {
    /**
     * Convert message db representation to message
     * @param message
     */
    private static fromDbEntity(message: IFindMsgChannelMessageDb, cache: FindMsgUserCache): IFindMsgChannelMessage {
        const {
            created, modified, deleted, synced, author: authorId,
            touched, // eslint-disable-line @typescript-eslint/no-unused-vars
            ...rest
        } = message;

        return {
            ...rest,
            created: du.numberToDate(created),
            deleted: du.numberToDate(deleted),
            modified: du.numberToDate(modified),
            synced: du.numberToDate(synced),
            authorId,
            authorName: cache.getDisplayName(authorId) ?? "",
        };
    }


    /**
     * Convert message to db representation
     * @param message
     */
    private static toDbEntity(message: IFindMsgChannelMessage): IFindMsgChannelMessageDb {
        const {
            created, modified, deleted, synced, authorId: author,
            authorName, // eslint-disable-line @typescript-eslint/no-unused-vars
            ...rest
        } = message;

        const c = du.dateToNumber(created);
        const m = du.dateToNumber(modified);
        const d = du.dateToNumber(deleted);

        return {
            ...rest,
            created: c,
            deleted: m,
            modified: d,
            synced: du.dateToNumber(synced),
            touched: Math.max(c, m, d),
            author,
        };
    }


    /**
     * Convert Microsoft Graph ChatMessage to IFindMsgChannelMessage
     * @param message
     * @param channelId
     */
    private static fromChatMessage(message: ChatMessage, channelId: string, userCache: FindMsgUserCache): IFindMsgChannelMessage | null {
        const {
            id,
            replyToId,
            createdDateTime,
            lastModifiedDateTime,
            deletedDateTime,
            from,
            subject,
            summary,
            body,
            webUrl,
        } = message;

        if (!id) {
            log.error("Ignoring message without id:", message);
            return null;
        }

        if (!createdDateTime) {
            log.error("Ignoring message without createdDateTime:", message);
            return null;
        }

        let type: BodyType;
        let content: string;
        let text: string | null;

        if (!body) {
            log.warn("Message has no body:", message);
            type = "text";
            content = "";
            text = subject ?? null;
        } else {
            if (body.contentType === "text") {
                type = "text";
                content = body.content ?? "";
                text = collapseWhitespace((subject ?? "") + " " + content).toLowerCase();
            } else if (body.contentType === "html") {
                type = "html";
                content = sanitize(body.content ?? "");
                text = collapseWhitespace((subject ?? "") + " " + stripHtml(content)).toLowerCase();
            } else {
                type = "text";
                content = "";
                text = null;
            }

            // only store text if it is different from the other fields
            if (content === text) text = null;
        }

        const author = from?.user ?? from?.device ?? from?.application ?? nullUser;

        if (author === nullUser) {
            if (deletedDateTime) {
                // this is ok like this
            } else {
                // the from field should be present even if the user was deleted.
                // It seems that this can be null if the message was sent to the channel via the channel's email address AND the sender is not a member of the team
                // This seems to be a bug that microsoft is aware of and tracking internally as indicated here:
                // https://stackoverflow.com/questions/63771540/how-to-obtain-channel-message-from-when-message-was-sent-by-email?noredirect=1#comment112801213_63771540
                log.warn("Message author not present.");
            }
        }

        if (typeof author.id !== "string") {
            log.warn("Author id not present");
            author.id = nullUser.id;
        }

        if (typeof author.displayName !== "string") {
            log.warn("Author name not present");
            author.displayName = nullUser.displayName;
        }

        const created = du.parseISO(createdDateTime);
        const modified = lastModifiedDateTime ? du.parseISO(lastModifiedDateTime) : du.invalidDate();
        const deleted = deletedDateTime ? du.parseISO(deletedDateTime) : du.invalidDate();

        const ts = du.isValid(modified) ? modified : created;

        if (typeof author.id === "string" && typeof author.displayName === "string") {
            userCache.updateUserName(author.id, author.displayName, ts);
        }

        if (!webUrl) {
            log.warn("Message has no webUrl:", message);
        }

        return {
            id,
            channelId,
            created,
            modified,
            deleted,
            synced: du.invalidDate(),
            authorId: author.id,
            authorName: author.displayName,
            body: content,
            type,
            replyToId: replyToId ?? "", // make undefined to "", for indexing
            subject: subject?.trim() ? subject : null, // distinction: make empty string null, for indexing
            summary: summary?.trim() ? summary : null,
            text,
            url: fixMessageLink(webUrl ?? ""),
        };
    }


    /**
     * Fetch a single message from the db
     * @param id
     */
    static async get(id: string): Promise<IFindMsgChannelMessage | null> {
        const result = await db.channelMessages.get(id);

        if (result) {
            const cache = await FindMsgUserCache.getInstance();
            return FindMsgChannelMessage.fromDbEntity(result, cache);
        }

        return null;
    }


    /**
     * store a single message to the to (overwrite existing)
     * @param message
     */
    static async put(message: IFindMsgChannelMessage): Promise<void> {
        await db.channelMessages.put(FindMsgChannelMessage.toDbEntity(message));
    }


    /**
     * store all messages to the db
     * @param messages
     * @param channel
     */
    static async putAll(messages: ChatMessage[], channel: IFindMsgChannel): Promise<IFindMsgChannelMessage[]> {
        if (messages.length > 0) {
            const uc = await FindMsgUserCache.getInstance();
            const msgs = messages.map(m => FindMsgChannelMessage.fromChatMessage(m, channel.id, uc))
                .filter(filterNull)
            const dbMsgs = msgs.map(FindMsgChannelMessage.toDbEntity);
            await db.channelMessages.bulkPut(dbMsgs);
            await uc.persistCache();
            return msgs;
        } else {
            return [];
        }
    }


    /**
     * Get all replies to the message
     * @param message
     */
    static async getReplies(message: IFindMsgChannelMessage): Promise<IFindMsgChannelMessage[]> {
        const result = await db.channelMessages.where({ channelId: message.channelId, replyToId: message.id }).toArray();

        if (result.length > 0) {
            const uc = await FindMsgUserCache.getInstance();
            return result.map(r => FindMsgChannelMessage.fromDbEntity(r, uc));
        }

        return [];
    }


    /**
     * Get a batch of top level messages (replyToId=null)
     * @param channel
     * @param limit
     */
    static async getTopLevelMessagesToSync(channel: IFindMsgChannel, limit: number): Promise<IFindMsgChannelMessage[]> {
        const cutoff: number = du.subMonths(du.now(), 7).getTime();
        const result = await db.channelMessages
            .where(idx.messages.channelId$replyToId$synced)
            .between([channel.id, "", Dexie.minKey], [channel.id, "", cutoff], true, true)
            .limit(limit)
            .toArray();

        if (result.length > 0) {
            const uc = await FindMsgUserCache.getInstance();
            return result.map(m => FindMsgChannelMessage.fromDbEntity(m, uc));
        }

        return [];
    }


    /**
     * Compare function to use in Array.sort()
     * sorts by last modified in descending order
     * @param a
     * @param b
     */
    static compareByTouched(a: IFindMsgChannelMessage, b: IFindMsgChannelMessage): number {
        const at = du.isValid(a.modified) ? a.modified : a.created;
        const bt = du.isValid(b.modified) ? b.modified : b.created;

        return du.differenceInMilliseconds(bt, at);
    }


    /**
     * Create a filter function that searches message text if available and subject and body if not.
     * Note: filter is always case insensitive
     * @param searchTerm the term to filter for
     * @returns a function that takes a messages and returns whether the message contains the search term
     */
    static createFilter(searchTerm: string): (m: IFindMsgChannelMessage | IFindMsgChannelMessageDb) => boolean {
        const t = searchTerm.toLowerCase();
        return ({ subject, body, text }) => typeof text === "string" ? text.includes(t) : subject?.toLowerCase().includes(t) || body.toLowerCase().includes(t);
    }


    /**
     * Super simple search
     * @param term
     * @param from display only messages that where created or modified after this
     * @param to display only messages that where created or modified before this
     * @param channelIds search only in these channels. serach all channels if empty.
     */
    @traceAsync()
    static async search(term: string, from: Date, to: Date, channelIds: Set<string>, userIds: Set<string>, cancelledCheck: checkFn): Promise<IFindMsgChannelMessage[]> {
        let messages = db.channelMessages.toCollection();

        // apply combined channel filter
        if (channelIds.size > 0) {
            messages = messages.filter(m => channelIds.has(m.channelId));
        }

        // apply date filter
        const fromValid = du.isValid(from);
        const toValid = du.isValid(to);
        if (fromValid || toValid) {
            if (fromValid && toValid && du.isAfter(from, to)) [from, to] = [du.startOfDay(to), du.endOfDay(from)];
            const fromN = from.valueOf();
            const toN = to.valueOf();

            if (fromValid && toValid) {
                messages = messages.filter(m => m.touched >= fromN && m.touched <= toN);
            } else if (fromValid) {
                messages = messages.filter(m => m.touched >= fromN);
            } else {
                messages = messages.filter(m => m.touched <= toN);
            }
        }

        // apply user filter
        if (userIds.size > 0) {
            messages = messages.filter(m => userIds.has(m.author));
        }

        // filter by search term
        // Note: filter only when the term is contains something other than pure whitespace, but actually require the whitespace when searching
        if (term.trim()) {
            messages = messages.filter(FindMsgChannelMessage.createFilter(term));
        }

        const results: IFindMsgChannelMessage[] = [];
        const uc = await FindMsgUserCache.getInstance();
        await Promise.all([
            messages.until(cancelledCheck).each(m => results.push(FindMsgChannelMessage.fromDbEntity(m, uc))),
            // delay(15000, cancelledCheck), // for testing: import { delay } from '../utils';
        ]);

        return results;
    }


    /**
     * Get a list of top level channel messages with non-empty subject
     * @param channelId
     * @param order
     * @param dir
     * @param offset
     * @param limit
     */
    @traceAsync()
    static async getTopLevelMessagesWithSubject(channelId: string, channelIds: Set<string>, order: MessageOrder, dir: Direction, offset = 0, limit = 0, filter = ""): Promise<[IFindMsgChannelMessage[], boolean]> {
        const index = order2IdxMap[order];

        const collection = db.channelMessages.where(index).between([channelId || Dexie.minKey, Dexie.minKey, Dexie.minKey], [channelId || Dexie.maxKey, Dexie.maxKey, Dexie.maxKey], true, true);

        if (dir === Direction.descending) collection.reverse();
        if (offset > 0) collection.offset(offset);
        if (limit > 0) collection.limit(limit + 1);
        if (filter.trim()) collection.filter(FindMsgChannelMessage.createFilter(filter));

        if (channelIds.size > 0){
            collection.filter(m => channelIds.has(m.channelId));
        }

        const result = await collection.toArray();

        let hasMore = false;

        if (result.length > 0) {
            const uc = await FindMsgUserCache.getInstance();

            if (result.length > limit) {
                result.pop();
                hasMore = true;
            }

            return [result.map(r => FindMsgChannelMessage.fromDbEntity(r, uc)), hasMore];
        }

        return [[], hasMore];
    }
}

