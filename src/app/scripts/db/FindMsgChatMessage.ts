import * as du from "../dateUtils";
import { ChatMessage, BodyType } from '@microsoft/microsoft-graph-types-beta';
import { IFindMsgChat } from './IFindMsgChat';
import { IFindMsgChatMessage } from './IFindMsgChatMessage';
import { IFindMsgChatMessageDb } from './IFindMsgChatMessageDb';
import { checkFn, filterNull } from '../utils';
import { db } from './Database';
import { traceAsync } from '../logger';
import * as log from '../logger';
import { FindMsgUserCache } from './FindMsgUserCache';
import { collapseWhitespace, sanitize, stripHtml } from '../purify';

const nullUser = {
    id: "",
    displayName: "",
};

export enum MessageOrder {
    author = "author",
    subject = "subject",
}

export enum Direction {
    ascending = "ascending",
    descending = "descending",
}


export class FindMsgChatMessage {
    /**
     * Convert message db representation to message
     * @param message
     */
    private static fromDbEntity(message: IFindMsgChatMessageDb, cache: FindMsgUserCache): IFindMsgChatMessage {
        const {
            created, modified, deleted, authorId: authorId,
            ...rest
        } = message;

        return {
            ...rest,
            created: du.numberToDate(created),
            deleted: du.numberToDate(deleted),
            modified: du.numberToDate(modified),
            authorId,
            authorName: cache.getDisplayName(authorId) ?? "",
        };
    }


    /**
     * Convert message to db representation
     * @param message
     */
    private static toDbEntity(message: IFindMsgChatMessage): IFindMsgChatMessageDb {
        const {
            created, modified, deleted, authorId: authorId,
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
            authorId,
        };
    }


    /**
     * Convert Microsoft Graph ChatMessage to IFindMsgChatMessage
     * @param message
     * @param chatId
     */
    private static fromChatMessage(message: ChatMessage, chatId: string, userCache: FindMsgUserCache): IFindMsgChatMessage | null {
        const {
            id,
            createdDateTime,
            lastModifiedDateTime,
            deletedDateTime,
            from,
            subject,
            body,
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
                // It seems that this can be null if the message was sent to the chat via the chat's email address AND the sender is not a member of the team
                // This seems to be a bug that microsoft is aware of and tracking internally as indicated here:
                // https://stackoverflow.com/questions/63771540/how-to-obtain-chat-message-from-when-message-was-sent-by-email?noredirect=1#comment112801213_63771540
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

        return {
            id,
            chatId,
            created,
            modified,
            deleted,
            authorId: author.id,
            authorName: author.displayName,
            body: content,
            type,
            text,
        };
    }


    /**
     * Fetch a single message from the db
     * @param id
     */
    static async get(id: string): Promise<IFindMsgChatMessage | null> {
        const result = await db.chatMessages.get(id);

        if (result) {
            const cache = await FindMsgUserCache.getInstance();
            return FindMsgChatMessage.fromDbEntity(result, cache);
        }

        return null;
    }


    /**
     * store a single message to the to (overwrite existing)
     * @param message
     */
    static async put(message: IFindMsgChatMessage): Promise<void> {
        await db.chatMessages.put(FindMsgChatMessage.toDbEntity(message));
    }


    /**
     * store all messages to the db
     * @param messages
     * @param chat
     */
    static async putAll(messages: ChatMessage[], chat: IFindMsgChat): Promise<IFindMsgChatMessage[]> {
        if (messages.length > 0) {
            const uc = await FindMsgUserCache.getInstance();
            const msgs = messages.map(m => FindMsgChatMessage.fromChatMessage(m, chat.id, uc))
                .filter(filterNull);
            const dbMsgs = msgs.map(FindMsgChatMessage.toDbEntity);
            await db.chatMessages.bulkPut(dbMsgs);
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
    static async getReplies(message: IFindMsgChatMessage): Promise<IFindMsgChatMessage[]> {
        const result = await db.chatMessages.where({ chatId: message.chatId, replyToId: message.id }).toArray();

        if (result.length > 0) {
            const uc = await FindMsgUserCache.getInstance();
            return result.map(r => FindMsgChatMessage.fromDbEntity(r, uc));
        }

        return [];
    }


    /**
     * Compare function to use in Array.sort()
     * sorts by last modified in descending order
     * @param a
     * @param b
     */
    static compareByTouched(a: IFindMsgChatMessage, b: IFindMsgChatMessage): number {
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
    static createFilter(searchTerm: string): (m: IFindMsgChatMessage | IFindMsgChatMessageDb) => boolean {
        const t = searchTerm.toLowerCase();
        return ({ body, text }) => typeof text === "string" ? text.includes(t) : body.toLowerCase().includes(t);
    }


    /**
     * Super simple search
     * @param term
     * @param from display only messages that where created or modified after this
     * @param to display only messages that where created or modified before this
     * @param chatIds search only in these chats. serach all chats if empty.
     */
    @traceAsync()
    static async search(term: string, from: Date, to: Date, chatIds: Set<string>, userIds: Set<string>, cancelledCheck: checkFn): Promise<IFindMsgChatMessage[]> {
        let messages = db.chatMessages.toCollection();

        // apply combined chat filter
        if (chatIds.size > 0) {
            messages = messages.filter(m => chatIds.has(m.chatId));
        }

        // apply date filter
        const fromValid = du.isValid(from);
        const toValid = du.isValid(to);
        if (fromValid || toValid) {
            if (fromValid && toValid && du.isAfter(from, to)) [from, to] = [du.startOfDay(to), du.endOfDay(from)];
            const fromN = from.valueOf();
            const toN = to.valueOf();

            if (fromValid && toValid) {
                messages = messages.filter(m => m.created >= fromN && m.created <= toN);
            } else if (fromValid) {
                messages = messages.filter(m => m.created >= fromN);
            } else {
                messages = messages.filter(m => m.created <= toN);
            }
        }

        // apply user filter
        if (userIds.size > 0) {
            messages = messages.filter(m => userIds.has(m.authorId));
        }

        // filter by search term
        // Note: filter only when the term is contains something other than pure whitespace, but actually require the whitespace when searching
        if (term.trim()) {
            messages = messages.filter(FindMsgChatMessage.createFilter(term));
        }

        const results: IFindMsgChatMessage[] = [];
        const uc = await FindMsgUserCache.getInstance();
        await Promise.all([
            messages.until(cancelledCheck).each(m => results.push(FindMsgChatMessage.fromDbEntity(m, uc))),
            // delay(15000, cancelledCheck), // for testing: import { delay } from '../utils';
        ]);

        return results;
    }
}

