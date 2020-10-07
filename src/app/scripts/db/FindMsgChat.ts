import { Chat } from '@microsoft/microsoft-graph-types-beta';
import { IFindMsgChat } from './IFindMsgChat';
import { IFindMsgChatDb } from './IFindMsgChatDb';
import { assert } from '../utils';
import { dateToNumber, numberToDate } from "../dateUtils";
import { db } from './Database';


export class FindMsgChat {
    /**
     * Convert from storage entity to
     * @param chat
     */
    static fromDbEntity(chat: IFindMsgChatDb): IFindMsgChat {
        const { createdDateTime, lastUpdatedDateTime, ...rest } = chat;
        return {
            ...rest,
            createdDateTime: numberToDate(createdDateTime),
            lastUpdatedDateTime: numberToDate(lastUpdatedDateTime),
        };
    }


    /**
     * Convert to storage entity
     * @param chat
     */
    static toDbEntity(chat: IFindMsgChat): IFindMsgChatDb {
        const { createdDateTime, lastUpdatedDateTime, ...rest } = chat;
        return {
            ...rest,
            createdDateTime: dateToNumber(createdDateTime),
            lastUpdatedDateTime: dateToNumber(lastUpdatedDateTime),
        };
    }


    /**
     * Convert from Microsoft Graph API Chat
     * @param chat
     * @param getSyncDateFromDb
     */
    static async fromChat(chat: Chat): Promise<IFindMsgChat> {

        const id = assert(chat.id, nameof(chat.id));
        const syncDate = -1;

        return {
            id,
            topic: chat.topic ?? null,
            createdDateTime: numberToDate(syncDate),
            lastUpdatedDateTime: numberToDate(syncDate),
        };
    }


    /**
     * Get chat from DB by IDs
     * @param id
     */
    static async get(id: string): Promise<IFindMsgChat | null> {
        const result = await db.chats.get(id);
        return result ? FindMsgChat.fromDbEntity(result) : null;
    }



    /**
     * Get all chats from DB
     */
    static async getAll(): Promise<IFindMsgChat[]> {
        return (await db.chats.toArray()).map(FindMsgChat.fromDbEntity);
    }



    /**
     * Store chat in DB (overwrite existing)
     * @param chat
     */
    static async put(chat: IFindMsgChat): Promise<void> {
        await db.chats.put(FindMsgChat.toDbEntity(chat));
    }
}
