import { AadUserConversationMember, ConversationMember } from '@microsoft/microsoft-graph-types-beta';
import { FindMsgUserCache } from './FindMsgUserCache';
import { IFindMsgChat } from './IFindMsgChat';
import { IFindMsgChatMemberDb } from './IFindMsgChatMemberDb';
import { assert } from '../utils';
import { db } from './Database';


export class FindMsgChatMember {
    public static async putAll(members: ConversationMember[], chat: IFindMsgChat): Promise<void> {
        if (members.length > 0) {
            const uc = await FindMsgUserCache.getInstance();
            const msgs = members.map(m => FindMsgChatMember.fromConversationMember(m, chat.id, uc));
            await db.chatMembers.bulkPut(msgs);
            await uc.persistCache();
        }
    }


    public static fromConversationMember(m: ConversationMember | AadUserConversationMember, chatId: string, uc: FindMsgUserCache): IFindMsgChatMemberDb {
        const uid = ('userId' in m ? m.userId : m.id) ?? "";

        if (uid && m.displayName) {
            uc.updateUserName(uid, m.displayName, new Date());
        }

        return {
            chatId,
            id: assert(m.id, nameof<ConversationMember>(m => m.id)),
            userId: ('userId' in m ? m.userId : m.id) ?? "",
        }
    }
}
