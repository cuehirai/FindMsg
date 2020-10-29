// IMPORTANT NOTE: null/undefined/NaN values are NOT indexed!
// A query of the form "where property == null" is NOT possible.


import Dexie from 'dexie';
import { IFindMsgTeamDb } from './IFindMsgTeamDb';
import { IFindMsgChannelDb } from './IFindMsgChannelDb';
import { IFindMsgChannelMessageDb } from './IFindMsgChannelMessageDb';
import { IFindMsgUserDb } from './IFindMsgUserDb';
import { info, traceAsync } from '../logger';
import { DbStatAggregator } from './DbStatAggregator';
import { AI } from '../appInsights';
import { collapseWhitespace, sanitize, stripHtml } from '../purify';
import { IFindMsgChatDb } from './IFindMsgChatDb';
import { IFindMsgChatMemberDb } from './IFindMsgChatMemberDb';
import { IFindMsgChatMessageDb } from './IFindMsgChatMessageDb';
import { IFindMsgImageDb } from './IFindMsgImageDb';
import { IFindMsgEventDb } from './Event/IFindMsgEventDb';
import { IFindMsgAttendeeDb } from './Attendee/IFindMsgAttendeeDb';

/**
 * エンティティ(テーブル)名リソース用インターフェース
 * ※テーブルを追加時に随時登録
 */
export interface IEntityNames {
    teams: string;
    channels: string;
    messages: string;
    users: string;
    chats: string;
    chatMembers: string;
    images: string;
    events: string;
    attendees: string;
}

/**
 * Generate a compound index definition
 * @param args
 */
const compound = (...args: string[]) => `[${args.join("+")}]`;


/**
 * Generate an index specifier for dexie
 * Sort to make sure $ comes in front, since in dexie the first index is the primary index.
 * @param def
 */
const indexSpec = (def: { [key: string]: string }): string => Object.keys(def).sort().map(key => def[key]).join(", ");


/**
 * Indexes defined on the database
 * The index beginning with $ is the primary index.
 * $ in the middle is used as a property separator for compound indexes
 * IMPORTANT NOTE: if indexes are changed, increase database version. see https://dexie.org/docs/Tutorial/Design#database-versioning
 */
const indexes = Object.freeze({
    /**
     * Indexes on the teams store
     */
    teams: Object.freeze({
        $id: nameof<IFindMsgTeamDb>(t => t.id),
    }),

    /**
     * Indexes on the channels store
     */
    channels: Object.freeze({
        $id: nameof<IFindMsgChannelDb>(c => c.id),
        teamId: nameof<IFindMsgChannelDb>(c => c.teamId),
    }),

    /**
     * Indexes on the messages store
     */
    messages: Object.freeze({
        $channelId$id: compound(nameof<IFindMsgChannelMessageDb>(m => m.channelId), nameof<IFindMsgChannelMessageDb>(m => m.id)),

        channelId$replyToId$synced: compound(nameof<IFindMsgChannelMessageDb>(m => m.channelId), nameof<IFindMsgChannelMessageDb>(m => m.replyToId), nameof<IFindMsgChannelMessageDb>(m => m.synced)),

        // extra subject index item is to ensure that items with null subject are ignored
        channelId$touched$subject: compound(nameof<IFindMsgChannelMessageDb>(m => m.channelId), nameof<IFindMsgChannelMessageDb>(m => m.touched), nameof<IFindMsgChannelMessageDb>(m => m.subject)),
        channelId$author$subject: compound(nameof<IFindMsgChannelMessageDb>(m => m.channelId), nameof<IFindMsgChannelMessageDb>(m => m.author), nameof<IFindMsgChannelMessageDb>(m => m.subject)),
        channelId$subject: compound(nameof<IFindMsgChannelMessageDb>(m => m.channelId), nameof<IFindMsgChannelMessageDb>(m => m.subject)),
    }),

    /**
     * Indexes on the users store
     */
    users: Object.freeze({
        $id: nameof<IFindMsgUserDb>(u => u.id),
    }),

    chats: Object.freeze({
        $id: nameof<IFindMsgChatDb>(c => c.id),
    }),

    chatMessages: Object.freeze({
        $chatId$id: compound(nameof<IFindMsgChatMessageDb>(m => m.chatId), nameof<IFindMsgChatMessageDb>(m => m.id)),
    }),

    chatMembers: Object.freeze({
        $chatId$id: compound(nameof<IFindMsgChatMemberDb>(m => m.chatId), nameof<IFindMsgChatMemberDb>(m => m.id)),
    }),

    images: Object.freeze({
        $id: nameof<IFindMsgImageDb>(i => i.id),
    }),

    events: Object.freeze({
        $id: nameof<IFindMsgEventDb>(e => e.id),

        organizer$start$subject: compound(nameof<IFindMsgEventDb>(m => m.organizerName), nameof<IFindMsgEventDb>(m => m.start), nameof<IFindMsgEventDb>(m => m.subject)),
        start$subject: compound(nameof<IFindMsgEventDb>(m => m.start), nameof<IFindMsgEventDb>(m => m.subject)),
        subject: nameof<IFindMsgEventDb>(m => m.subject),
    }),

    attendees: Object.freeze({
        $eventId$id: compound(nameof<IFindMsgAttendeeDb>(a => a.eventId), nameof<IFindMsgAttendeeDb>(a => a.id)),
    })
});


/**
 * App database
 */
class Database extends Dexie {
    teams: Dexie.Table<IFindMsgTeamDb, string>;
    channels: Dexie.Table<IFindMsgChannelDb, string>;
    channelMessages: Dexie.Table<IFindMsgChannelMessageDb, string>;
    users: Dexie.Table<IFindMsgUserDb, string>;

    chats: Dexie.Table<IFindMsgChatDb, string>;
    chatMembers: Dexie.Table<IFindMsgChatMemberDb, string>;
    chatMessages: Dexie.Table<IFindMsgChatMessageDb, string>;

    events: Dexie.Table<IFindMsgEventDb, string>;
    attendees: Dexie.Table<IFindMsgAttendeeDb, string>;

    /** Stores images attached to messages */
    images: Dexie.Table<IFindMsgImageDb, string>;

    constructor(dbName: string) {
        super(dbName);

        this.version(3).stores({
            teams: indexSpec(indexes.teams),
            channels: indexSpec(indexes.channels),
            messages: indexSpec(indexes.messages),
            users: indexSpec(indexes.users),
        }).upgrade(tx => {
            Database._onUpgrade(3);
            (<Dexie.Table<IFindMsgChannelMessageDb, string>>tx.table('messages')).toCollection().modify(m => {
                if (m.type === "html") {
                    m.body = sanitize(m.body ?? "");
                    m.text = collapseWhitespace((m.subject || "") + " " + stripHtml(m.body).toLowerCase());
                } else if (m.type === "text") {
                    m.text = (m.subject?.toLowerCase() ?? "") + m.body.toLowerCase();
                } else {
                    m.type = "text";
                    m.body = "";
                    m.text = null;
                }
            })
        });

        this.version(4).stores({
            teams: indexSpec(indexes.teams),
            channels: indexSpec(indexes.channels),
            messages: indexSpec(indexes.messages),
            users: indexSpec(indexes.users),
        }).upgrade(tx => {
            Database._onUpgrade(4);
            (<Dexie.Table<IFindMsgChannelMessageDb, string>>tx.table('messages')).toCollection().modify(m => {
                if (!m.subject?.trim()) m.subject = null;
            })
        });

        this.version(5).stores({
            chats: indexSpec(indexes.chats),
            chatMembers: indexSpec(indexes.chatMembers),
            chatMessages: indexSpec(indexes.chatMessages),
        });

        this.version(6).stores({}).upgrade(tx => {
            Database._onUpgrade(6);
            (<Dexie.Table<IFindMsgChannelMessageDb, string>>tx.table('messages')).toCollection().modify(m => {
                if (m.type === "html") {
                    m.body = sanitize(m.body ?? "");
                    m.text = collapseWhitespace((m.subject ?? "") + " " + stripHtml(m.body).toLowerCase());
                } else if (m.type === "text") {
                    m.text = collapseWhitespace((m.subject ?? "") + " " + m.body).toLowerCase();
                } else {
                    m.type = "text";
                    m.body = "";
                    m.text = null;
                }
            })
        });

        this.version(7).stores({
            images: indexSpec(indexes.images),
        });

        this.version(8).stores({
            events: indexSpec(indexes.events),
            attendees: indexSpec(indexes.attendees),
        });

        this.teams = this.table('teams');
        this.channels = this.table('channels');
        this.channelMessages = this.table('messages');
        this.users = this.table('users');
        this.chats = this.table('chats');
        this.chatMembers = this.table('chatMembers');
        this.chatMessages = this.table('chatMessages');
        this.images = this.table('images');
        this.events = this.table('events');
        this.attendees = this.table('attendees');
    }

    private static _onUpgrade(version: number) {
        AI.trackEvent({
            name: "DB_upgrade",
            properties: { version }
        });
    }

    @traceAsync()
    async stats() {
        const statagg = new DbStatAggregator();
        await statagg.analyzeDb(this);
    }

    @traceAsync()
    async messageStats() {
        info(`Checking message table...`);

        let count = 0;
        let topCount = 0;
        let len = 0;
        let minLen = Infinity;
        let maxLen = 0;

        await this.channelMessages.each(msg => {
            count++;
            if (!msg.replyToId) topCount++;
            const l = msg.body.length + (msg.subject?.length ?? 0);
            len += l;
            minLen = Math.min(minLen, l);
            maxLen = Math.max(maxLen, l);
        });

        info(`${count} messages (${topCount} top level)`);
        info(`Average message length: ${(len / count).toFixed(2)}`);
        info(`Minimum message length: ${minLen}`);
        info(`Maximum message length: ${maxLen}`);
    }

    async d_list(): Promise<void> {
        const teams = await this.teams.toArray();

        for (const team of teams) {
            console.info(`${team.id}   ${team.displayName}`);
            await this.channels.filter(c => c.teamId === team.id).each(c => console.info(`   ${c.id}   ${c.displayName}`));
        }
    }

    async d_resetTeamSynced(tid: string): Promise<void> {
        await this.teams.where(indexes.teams.$id).equals(tid).modify(t => { t.lastChannelListSync = -1 });
    }

    async d_resetChannelSynced(cid: string, full = false): Promise<void> {
        await this.channels.where(indexes.channels.$id).equals(cid).modify(c => {
            if (full) c.lastFullMessageSync = -1;
            c.lastDeltaUpdate = -1;
        });
    }

    async d_delChannelMessages(cid: string): Promise<void> {
        await this.d_resetChannelSynced(cid, true);
        await this.channelMessages.where(indexes.messages.$channelId$id).between([cid, Dexie.minKey], [cid, Dexie.maxKey]).delete();
    }
}


export const db = new Database("FindMsg-database");
export const idx = indexes;
