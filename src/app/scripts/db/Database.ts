// IMPORTANT NOTE: null/undefined/NaN values are NOT indexed!
// A query of the form "where property == null" is NOT possible.


import Dexie from 'dexie';
import { IFindMsgTeamDb } from './IFindMsgTeamDb';
import { IFindMsgChannelDb } from './IFindMsgChannelDb';
import { IFindMsgChannelMessageDb } from './IFindMsgChannelMessageDb';
import { IFindMsgUserDb } from './IFindMsgUserDb';
import { info, traceAsync, warn } from '../logger';
import { DbStatAggregator } from './DbStatAggregator';
import { AI } from '../appInsights';
import { collapseWhitespace, sanitize, stripHtml } from '../purify';
import { IFindMsgChatDb } from './IFindMsgChatDb';
import { IFindMsgChatMemberDb } from './IFindMsgChatMemberDb';
import { IFindMsgChatMessageDb } from './IFindMsgChatMessageDb';
import { IFindMsgImageDb } from './IFindMsgImageDb';
import { IFindMsgEventDb } from './Event/IFindMsgEventDb';
import { IFindMsgAttendeeDb } from './Attendee/IFindMsgAttendeeDb';
import IDBExportImport from 'indexeddb-export-import';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from '../../../config/AppConfig';
import * as du from "../dateUtils";
import { FileUtil } from '../fileUtil';


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

export interface ITableLastSync {
    id: string;
    lastSync: number;
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
    }),

    lastsync: Object.freeze({
        $id: nameof<ITableLastSync>(l => l.id),
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

    lastsync: Dexie.Table<ITableLastSync, string>;

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

        this.version(9).stores({
            lastsync: indexSpec(indexes.lastsync),
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
        this.lastsync = this.table('lastsync');

        const lastUser = localStorage.getItem(this.lastUserKey());
        this.userPrincipalName = lastUser?? "";
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

    /**
     * DBを使用する前に必ずログインしてください。
     * @param client Graphクライアント
     * @param userPrincipalName ログインヘルプ
     */
    async login(client: Client, userPrincipalName: string): Promise<boolean> {
        info(`▼▼▼ Database.login START ▼▼▼`);
        let res = false;

        this.msGraphClient = client;
        if (userPrincipalName != this.userPrincipalName) {
            info(`★★★ Login user changed [${this.userPrincipalName}] => [${userPrincipalName}] ★★★`);
            // アプリへの初回ログインまたはユーザが変わった場合は強制的にリロード
            this.userPrincipalName = userPrincipalName;
            localStorage.setItem(this.lastUserKey(), userPrincipalName);
            res = await this.import();
        } else {
            // 同じユーザが使い続けている場合、現在のデバイス/ブラウザで最後にエクスポートまたはインポートした日付よりも
            // エクスポートファイルの最終更新日時が新しければリロードする
            // ※他のデバイス/ブラウザで同期・エクスポートした内容を取り込む想定
            const lastExport = (): Date => {
                const res = localStorage.getItem(this.lastExportKey());
                return res? du.parseISO(res) : du.invalidDate();
            };
            const lastImport = (): Date => {
                const res = localStorage.getItem(this.lastImportKey());
                return res? du.parseISO(res) : du.invalidDate();
            };

            const latest = (lastExport() > lastImport())? lastExport() : lastImport();
            
            const file = await FileUtil.getFile(client, this.exportfileName(), this.exportfilePath());
            if (file) {
                const lastModified = file.lastModifiedDateTime? du.parseISO(file.lastModifiedDateTime) : du.invalidDate();

                info(`★★★ lastExport:[${lastExport()}] lastImport:[${lastImport()}] lastModified:[${lastModified}] ★★★`);
                if (lastModified > latest) {
                    res = await this.import();
                } else {
                    res = true;
                }
            } else {
                res = true;
            }
        }
        info(`▲▲▲ Database.login END ▲▲▲`);
        return res;
    }

    /**
     * テーブルの最終同期日時を取得します
     * @param id テーブルを識別するID
     */
    async getLastSync(id: string): Promise<Date> {
        info(`▼▼▼ getLastSync START... ID: [${id}] ▼▼▼`);
        let res = du.invalidDate();
        const rec = await this.lastsync.get(id);
        if (rec) {
            res = du.numberToDate(rec.lastSync);
        }
        info(`▲▲▲ getLastSync END... ID: [${id}] result: [${res}] ▲▲▲`);
        return res;
    }

    /**
     * テーブルの最終同期日時を保存します
     * @param id テーブルを識別するID
     * @param date 同期実行した日時
     * @param doExport trueを指定するとDBエクスポートを実施します(省略値はfalse)
     */
    async storeLastSync(id: string, date: Date, doExport?: boolean): Promise<void> {
        info(`▼▼▼ storeLastSync START... ID: [${id}] date: [${date}] doExport: [${doExport}] ▼▼▼`);
        await this.transaction("rw", this.lastsync, async() => {
            const rec: ITableLastSync = {id: id, lastSync: du.dateToNumber(date)};
            await this.lastsync.put(rec);
        });
        if (doExport?? false) {
            // 最終同期日時は常にエクスポートファイルの最終更新日時より古くなるのでエクスポートの引数としては使用しない
            // ※同一ユーザが使い続けているのに、タブを選択するたびにインポートが必要と判断されてしまうため
            await this.export();
        }
        info(`▲▲▲ storeLastSync END... ID: [${id}] ▲▲▲`);
    }

    /**
     * DBをクリアします
     */
    async clear(): Promise<boolean> {
        info(`▼▼▼ clear START ▼▼▼`);
        let res = false;
        const idbDatabase = this.backendDB();
        const clearDbCallback = async (error: any) => {
            if (!error) res = true;
        }
        // Clear処理本体
        IDBExportImport.clearDatabase(idbDatabase, clearDbCallback);
        info(`▲▲▲ clear END ▲▲▲`);
        return res
    }

    /**
     * DBの全データをOneDriveにエクスポートします
     * @param syncDatetime 最終エクスポート日時として保存する日時※省略時はnow()が保存されます
     */
    async export(syncDatetime?: Date): Promise<boolean> {
        info(`▼▼▼ export START ▼▼▼`);
        let res = false;

        if (!this.msGraphClient) {
            warn(`Database is not logged in!`);
        } else {
            const idbDatabase = this.backendDB();
            const exportCallback = async (error: any, jsonString: string) => {
                if (error) {
                    warn(`Error in exportToJsonString: [${error}]`);
                } else {
                    // info(`DB exported: [${jsonString}]`);
                    if (this.msGraphClient? await FileUtil.writeFile(this.msGraphClient, this.exportfileName(), jsonString, this.exportfilePath(), true) : false) {
                        res = true;
                        const lastExport = syncDatetime?? du.now();
                        localStorage.setItem(this.lastExportKey(), du.formatISO(lastExport));
                    }
                }
            }
            // Export処理本体    
            IDBExportImport.exportToJsonString(idbDatabase, exportCallback);
        }

        info(`▲▲▲ export END ▲▲▲`);

        return res;
    }

    /**
     * DBのデータをOneDriveからインポートします
     */
    async import(): Promise<boolean> {
        info(`▼▼▼ import START ▼▼▼`);
        let res = false;

        const idbDatabase = this.backendDB();
        const importCallback = (error: any) => {
            if (error) {
                error(`Error in importFromJsonString: [${error}]`);
            } else {
                res = true;
                localStorage.setItem(this.lastImportKey(), du.formatISO(du.now()));
                info(`DB import successfully completed!`);
            }
        }
        // インポートを実施する前に、有無を言わさずDBをクリアする
        if (this.clear()) {
            if (!this.msGraphClient) {
                warn(`Database is not logged in!`);
            } else {
                // インポートすべきファイルがあるかどうかを確認
                const file = await FileUtil.getFile(this.msGraphClient, this.exportfileName(), this.exportfilePath(), true);
                if (file) {
                    const jsonString = await FileUtil.readFile(this.msGraphClient, this.exportfileName(), this.exportfilePath());
                    // info(`Data importing into DB: [${jsonString}]`);
                    if (jsonString) {
                        // Import処理本体
                        IDBExportImport.importFromJsonString(idbDatabase, jsonString, importCallback);
                    } else {
                        warn(`DB import failed due to either file not found or read failure`);
                    }
                } else {
                    res = true;
                    info(`Exported file does not exist; import command ignored.`);
                }
            }
        }
    info(`▲▲▲ import END ▲▲▲`);

        return res;
    }

 
    // セキュリティ関連プロパティ
    private msGraphClient: Client | undefined = undefined;
    private userPrincipalName = "";
    // Export/Import用のワークプロパティ
    private lastUserKey(): string { return `${this.name}-lastUser`;}
    private lastExportKey(): string { return `${this.name}-${this.userPrincipalName}-lastExported`; }
    private lastImportKey(): string { return `${this.name}-${this.userPrincipalName}-lastImported`; }
    private accountDomain(): string {
        let res = "";
        if (this.userPrincipalName.indexOf("@") > 0) {
            res = this.userPrincipalName.split("@")[1];
        }
        return res;
    }
    private accountName(): string {
        let res = "";
        if (this.userPrincipalName.indexOf("@") > 0) {
            res = this.userPrincipalName.split("@")[0];
        }
        return res;
    }
    private exportfileName(): string { return "db.dat";}
    private exportfilePath(): string { return `AppData/${AppConfig.AppInfo.name}/${this.accountDomain()}/${this.accountName()}`;}
    
}

export const db = new Database(`${AppConfig.AppInfo.name}-database`);
export const idx = indexes;
