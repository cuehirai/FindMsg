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
import { IImageDb } from './Image/IImageDb';
import { IFindMsgEventDb } from './Event/IFindMsgEventDb';
import { IFindMsgAttendeeDb } from './Attendee/IFindMsgAttendeeDb';
import IDBExportImport from 'indexeddb-export-import';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from '../../../config/AppConfig';
import * as du from "../dateUtils";
import { FileUtil } from '../fileUtil';
import sizeof from 'object-sizeof';
import { ImageTable } from './Image/ImageEntity';
import { b64toBlob } from '../utils';


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

// interface ITableIndex {
//     name: string;
//     files: Array<string>;
// }

// interface ITableOfContent {
//     exported: string;
//     tables: Array<ITableIndex>;
// }

interface IExportFile {
    table: string;
    data: any[];
}

interface IExportIndexFile {
    exported: string;
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
        $id: nameof<IImageDb>(i => i.id),
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
    images: Dexie.Table<IImageDb, string>;

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
     * 戻り値は、インポートを実行してその結果が成功の場合のみTrueです。Falseが戻ってもエラーというわけではありません。
     * @param client Graphクライアント
     * @param userPrincipalName ログインヘルプ
     * @param avoidReload 自動リロードの判定を回避するかどうか※特殊なケースを除きこのパラメータは指定しないでください。
     * @param getLastSync インポートが実行されたあと最終同期日時を取得することができるタイミングで実行されるコールバックです。
     */
    async login(arg: {client: Client, userPrincipalName: string, avoidReload?: boolean, getLastSync?: () => void}): Promise<boolean> {
        info(`▼▼▼ Database.login START ▼▼▼`);
        let res = false;

        this.msGraphClient = arg.client;
        if ( !(arg.avoidReload?? false)) {
            if (arg.userPrincipalName != this.userPrincipalName) {
                info(`★★★ Login user changed [${this.userPrincipalName}] => [${arg.userPrincipalName}] ★★★`);
                // アプリへの初回ログインまたはユーザが変わった場合は強制的にリロード
                this.userPrincipalName = arg.userPrincipalName;
                localStorage.setItem(this.lastUserKey(), arg.userPrincipalName);
                res = await this.import({getLastSync: arg.getLastSync});
            } else {
                // 同じユーザが使い続けている場合、現在のデバイス/ブラウザで最後にエクスポートまたはインポートした日付よりも
                // エクスポートファイルの最終更新日時が新しければリロードする
                // ※他のデバイス/ブラウザで同期・エクスポートした内容を取り込む想定
                const lastExport = (): Date => {
                    const local = localStorage.getItem(this.lastExportKey());
                    return local? du.parseISO(local) : du.invalidDate();
                };
                const lastImport = (): Date => {
                    const local = localStorage.getItem(this.lastImportKey());
                    return local? du.parseISO(local) : du.invalidDate();
                };
    
                let latest = du.invalidDate();
                if (du.isValid(lastExport())) {
                    if (du.isValid(lastImport())) {
                        latest = (lastExport() > lastImport())? lastExport() : lastImport();
                    } else {
                        latest = lastExport();
                    }
                } else {
                    if (du.isValid(lastImport())) {
                        latest = lastImport();
                    }
                }
                
                const indexFile = await FileUtil.readFile(arg.client, this.exportIndexName(), this.exportfilePath());
                if (indexFile) {
                    try {
                        const index: IExportIndexFile = JSON.parse(indexFile);
                        if (index.exported) {
                            const exported = du.parseISO(index.exported);
                            info(`★★★ lastExport:[${lastExport()}] lastImport:[${lastImport()}] lastModified:[${exported}] ★★★`);
                            if (!du.isValid(latest) || exported > latest) {
                                info(`★★★ executing db import!! ★★★`);
                                res = await this.import({getLastSync: arg.getLastSync});
                            } else {
                                info(`★★★ should be no need to import ★★★`);
                            }
                        }
                    } catch (e) {
                        warn(`★★★ ${this.exportIndexName()} is invalid... Import is not performed. ★★★`);
                    }
                } else {
                    warn(`★★★ ${this.exportIndexName()} not found... Import is not performed. ★★★`);
                }            
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
            await this.export({});
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
     * @param exportTo 既定のエクスポートパスと異なる場所に保存したい場合に指定するパス文字列です
     * @param callback 全テーブルのエクスポートが完了した時点で呼び出されるコールバックです
     */
    async export(arg: {syncDatetime?: Date, exportTo?: string, callback?: () => void}): Promise<boolean> {
        info(`▼▼▼ export START ▼▼▼`);
        let res = false;

        if (!this.msGraphClient) {
            warn(`Database is not logged in!`);
        } else {
            const exportPath = arg.exportTo?? this.exportfilePath();

            // Export処理本体    
            const client = this.msGraphClient;

            const exportTables = async(client: Client): Promise<boolean> => {
                let res = true;
                for (let i = 0; i < this.tables.length; i++) {
                    const table = this.tables[i];
                    await exportToFiles(client, table).then(
                        (value) => {
                            if (i == this.tables.length - 1) {
                                arg.callback && arg.callback();
                            }
                            if (!value) {
                                res = false;
                                warn(`exportToFiles failed`);
                            }
                        }
                    ).catch(
                        (reason) => {
                            res = false;
                            warn(`exportToFiles failed: reason [${reason}]`);
                        }
                    )

                }
                return res;
            }

            const exportToFiles = async(client: Client, table: Dexie.Table): Promise<boolean> => {
                let res = true;

               let arr: Array<any> = [];
                let fileNo = 0;

                await table.each(async (rec) => {
                    arr.push(rec);
                    let mem = sizeof(arr);
                    // 20MBを超えたらいったんファイルに出力
                    if (sizeof(arr) > 20971520) {
                        info(`@@@@@ current memory size of output record array: ${mem} @@@@@`);
                        const fileName = this.exportfileName(table.name, fileNo++);
                        // files.push(fileName);
                        const content: IExportFile = {
                            table: table.name,
                            data: arr,
                        };
                        const jsonString = JSON.stringify(content);
                        arr = [];
                        mem = sizeof(arr);
                        info(`@@@@@ memory size of output record array after clear: ${mem} @@@@@`);
                        if (await FileUtil.writeFile(client, fileName, jsonString, exportPath, true)) {
                            info(`@@@@@ file [${fileName}] written into OneDrive @@@@@`);
                        } else {
                            res = false;
                            warn(`@@@@@ write file [${fileName}] into OnDrive failed @@@@@`)
                        }
                    }
                })
                if (arr.length > 0) {
                    const fileName = this.exportfileName(table.name, fileNo++);
                    // files.push(fileName);
                    const content: IExportFile = {
                        table: table.name,
                        data: arr,
                    };
                    const jsonString = JSON.stringify(content);
                    if (await FileUtil.writeFile(client, fileName, jsonString, exportPath, true)) {
                        info(`@@@@@ file [${fileName}] written into OneDrive @@@@@`);
                    } else {
                        res = false;
                        warn(`@@@@@ write file [${fileName}] into OnDrive failed @@@@@`)
                    }
                }

                return res;
            }

            exportTables(client).then(
                async (value) => {
                    if (value) {
                        res = true;
                        const lastExport = arg.syncDatetime?? du.now();
                        if (this.userPrincipalName != "") {
                            localStorage.setItem(this.lastExportKey(), du.formatISO(lastExport));
                        }

                        const index: IExportIndexFile = {exported: du.formatISO(lastExport)};
                        const indexFile = JSON.stringify(index);
                        await FileUtil.writeFile(client, this.exportIndexName(), indexFile, exportPath, true);
                    }
                }
            ).catch(
                (reason) => {
                    warn(`exportTables failed: reason [${reason}]`);
                }
            )    
        }

        info(`▲▲▲ export END ▲▲▲`);

        return res;
    }

    /**
     * DBのデータをOneDriveからインポートします
     * @param importFrom 既定のインポートパスと異なる場所からインポートしたい場合に指定するパス文字列です
     * @param getLastSync インポートが実行されたあと最終同期日時を取得することができるタイミングで実行されるコールバックです
     */
    async import(arg: {importFrom?: string, getLastSync?: () => void}): Promise<boolean> {
        info(`▼▼▼ import START ▼▼▼`);
        let res = false;
        const importFrom = arg.importFrom?? this.exportfilePath();

        // インポートを実施する前に、有無を言わさずDBをクリアする
        if (this.clear()) {
            if (!this.msGraphClient) {
                warn(`Database is not logged in!`);
            } else {
                // インポートすべきファイルがあるかどうかを確認
                const client = this.msGraphClient;
                const items = await FileUtil.getItems(client, importFrom, true);
                let index = 0
                await Promise.all(
                    items.map(async item =>{
                        index += 1;
                        if (item.file && item.name && item.name != this.exportIndexName()) {
                            const content = await FileUtil.readFile(client, item.name, importFrom);
                            if (content) {
                                try {
                                    const fileObj: IExportFile = JSON.parse(content);
                                    if (fileObj.table && fileObj.data) {
                                        info(`@@@@@ importing file [${item.name}] into table [${fileObj.table}] @@@@@`);
                                        const table = this.table(fileObj.table);
                                        await table.bulkPut(fileObj.data);
                                        if (this.userPrincipalName != "") {
                                            localStorage.setItem(this.lastImportKey(), du.formatISO(du.now()));
                                        }
                                        info(`@@@@@ file [${item.name}] imported into table [${fileObj.table}] @@@@@`);
                                        if (index > items.length) {
                                            arg.getLastSync && arg.getLastSync();
                                        }
                                    }    
                                } catch (e) {
                                    warn(`failed to import file [${item.name}]`);
                                }
                            }
                        }    
                    })
                )

                res = true;
            }
        }
        info(`▲▲▲ import END ▲▲▲`);

        return res;
    }

 
    // セキュリティ関連プロパティ
    private msGraphClient: Client | undefined = undefined;
    /** このデバイス/ブラウザで最後にこのアプリを使用したユーザ */
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
    private exportIndexName(): string {return `_index.dat`;}
    private exportfileName(table: string, no: number): string { return `db.${table}-${no}.dat`;}
    private exportfilePath(): string { return `AppData/${AppConfig.AppInfo.name}/${this.accountDomain()}/${this.accountName()}`;}
    
    async importFromFindMsg(callback1: () => void, callback2: () => void): Promise<void> {
        if (this.msGraphClient) {
            await this.clear();

            // 旧DBインスタンスを生成
            const client = this.msGraphClient;
            const oldDb = new Database("FindMsg-database");
            oldDb.login({client: client, userPrincipalName: this.userPrincipalName, avoidReload: true});

            // 旧DBをバックアップ
            await oldDb.export({exportTo: `AppData/FindMsg/${this.accountDomain()}/${this.accountName()}`});

            // lastsyncテーブルを移行
            const createLastSynced = async(lastSyncedKey: string) => {
                const local = localStorage.getItem(lastSyncedKey);
                if (local) {
                    const lastSynced = du.parseISO(local);
                    await this.storeLastSync(lastSyncedKey, lastSynced, false);    
                }
            }
            await this.teams.bulkPut(await oldDb.teams.toArray());
            await this.channels.bulkPut(await oldDb.channels.toArray());
            await this.channelMessages.bulkPut(await oldDb.channelMessages.toArray());
            await this.users.bulkPut(await oldDb.users.toArray());
            await this.chats.bulkPut(await oldDb.chats.toArray());
            await this.chatMembers.bulkPut(await oldDb.chatMembers.toArray());
            await this.chatMessages.bulkPut(await oldDb.chatMessages.toArray());
            await this.events.bulkPut(await oldDb.events.toArray());
            await this.attendees.bulkPut(await oldDb.attendees.toArray());
            await oldDb.images.each( async(image) => {
                const srcUrl = image.srcUrl;
                const data = image.dataUrl? b64toBlob(image.dataUrl) : image.data;
                const imgdb: IImageDb ={
                    id: image.id,
                    srcUrl: srcUrl,
                    fetched: image.fetched,
                    data: data,
                    dataUrl: image.dataUrl,
                    dataChunk: [],
                    parentId: null,
                }
                await ImageTable.put(imgdb);
            })

            await createLastSynced("FindMsg_teams_last_synced");
            await createLastSynced("FindMsg_toplevel_messages_last_synced");
            await createLastSynced("FindMsgSearch_last_synced");
            await createLastSynced("FindMsg_events_last_synced");
            await createLastSynced("FindMsg_attendees_last_synced");
            await createLastSynced("FindMsgSearchChat_last_synced");
            callback1();

            // 移行済みDBをエクスポート
            await this.export({callback: callback2});
        }
    }
}

export const db = new Database(`${AppConfig.AppInfo.name}-database`);
export const idx = indexes;
