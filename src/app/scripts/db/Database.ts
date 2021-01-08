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
import { FileChooser, FileUtil } from '../fileUtil';
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

interface IExportFile {
    table: string;
    data: any[];
}

interface IFile {
    file: string;
    compressed: boolean;
}

interface IExportIndexFile {
    exported: string;
    files: IFile[];
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

    private serverFiles: IExportIndexFile | null = null;

    /**
     * DBを使用する前に必ずログインしてください。
     * @param client Graphクライアント
     * @param userPrincipalName ログインヘルプ
     */
    async login(arg: {client: Client, userPrincipalName: string, resultCallback: (result: DbLoginResult) => void}): Promise<DbLoginResult> {
        info(`▼▼▼ Database.login START ▼▼▼`);

        const userChanged = async (): Promise<void> => {
            const index = await this.getFileIndex();
            if (index) {
                this.serverFiles = index;
                arg.resultCallback("SHOULD_IMPORT");
            } else {
                this.serverFiles = null;
                arg.resultCallback("OK");
            }
        }

        const checkIndex = async (): Promise<void> => {
            // 同じユーザが使い続けている場合、現在のデバイス/ブラウザで最後にエクスポートまたはインポートした日付よりも
            // エクスポートファイルの最終更新日時が新しければリロードを勧める
            // ※他のデバイス/ブラウザで同期・エクスポートした内容を取り込む想定
            const index = await this.getFileIndex();
            if (index) {
                this.serverFiles = index;

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
                const exported = du.parseISO(index.exported);
                info(`★★★ lastExport:[${lastExport()}] lastImport:[${lastImport()}] lastModified:[${exported}] ★★★`);
                if (!du.isValid(latest) || exported > latest) {
                    info(`★★★ user had better import!! ★★★`);
                    arg.resultCallback("RECOMMEND_IMPORT");
                } else {
                    info(`★★★ should be no need to import ★★★`);
                    arg.resultCallback("OK");
                }
            } else {
                this.serverFiles = null;
                arg.resultCallback("OK");
            }
        }

        this.msGraphClient = arg.client;
        if (arg.userPrincipalName != this.userPrincipalName) {
            info(`★★★ Login user changed [${this.userPrincipalName}] => [${arg.userPrincipalName}] ★★★`);
            // アプリへの初回ログインまたはユーザが変わった場合=>SHOULD_IMPORT
            if (this.userPrincipalName !== "") {
                // ユーザが変わった場合は強制的にクリア
                await this.clear();
            }
            this.userPrincipalName = arg.userPrincipalName;
            localStorage.setItem(this.lastUserKey(), arg.userPrincipalName);
            userChanged();
        } else {
            checkIndex();
        }    
        info(`▲▲▲ Database.login END ▲▲▲`);
        return "OK";
    }

    /**
     * エクスポートインデクスファイルを読み込みます
     * @param client Graphクライアント
     */
    async getFileIndex(): Promise<IExportIndexFile | null> {
        let res: IExportIndexFile | null = null;
        if (this.msGraphClient) {
            const indexFile = await FileUtil.readFile(this.msGraphClient, this.exportIndexName(), this.exportfilePath());
            if (indexFile) {
                const json: IExportIndexFile = JSON.parse(indexFile);
                info(`***** index file found. JSON.parse result: [${JSON.stringify(json)}]`);
                res = {
                    exported: json.exported,
                    files: json.files,
                }
            } else {
                info(`***** index file not found.`);
            }    
        } else {
            warn(`Database not logged in => cannot read index file`);
        }
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
    async storeLastSync(id: string, date: Date, doExport?: boolean, progressCallback?: (tableName: string, progress: number) => void, callback?: (message: number) => void): Promise<void> {
        info(`▼▼▼ storeLastSync START... ID: [${id}] date: [${date}] doExport: [${doExport}] ▼▼▼`);
        await this.transaction("rw", this.lastsync, async() => {
            const rec: ITableLastSync = {id: id, lastSync: du.dateToNumber(date)};
            await this.lastsync.put(rec);
        });
        if (doExport?? false) {
            // 最終同期日時は常にエクスポートファイルの最終更新日時より古くなるのでエクスポートの引数としては使用しない
            // ※同一ユーザが使い続けているのに、タブを選択するたびにインポートが必要と判断されてしまうため
            await this.export({includeImages: true, progressCallback: progressCallback, callback: callback});
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
    async export(arg: {syncDatetime?: Date, exportTo?: string, includeImages: boolean, targetTables?: Dexie.Table[], progressCallback?: (tableName: string, done: number, all: number, progress: number) => void, callback?: (message: number) => void}): Promise<boolean> {
        info(`▼▼▼ export START ▼▼▼`);
        let res = false;
        let message = 0;

        if (!this.msGraphClient) {
            warn(`Database is not logged in!`);
        } else {
            const client = this.msGraphClient;
            const maxMem = 15728640;
            const target = new FileChooser;
            // テーブルの選択用にもFileChooserを流用
            const tableChooser = new FileChooser;
            // 対象テーブルが指定されている場合はファイル条件を追加しておく(条件=0件の場合はすべてのテーブル/ファイルが対象となる)
            if (arg.targetTables && arg.targetTables.length > 0) {
                target.add(this.exportIndexName(), false);
                arg.targetTables.forEach(table => {
                    if (arg.includeImages || table.name !== this.images.name) {
                        const name = `^db\\.${table.name}-\\d*\\.dat`;
                        target.add(name, true);
                        tableChooser.add(table.name, false);
                        info(`@@@@@ target file added: pattern [${name}] @@@@@`);
                    }
                })
            } else {
                this.tables.forEach(table => {
                    if (arg.includeImages || table.name !== this.images.name) {
                        tableChooser.add(table.name, false);
                    }
                })
            }

            // OneDriveに必要な容量が残っているかをチェック
            let required = 0;
            let remain = 0;
            try {
                const drive = await FileUtil.getDrive(client);
                remain = drive? drive.quota? drive.quota.remaining?? 0 : 0 : 0;

                const exportPath = arg.exportTo?? this.exportfilePath();
                const files: IFile[] = [];
                // 最新のインデックスファイルを読み込む
                const index = await this.getFileIndex();
                if (index && index.files) {
                    // 対象テーブルのファイルでないものは予めファイルリストに追加しておく
                    index.files.forEach(file => {
                        if (!target.test(file.file)) {
                            files.push({file: file.file, compressed: file.compressed});
                        }
                    })
                }

                // 既存のエクスポートファイルをバックアップ
                const bkup = await FileUtil.backupFolder(client, exportPath, "bkup", target)
    
                // Export処理本体    
                let allCount = 0
                await Promise.all(
                    this.tables.map(async table => {
                        if (tableChooser.test(table.name)) {
                            const recs = await table.count();
                            info(`@@@@@ table [${table.name}].count() = [${recs}] @@@@@`)
                            allCount += recs;
                        }
                    })    
                )
                let doneCount = 0
                const progPercent = () => {
                    let res = 0;
                    if (allCount !==0) {
                        res = Math.floor(doneCount / allCount * 100);
                    }
                    return res;
                }
    
                const exportTables = async(client: Client): Promise<boolean> => {
                    let res = true;
                    await Promise.all(this.tables.map(async (table) => {
                        if (tableChooser.test(table.name)) {
                            if ( !(await exportToFiles(client, table)) ) {
                                res = false;
                                warn(`exportToFiles failed`);
                            }    
                        }
                    }));
                    info(`@@@@@ calculated quota required: [${required}] current remaining in OneDrive: [${remain}] @@@@@`);

                    return res;
                }
    
                const exportToFiles = async(client: Client, table: Dexie.Table): Promise<boolean> => {
                    let res = true;
    
                    let arr: Array<any> = [];
                    let fileNo = 0;
                    let prevPercent = 0;
                    let prevTimestamp = du.now();
    
                    await table.each(async (rec) => {
                        arr.push(rec);
                        doneCount += 1;
                        let mem = sizeof(arr);
                        // メモリサイズが15MBを超えたらいったんファイルに出力
                        if (sizeof(arr) > maxMem) {
                            info(`@@@@@ current memory size of output record array: ${mem} @@@@@`);
                            const fileName = this.exportfileName(table.name, fileNo++);
                            // files.push(fileName);
                            const content: IExportFile = {
                                table: table.name,
                                data: arr,
                            };
                            const jsonString = JSON.stringify(content);
                            // const compressed = compress(jsonString);
                            required += jsonString.length;
                            files.push({file: fileName, compressed: false});
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
                        const timestamp = du.subSeconds(du.now(), 10);
                        const percent = progPercent();
                        if (percent >= prevPercent + 5 || timestamp > prevTimestamp || percent === 100) {
                            arg.progressCallback && arg.progressCallback(table.name, doneCount, allCount, percent);
                            prevPercent = percent;
                            prevTimestamp = du.now();
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
                        // const compressed = compress(jsonString);
                        required += jsonString.length;
                        files.push({file: fileName, compressed: false});
                        if (await FileUtil.writeFile(client, fileName, jsonString, exportPath, true)) {
                            info(`@@@@@ file [${fileName}] written into OneDrive @@@@@`);
                        } else {
                            res = false;
                            warn(`@@@@@ write file [${fileName}] into OnDrive failed @@@@@`)
                        }
                        arg.progressCallback && arg.progressCallback(table.name, doneCount, allCount, progPercent());
                    }
    
                    return res;
                }
    
                await exportTables(client).then(
                    async (value) => {
                        if (value) {
                            res = true;
                            const lastExport = arg.syncDatetime?? du.now();
    
                            const index: IExportIndexFile = {exported: du.formatISO(lastExport), files: files};
                            const indexFile = JSON.stringify(index);
                            await FileUtil.writeFile(client, this.exportIndexName(), indexFile, exportPath, true);
                            if (this.userPrincipalName != "") {
                                localStorage.setItem(this.lastExportKey(), du.formatISO(lastExport));
                            }
                        } else {
                            // バックアップが成功していたならリストア（ロールバックに近い）
                            if (bkup) {
                                await FileUtil.restoreFromBackup(client, exportPath, "bkup", target);
                            }
                            // 計算上容量が足りていなかった場合はメッセージ出力
                            if (remain < required) {
                                message = Math.ceil(required / 1024);
                            }    
                        }
                    }
                ).catch(
                    async (reason) => {
                        warn(`exportTables failed: reason [${reason}]`);
                    }
                )                    
            } catch (e) {
                warn(`export failed: error [${e}]`);
            }
        }

        arg.callback && arg.callback(message);

        info(`▲▲▲ export END ▲▲▲`);

        return res;
    }

    /**
     * DBのデータをOneDriveからインポートします
     * @param importFrom 既定のインポートパスと異なる場所からインポートしたい場合に指定するパス文字列です
     * @param getLastSync インポートが実行されたあと最終同期日時を取得することができるタイミングで実行されるコールバックです
     */
    async import(arg: {importFrom?: string, progressCallback?: (tableName: string, done: number, all: number, progress: number) => void, getLastSync?: () => void}): Promise<boolean> {
        let res = false;
        const importFrom = arg.importFrom?? this.exportfilePath();
        info(`▼▼▼ import START importFrom: [${importFrom}] ▼▼▼ `);

        try {
            const allCount = this.serverFiles? this.serverFiles.files? this.serverFiles.files.length : 0 : 0;
            let doneCount = 0
            const progPercent = () => {
                let res = 0;
                if (allCount !== 0) {
                    res = Math.floor(doneCount / allCount * 100);
                }
                return res;
            }
    
            const importContent = async(fileName: string, compressed: boolean, content: string) => {
                try {
                    let fileObj: IExportFile | null = null;
                    if (compressed) {
                        const decomp = content;
                        if (decomp) {
                            fileObj = JSON.parse(decomp);
                        } else {
                            warn(`failed to decompress data`)
                        }
                    } else {
                        fileObj = JSON.parse(content);
                    }
                    if (fileObj) {
                        if (fileObj.table && fileObj.data) {
                            info(`@@@@@ importing file [${fileName}] into table [${fileObj.table}] @@@@@`);
                            const table = this.table(fileObj.table);
                            await table.bulkPut(fileObj.data);
                            if (this.userPrincipalName != "") {
                                localStorage.setItem(this.lastImportKey(), du.formatISO(du.now()));
                            }
                            doneCount += 1;
                            arg.progressCallback && arg.progressCallback(table.name, doneCount, allCount, progPercent())
                            info(`@@@@@ file [${fileName}] imported into table [${fileObj.table}] @@@@@`);
                        }    
                    }
                } catch (e) {
                    warn(`failed to import file [${fileName}]; error [${e}]`);
                }
            }
            
            if (!this.msGraphClient) {
                warn(`Database is not logged in!`);
            } else {
                // インポートすべきファイルがあるかどうかを確認
                const client = this.msGraphClient;
                if (!(this.serverFiles) || !this.serverFiles.files || this.serverFiles.files.length == 0) {
                    info(`***** assumed as old version export files`);
                    // インデックスファイルが指定されていない/インデックスファイルにファイル一覧がない場合は旧バージョンエクスポートファイルとみなし
                    // ディレクトリ内のファイルをすべてインポートする（圧縮なし）
                    const items = await FileUtil.getItems(client, importFrom, true);
                    await Promise.all(
                        items.map(async item =>{
                            if (item.file && item.name && item.name != this.exportIndexName()) {
                                const content = await FileUtil.readFile(client, item.name, importFrom);
                                if (content) {
                                    await importContent(item.name, false, content);
                                }
                            }    
                        })
                    )    
                } else {
                    info(`***** assumed as new version (compressed) export files`);
                    // インデックスファイルのファイル一覧にデータがある場合は指定されているファイルをインポートする（圧縮：指定に従う）
                    await Promise.all(
                        this.serverFiles.files.map(async file => {
                            const content = await FileUtil.readFile(client, file.file, importFrom);
                            if (content) {
                                await importContent(file.file, file.compressed, content);
                            }
                        })
                    )
                }
                res = true;
            }                
        } catch (e) {
            warn(`import failed: e = [${e}]`);
        }

        arg.getLastSync && arg.getLastSync();
    
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
            const callback = () => {info(`callback called`);};
            oldDb.login({client: client, userPrincipalName: this.userPrincipalName, resultCallback: callback});

            // 旧DBをバックアップ
            await oldDb.export({includeImages: true, exportTo: `AppData/FindMsg/${this.accountDomain()}/${this.accountName()}`});

            // lastsyncテーブルを移行
            const createLastSynced = async(lastSyncedKey: string) => {
                const local = localStorage.getItem(lastSyncedKey);
                if (local) {
                    const lastSynced = du.parseISO(local);
                    await this.storeLastSync(lastSyncedKey, lastSynced);    
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
            await this.export({includeImages: true, callback: callback2});
        }
    }
}

export declare type DbLoginResult = "OK" | "NG" | "SHOULD_IMPORT" | "RECOMMEND_IMPORT";
export const db = new Database(`${AppConfig.AppInfo.name}-database`);
export const idx = indexes;
