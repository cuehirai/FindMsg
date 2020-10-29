import { Client } from "@microsoft/microsoft-graph-client";
import { db } from "./Database";
import * as du from "../dateUtils";
import * as log from '../logger';
import { progressFn, throwFn } from "../utils";
import { IMessageTranslation } from "../i18n/IMessageTranslation";
import { AI } from "../appInsights";

export class SyncError extends Error {
    constructor(message: string) {
        super(message);
    }
}

export enum OrderByDirection {
    ascending = "ascending",
    descending = "descending",
}


/** 従属テーブルの操作を行うメソッド */
export type SubEntityFunction = {
    (arg: ISubEntityFunctionArg): Promise<void>;
}

/** 従属テーブルの操作を行うメソッド(一括) */
export type SubEntytyAllFunction = {
    (args: ISubEntityFunctionArg[]): Promise<void>;
}

export type SyncFunction = {
    (arg: ISyncFunctionArg): Promise<boolean>;
}

/** 従属テーブル操作メソッドの引数ベース */
export interface ISubEntityFunctionArg {
    parent: ITeamsEntityBase | null;
    parentDb: IDbEntityBase | null;
}

/** DBレコードベース */
export interface IDbEntityBase {
    id: string;
}

/** アプリ用に成形されたレコードベース */
export interface ITeamsEntityBase {
    id: string;
}

/** 同期処理メソッドの引数ベース */
export interface ISyncFunctionArg {
    /** graphクライアント */
    client: Client;
    /** キャンセル時処理 */
    checkCancel: throwFn;
    /** 進捗表示処理 */
    progress: progressFn;
    /** サブエンティティも同期するか */
    subentity?: boolean;
    /** サブエンティティとして同期している場合の親エンティティレコード */
    parent?: ITeamsEntityBase;
    /** メッセージ辞書 */
    translate: IMessageTranslation;
}

/** DBアクセスクラスの基本形 */
interface IDbAccessorClasBase<D, T, A>{
    /** 対象のテーブル名 */
    tableName: string;

    /** 最終同期日時を保存するキー */
    lastSyncedKey: string;

    /** デルタ同期が可能かどうか */
    isDeltaSyncAvailable: boolean;

    /** APIから取得したレコードをアプリ用に生成 */
    parseApi(api: A, arg?: unknown): Promise<T | null>;

    /** DBから取得したレコードをアプリ用に成形 */
    fromDbEntity(dbrec: D): T;

    /** アプリ用に成形されたレコードをDB登録用に変換 */
    toDbEntity(teamsrec: T): D;

    /**
     * DBからidを指定して1件のレコードを抽出
     * @param id 
     * @param subentity 
     */
    get(id: string, subentity?: boolean): Promise< T | null>;

    /**
     * DBの全レコードを抽出
     * @param subentity 
     */
    getAll(subentity?: boolean): Promise<T[]>;

    /**
     * 指定レコードをDBに書き込む
     * @param teamsrec 
     * @param subentity 
     */
    put(teamsrec: T, subentity?: boolean): Promise<void>;

    /**
     * 指定レコード(複数)をまとめてDBに書き込む
     * @param teamsrecs 
     * @param subentity 
     */
    putAll(teamsrecs: T[], subentity?: boolean): Promise<void>;

    /**
     * エンティティをapiから取得したデータで同期
     * @param arg 
     */
    sync(arg: ISyncFunctionArg): Promise<boolean>;

}

/** DBアクセス部品のベース */
export abstract class DbAccessorBaseComponent<D extends IDbEntityBase, T extends ITeamsEntityBase, A> implements IDbAccessorClasBase<D, T, A> {
    abstract tableName: string;

    abstract lastSyncedKey: string;

    abstract isDeltaSyncAvailable: boolean;

    abstract parseApi(api: A, arg?: unknown): Promise<T | null>;

    abstract fromDbEntity(dbrec: D): T;

    abstract toDbEntity(teamsrec: T): D;

    /** 従属テーブルをDBから取得するデリゲートメソッド */
    protected abstract getSubEntity: SubEntityFunction;

    /** 従属テーブルをDBに登録するデリゲートメソッド */
    protected abstract putSubEntity: SubEntityFunction;

    /** 従属テーブルを一括DBに一括登録するデリゲートメソッド */
    protected abstract putAllSubEntity: SubEntytyAllFunction;

    /** 従属テーブル操作メソッド用引数を作成 */
    protected abstract createSubEntityArg(parent: T | null, parentDb: D | null): ISubEntityFunctionArg;

    /** apiから全件同期するデリゲートメソッド */
    protected abstract fetchApiAll(arg: ISyncFunctionArg): Promise<T[]>;

    /** apiからデルタ同期するデリゲートメソッド */
    protected abstract fetchApiDelta(arg: ISyncFunctionArg): Promise<T[]>;

    /** サブエンティティを同期するデリゲートメソッド */
    protected abstract syncSubentity(arg: ISyncFunctionArg, parents: T[]): Promise<boolean>;

    /** このエンティティの最終同期日時を取得 */
    getLastSynced(): Date {
        const res = du.parseISO(localStorage.getItem(this.lastSyncedKey) ?? "");
        return res
    }
    
    /**
     * このエンティティを同期した日時を保存
     * @param m タイムスタンプ
     */
    storeLastSynced(m: Date): void {
        localStorage.setItem(this.lastSyncedKey, du.formatISO(m));
    }

    async get(id: string, subentity?: boolean): Promise<T | null> {
        log.info(`▼▼▼ ` + this.tableName + `.get START (id=` + id +`) ▼▼▼`);
        const table: Dexie.Table<D, string> = db.table(this.tableName);
        const result: unknown = await table.get(id);
        const res =result ? this.fromDbEntity(result as D) : null;
        if (res && (subentity?? false)) {           
            this.getSubEntity(this.createSubEntityArg(res, result ? result as D : null));
        }
        log.info(`▲▲▲ ` + this.tableName + `.get END (id=` + id +`) ▲▲▲`);
        return res;        
    }

    async getAll(subentity?: boolean): Promise<T[]> {
        log.info(`▼▼▼ ` + this.tableName + `.getAll START ▼▼▼`);
        const table: Dexie.Table<D, string> = db.table(this.tableName);
        const results = await table.toArray();
        const res: Array<T> = [];
        results.forEach(async dbrec => {
            const teamsrec = this.fromDbEntity(dbrec);
            if (teamsrec && (subentity?? false)) {
                await this.getSubEntity(this.createSubEntityArg(teamsrec, dbrec));
            }
            res.push(teamsrec);
        });
        log.info(`▲▲▲ ` + this.tableName + `.getAll END ▲▲▲`);
        return res;
    }

    async put(teamsrec: T | null, subentity?: boolean): Promise<void> {
        log.info(`▼▼▼ ` + this.tableName + `.put START (id=` + teamsrec?.id +`) ▼▼▼`);
        const table: Dexie.Table<D, string> = db.table(this.tableName);
        const dbrec = teamsrec ? this.toDbEntity(teamsrec) : null;
        if (teamsrec && dbrec) {
            await table.put(dbrec);
            if (subentity?? false) {
                await this.putSubEntity(this.createSubEntityArg(teamsrec, dbrec));
            }
        }       
        log.info(`▲▲▲ ` + this.tableName + `.put END (id=` + teamsrec?.id +`) ▲▲▲`);
    }

    async putAll(teamsrecs: T[], subentity?: boolean): Promise<void> {
        log.info(`▼▼▼ ` + this.tableName + `.putAll START ▼▼▼`);
        const dbrecs: Array<D> = [];
        const args: Array<ISubEntityFunctionArg> = [];
        teamsrecs.forEach(async teamsrec => {
            const dbrec = this.toDbEntity(teamsrec);
            dbrecs.push(dbrec);
            args.push(this.createSubEntityArg(teamsrec,dbrec));
        })
        const table: Dexie.Table<D, string> = db.table(this.tableName);
        table.bulkPut(dbrecs);
        if (subentity?? false) {
            this.putAllSubEntity(args);
        }
        log.info(`▲▲▲ ` + this.tableName + `.putAll END ▲▲▲`);
    }

    @log.traceAsync()
    async sync(arg: ISyncFunctionArg): Promise<boolean> {
        log.info(`▼▼▼ ` + this.tableName + `.sync START ▼▼▼`);
        let res = false;
        
        const lastSync = this.getLastSynced();
        const neverSynced = !du.isValid(lastSync);

        // can only get messages in the last 8 months via delta endpoint, but add margin of error
        const canUseDelta = du.isValid(lastSync) && du.isAfter(lastSync, du.subMonths(du.now(), 7));

        const needFullSync = !this.isDeltaSyncAvailable || neverSynced || !canUseDelta;
        log.info(`full: [${du.isValid(lastSync) ? lastSync.toISOString() : "invalid"}], neverSynced: [${neverSynced}], canUseDelta: [${canUseDelta}]`);

        try {
            const result: Array<ITeamsEntityBase> = [];
            if (needFullSync) {
                log.info(`▼▼▼▼▼ Starting full sync for [${this.tableName}] ▼▼▼▼▼`);
                (await this.fetchApiAll(arg)).map(rec => result.push(rec));               
                log.info(`▲▲▲▲▲ Starting full sync for [${this.tableName}] ▲▲▲▲▲`);
            }
            else {
                log.info(`▼▼▼▼▼ Starting incremental sync for [${this.tableName}] ▼▼▼▼▼`);
                (await this.fetchApiDelta(arg)).map(rec => result.push(rec));
                log.info(`▲▲▲▲▲ Starting incremental sync for [${this.tableName}] ▲▲▲▲▲`);
            }
            res = true;

            if (arg.subentity?? false) {
                res = await this.syncSubentity(arg, result as T[]);
            }

        } catch(error) {
            AI.trackException({
                exception: error,
                properties: {
                    operation: nameof(this.sync),
                }
            });
            res = false;
        } finally {
            log.info(`▲▲▲ ` + this.tableName + `.sync END ▲▲▲`);
        }

        return res;
    }
}