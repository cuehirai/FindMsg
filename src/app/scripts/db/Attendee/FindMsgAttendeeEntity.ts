import { db, idx } from '../Database';
import * as du from "../../dateUtils";
import { Attendee, ResponseType } from "@microsoft/microsoft-graph-types";
import { DbAccessorBaseComponent, ISubEntityFunctionArg, ISyncFunctionArg, SubEntityFunction, SubEntytyAllFunction } from "../db-accessor-class-base";
import { IFindMsgAttendeeDb } from "./IFindMsgAttendeeDb";
import { IFindMsgAttendee } from "./IFindMsgAttendee";
import { IFindMsgEvent } from "../Event/IFindMsgEvent";
import * as log from '../../logger';
import Dexie from 'dexie';

/**
 * 参加者のAPIデータをアプリ用に変換する際の引数に含めるインターフェース
 * ※このテーブルはEventの属性として含まれるデータをもとにするので参加者のIDやEventのIDはアプリで与える必要があります。
 * このためparseApiの引数としてこれらの情報を渡します。
 */
export interface IParseAttendeeArg {
    /** 参加者のID */
    id: string;
    /** 親であるEventのID */
    eventId: string;
}

/** イベントの参加者エンティティ管理クラス */
class AttendeeEntity<D extends IFindMsgAttendeeDb, T extends IFindMsgAttendee, A extends Attendee> extends DbAccessorBaseComponent<D, T, A> {
    tableName = "attendees";

    lastSyncedKey = "FindMsg_attendees_last_synced";

    isDeltaSyncAvailable = false;
    
    parseApi = async (api: A, arg: IParseAttendeeArg): Promise<T | null>  => {
        const {id, eventId} = (arg as IParseAttendeeArg);
        const organizer:ResponseType = "organizer";
        let isOrganizer = false;
        if (api.status && api.status.response && api.status.response === organizer) {
            isOrganizer = true;
        }
        const res: IFindMsgAttendee = {
            id: id,
            eventId: eventId,
            isOrganizer: isOrganizer,
            name: api.emailAddress? api.emailAddress.name ?? "" : "",
            mail: api.emailAddress? api.emailAddress.address ?? "" : "",
            type: api.type ?? "required",
            status: api.status? api.status.response ?? "none" : "none",
        };
        return res as T;
    }

    fromDbEntity(dbrec: D): T {
        const res: IFindMsgAttendee = dbrec;
        return res as T;
    }
    toDbEntity(teamsrec: T): D {
        const res: IFindMsgAttendeeDb = teamsrec;
        return res as D;
    }
    protected getSubEntity: SubEntityFunction = async (): Promise<void> => {/* 実装なし */};

    protected putSubEntity: SubEntityFunction = async (): Promise<void> => {/* 実装なし */};

    protected putAllSubEntity: SubEntytyAllFunction = async (): Promise<void> => {/* 実装なし */};

    protected createSubEntityArg(parent: T | null, parentDb: D | null): ISubEntityFunctionArg {
        return {parent, parentDb};
    }

    protected async fetchApiAll(arg: ISyncFunctionArg): Promise<T[]> {
        const parent = arg.parent;
        const res: Array<IFindMsgAttendee> = (parent && (parent as IFindMsgEvent).attendees) ?? [];
        if (parent) {
            await db.transaction("rw", db.events, db.attendees, async () => {
                const delcount = db.attendees.where(idx.attendees.$eventId$id).between([parent.id || Dexie.minKey], [parent.id || Dexie.maxKey], true, true).delete;
                log.info(`deleted [${delcount}] existing attendees for event [${parent.id}]`);
                res.forEach(rec => this.put(rec as T));
                log.info(`registered [${res.length}] attendees for event [${parent.id}]`)
            });
            await this.storeLastSynced(du.now());
        }
        return res as T[];
    }

    protected async fetchApiDelta(arg: ISyncFunctionArg): Promise<T[]> {
        const res: Array<IFindMsgAttendee> = (arg.parent && (arg.parent as IFindMsgEvent).attendees) ?? [];
        await this.storeLastSynced(du.now());
        return res as T[];
    }

    protected async syncSubentity(): Promise<boolean> {
        return true;
    }
}

export const FindMsgAttendee = new AttendeeEntity<IFindMsgAttendeeDb, IFindMsgAttendee, Attendee>();
