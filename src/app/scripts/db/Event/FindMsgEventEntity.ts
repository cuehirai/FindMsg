import Dexie from 'dexie';
import { db, idx } from '../Database';
import * as du from "../../dateUtils";
import { BodyType, Event } from "@microsoft/microsoft-graph-types";
import { FindMsgAttendee, IParseAttendeeArg } from "../Attendee/FindMsgAttendeeEntity";
import { DbAccessorBaseComponent, ISubEntityFunctionArg, ISyncFunctionArg, OrderByDirection, SubEntityFunction, SubEntytyAllFunction, SyncError } from "../db-accessor-class-base";
import { IFindMsgEvent } from "./IFindMsgEvent";
import { IFindMsgEventDb } from "./IFindMsgEventDb";
import * as log from '../../logger';
import { collapseConsecutiveChar, collapseWhitespace, sanitize, stripHtml } from '../../purify';
import { IFindMsgAttendee } from '../Attendee/IFindMsgAttendee';
import { getAllPages } from '../../graph/getAllPages';
import { AI } from '../../appInsights';

/**
 * このテーブルがサポートするソートキー
 * ※テーブルに定義したインデックスと一致していること
 */
export enum EventOrder {
    organizer = "organizer",
    start = "start",
    subject = "subject",
}

/**
 * ソートキーからテーブルのインデックス名に変換するためのマップ
 */
const order2IdxMap = {
    organizer: idx.events.organizer$start$subject,
    start: idx.events.start$subject,
    subject: idx.events.subject,
}

/**
 * イベント（スケジュール）エンティティ管理クラス
 */
class EventEntity<D extends IFindMsgEventDb, T extends IFindMsgEvent, A extends Event> extends DbAccessorBaseComponent<D, T, A> {
    
    tableName = "events";

    lastSyncedKey = "FindMsg_events_last_synced";

    // デルタクエリは開始日と終了日の範囲指定でイベントを抽出してくれるだけなので
    // 削除されたイベントを掃除することができない＝＞デルタはサポートしないことにする
    isDeltaSyncAvailable = false;
    
    parseApi = async (api: A): Promise<T | null> => {
        let res: IFindMsgEvent | null = null;
        const {
            id,
            createdDateTime,
            lastModifiedDateTime,
            organizer,
            start,
            end,
            attendees
        } = api;

        if (!id) {
            log.error("Ignoring event without id:", api);
            return null;
        }
        if (!createdDateTime) {
            log.warn("Ignoring event without createdDateTime:", api);
            return null;
        }
        if (!start || !start.dateTime || !end || !end.dateTime) {
            log.warn("Ignoring event without start/end:", api);
            return null;
        }

        const organizerName = (organizer && organizer.emailAddress && organizer.emailAddress.name) ??  null 
        const organizerMail = (organizer && organizer.emailAddress && organizer.emailAddress.address) ?? null;
        
        let type: BodyType;
        let body: string;
        let text: string | null;

        if (!api.body) {
            log.warn("Event has no body:", api);
            type = "text";
            body = "";
            text = api.subject ?? null;
        } else {
            if (api.body.contentType === "text") {
                type = "text";
                body = api.body.content ?? "";
                text = collapseWhitespace((api.subject ?? "") + " " + body).toLowerCase();
            } else if (api.body.contentType === "html") {
                type = "html";
                body = api.body.content ?? "";
                text = collapseWhitespace((api.subject ?? "") + " " + stripHtml(collapseConsecutiveChar(sanitize(body), "_", 3))).toLowerCase();
            } else {
                type = "text";
                body = "";
                text = null;
            }

            // only store text if it is different from the other fields
            if (body === text) text = null;
        }
        const sub: IFindMsgAttendee[] = []
        let i = 0;
        attendees?.forEach(async attendee => {
            const arg: IParseAttendeeArg = {
                id: id + i++,
                eventId: id,
            };
            const rec = await FindMsgAttendee.parseApi(attendee, arg);
            if (rec) {
                sub.push(rec);
            }
        })

        res = {
            id: id,
            created: du.parseISO(createdDateTime),
            modified: lastModifiedDateTime ? du.parseISO(lastModifiedDateTime) : du.invalidDate(),
            organizerName: organizerName,
            organizerMail: organizerMail,
            start: du.parseISO(start.dateTime),
            end: du.parseISO(end.dateTime),
            subject: api.subject ?? "",
            body: body,
            type: type,
            hasAttachments: api.hasAttachments ?? false,
            importance: api.importance ?? "normal",
            sensitivity: api.sensitivity ?? "normal",
            isAllDay: api.isAllDay ?? false,
            isCancelled: api.isCancelled ?? false,
            webLink: api.webLink ?? "",
            text: text,
            attendees: sub,
        }
        
        return res as T;
    }
    
    fromDbEntity(dbrec: D): T {
        const {
            created,
            modified,
            start,
            end,
            ...rest
        } = dbrec as IFindMsgEventDb;

        const res: IFindMsgEvent = {
            created: du.numberToDate(created),
            modified: du.numberToDate(modified),
            start: du.numberToDate(start),
            end: du.numberToDate(end),
            attendees: [],
            ...rest
        };

        this.getSubEntity({parent: res, parentDb:null});

        return res as T;
    }
    
    toDbEntity(teamsrec: T): D {
        const {
            created,
            modified,
            start,
            end,
            ...rest
        } = teamsrec as IFindMsgEvent;

        const res: IFindMsgEventDb = {
            created: du.dateToNumber(created),
            modified: du.dateToNumber(modified),
            start: du.dateToNumber(start),
            end: du.dateToNumber(end),
            ...rest
        };

        return res as D;
    }
    
    protected getSubEntity: SubEntityFunction = async (arg: ISubEntityFunctionArg): Promise<void> => {
        const {parent} = arg;
        if (parent) {
            const attendees = await db.attendees.where(idx.attendees.$eventId$id).between([parent.id || Dexie.minKey], [parent.id || Dexie.maxKey], true, true).toArray();
            (parent as IFindMsgEvent).attendees = attendees;
        }
    };

    protected putSubEntity: SubEntityFunction = async (arg: ISubEntityFunctionArg): Promise<void> => {
        const {parent} = arg;
        const recs: Array<IFindMsgAttendee> = [];
        (parent as IFindMsgEvent).attendees.forEach(rec => recs.push(rec));
        await FindMsgAttendee.putAll(recs);
    };

    protected putAllSubEntity: SubEntytyAllFunction = async (args: ISubEntityFunctionArg[]): Promise<void> => {
        const recs: Array<IFindMsgAttendee> = [];
        args.forEach(arg => {
            const {parent} = arg;
            (parent as IFindMsgEvent).attendees.forEach(rec => recs.push(rec));
        })
        if (recs.length > 0) {
            FindMsgAttendee.putAll(recs);
        }
    };
    
    protected createSubEntityArg(parent: T | null, parentDb: D | null): ISubEntityFunctionArg {
        return {parent, parentDb};
    }

    protected async fetchApiAll(arg: ISyncFunctionArg): Promise<T[]> {
        const res: Array<IFindMsgEvent> = [];
        const client = arg.client;
        try {
            arg.progress(arg.translate.common.syncEntity(arg.translate.entities.events));
            const existingIds = await db.events.toCollection().primaryKeys();
            log.info(`existing eventIds: [${existingIds.join("], [")}]`);

            log.info(`calling API [/me/calendar/events]`);

            const response = await client.api('/me/calendar/events')
            .get();
            const fetchedAll = await getAllPages<Event>(client, response);

            if (fetchedAll === null) {
                if (existingIds.length === 0) {
                    throw new SyncError("Could not sync events");
                }
            } else {
                let ommitted = 0;
                const fetched: Array<Event> = [];
                fetchedAll.forEach(rec => {
                    let apply = true;
                    if (rec.seriesMasterId && rec.seriesMasterId != rec.id) {
                        apply = false;
                    }
                    log.info(`id:[${rec.id?? "null"}], subject:[${rec.subject?? "null"}], isCancelled:[${rec.isCancelled?? "null"}], organizer:[${rec.organizer? rec.organizer.emailAddress? rec.organizer.emailAddress.name?? "null": "null": "null"}], apply:[${apply}]`);
                    if (apply) {
                        fetched.push(rec);
                    } else {
                        ommitted += 1;
                    }
                })

                log.info(`API returned [${fetched.length}] events (${ommitted} records ommitted due to being recurrence data)`);
                await db.transaction("rw", db.events, db.attendees, async () => {
                    const events = await Promise.all(fetched.map(event => this.parseApi(event as A)));
                    let count = 0;
                    await Promise.all(events.map(t => {
                        arg.checkCancel();
                        if (t) {
                            this.put(t);
                            res.push(t);
                            count += 1;
                            arg.progress(arg.translate.common.syncEntityWithCount(arg.translate.entities.events, count));
                        }
                    }));
    
                    // DBにあってapiの戻りに存在しないイベントIDはDBから削除する.
                    const deletedIds = existingIds.filter(exist => !events.some(t => (t as IFindMsgEvent).id === exist));
                    log.info(`deleted eventIds: [${deletedIds.join("], [")}]`);
                    deletedIds.forEach( async eId => {
                        // 削除されるイベントに属する参加者レコードを先に削除しておく
                        const deletedAttendee = await db.attendees.where('eventId').equals(eId).primaryKeys();
                        log.info(`★★★ deleting Attendees due to event deletion; keys:[${deletedAttendee.join("], [")}]`);
                        await db.attendees.bulkDelete(deletedAttendee)
                        log.info(`★★★ Attendees deletion completed`);
                    })
                    log.info(`★★★ deleting events; keys:[${deletedIds.join("], [")}]`);
                    await db.events.bulkDelete(deletedIds);
                    log.info(`★★★ Events deletion completed`);
                });
    
                await this.storeLastSynced(du.now(), true);
    
            }
        } catch (error) {
            AI.trackException({
                exception: error,
                properties: {
                    operation: nameof(this.fetchApiAll),
                }
            });
        }

        return res as T[];
    }

    protected async fetchApiDelta(arg: ISyncFunctionArg): Promise<T[]>{
        const res: Array<IFindMsgEvent> = [];
        const client = arg.client;
        const last = await this.getLastSynced();

        if (!du.isValid(last)) {
            throw new Error("last delta sync invalid");
        }
        if (du.isBefore(last, du.subDays(du.subMonths(du.now(), 7), 1))) {
            throw new Error("last delta sync too old");
        }

        try {
            arg.progress(arg.translate.common.syncEntity(arg.translate.entities.events));

            const cutOffTime = du.subMinutes(last, 5);
            const endtime = du.now();
            endtime.setFullYear(endtime.getFullYear() + 1);

            // const delta = `/me/calendarView/delta?startdatetime=${cutOffTime.toISOString()}&enddatetime=${now.toISOString()}`;
            const delta = `/me/calendarView/delta?startdatetime=${cutOffTime.toISOString()}&enddatetime=${endtime.toISOString()}`;
            log.info(`calling API [${delta}]`);
            const response = await client.api(delta)
                .get();

            const fetched = await getAllPages<Event>(client, response);
            fetched.forEach(rec => {
                log.info(`id:[${rec.id?? "null"}], subject:[${rec.subject?? "null"}], isCancelled:[${rec.isCancelled?? "null"}], organizer:[${rec.organizer? rec.organizer.emailAddress? rec.organizer.emailAddress.name?? "null": "null": "null"}]`);
            })

            log.info(`API returned [${fetched.length}] events`);
            await db.transaction("rw", db.events, db.attendees, async () => {
                const events = await Promise.all(fetched.map(event => this.parseApi(event as A)));
                let count = 0;
                await Promise.all(events.map(t => {
                    arg.checkCancel();
                    if (t) {
                        this.put(t);
                        res.push(t);
                        count += 1;
                        arg.progress(arg.translate.common.syncEntityWithCount(arg.translate.entities.events, count));
                    }
                }));
            });

            await this.storeLastSynced(du.now(), true);
        } catch (error) {
            AI.trackException({
                exception: error,
                properties: {
                    operation: nameof(this.fetchApiDelta),
                }
            });
        }
        return res as T[];
    }

    protected async syncSubentity(arg: ISyncFunctionArg, parents: T[]): Promise<boolean>{
        log.info(`▼▼▼▼▼ syncSubentity START for Event id: [${arg.parent? arg.parent.id : "null"}] ▼▼▼▼▼`);
        let res = true;
        parents.forEach(async rec => {
            const subarg: ISyncFunctionArg = {
                client: arg.client,
                checkCancel: arg.checkCancel,
                progress: arg.progress,
                subentity: false,
                parent: rec,
                translate: arg.translate,
            };
            res = res && await FindMsgAttendee.sync(subarg);
        })
        log.info(`▲▲▲▲▲ syncSubentity END for Event id: [${arg.parent? arg.parent.id : "null"}] ▲▲▲▲▲`);
        return res;
    }

    /**
     * Create a filter function that searches message text if available and subject and body if not.
     * Note: filter is always case insensitive
     * @param searchTerm the term to filter for
     * @returns a function that takes a messages and returns whether the message contains the search term
     */
    createFilter(searchTerm: string): (m: IFindMsgEvent | IFindMsgEventDb) => boolean {
        const t = searchTerm.toLowerCase();
        return ({ subject, body, text }) => typeof text === "string" ? text.includes(t) : subject?.toLowerCase().includes(t) || body.toLowerCase().includes(t);
    }

    async fetch(order: EventOrder, dir: OrderByDirection, offset = 0, limit = 0, filter = "", from: Date, to: Date, organizer: Set<string>): Promise<[IFindMsgEvent[], boolean]> {
        log.info(`▼▼▼ fetch START ▼▼▼`);
        const index = order2IdxMap[order];

        const collection = db.events.orderBy(index);

        if (dir === OrderByDirection.descending) collection.reverse();
        if (offset > 0) collection.offset(offset);
        if (limit > 0) collection.limit(limit + 1);
        if (filter.trim()) collection.filter(this.createFilter(filter));

        const fromValid = du.isValid(from);
        const toValid = du.isValid(to);
        if (fromValid || toValid) {
            if (fromValid && toValid && du.isAfter(from, to)) [from, to] = [du.startOfDay(to), du.endOfDay(from)];
            const fromN = from.valueOf();
            const toN = to.valueOf();

            if (fromValid && toValid) {
                collection.filter(m => m.start >= fromN && m.start <= toN);
            } else if (fromValid) {
                collection.filter(m => m.start >= fromN);
            } else {
                collection.filter(m => m.start <= toN);
            }
        }
    
        if (organizer.size > 0) {
            collection.filter(rec => organizer.has(rec.organizerName?? ""));
        }

        const result = await collection.toArray();

        let hasMore = false;

        log.info(`▲▲▲ fetch END ▲▲▲`);
        if (result.length > 0) {
            if (result.length > limit) {
                result.pop();
                hasMore = true;
            }
            return [result.map(r => this.fromDbEntity(r as D) as IFindMsgEvent), hasMore];
        }
        return [[], hasMore];
    }

    /** 主催者をすべて取得 */
    async getOrganizers(): Promise<string[]> {
        const all = await db.events.orderBy(idx.events.organizer$start$subject).toArray();
        let orgbk = "";
        const res: Array<string> = [];
        all.forEach(rec => {
            if (rec.organizerName && rec.organizerName != orgbk) {
                orgbk = rec.organizerName;
                res.push(orgbk);
            }
        })
        return res;
    }

}

export const FindMsgEvent = new EventEntity<IFindMsgEventDb, IFindMsgEvent, Event>();
