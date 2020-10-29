import Dexie from 'dexie';
import { db, idx } from '../Database';
import * as du from "../../dateUtils";
import { BodyType, Event } from "@microsoft/microsoft-graph-types-beta";
import { FindMsgAttendee, IParseAttendeeArg } from "../Attendee/FindMsgAttendeeEntity";
import { IFindMsgAttendeeDb } from "../Attendee/IFindMsgAttendeeDb";
import { DbAccessorBaseComponent, ISubEntityFunctionArg, ISyncFunctionArg, OrderByDirection, SubEntityFunction, SubEntytyAllFunction, SyncError } from "../db-accessor-class-base";
import { IFindMsgEvent } from "./IFindMsgEvent";
import { IFindMsgEventDb } from "./IFindMsgEventDb";
import * as log from '../../logger';
import { collapseWhitespace, sanitize, stripHtml } from '../../purify';
import { IFindMsgAttendee } from '../Attendee/IFindMsgAttendee';
import { getAllPages } from '../../graph/getAllPages';
import { AI } from '../../appInsights';

export enum EventOrder {
    organizer = "organizer",
    start = "start",
    subject = "subject",
}

const order2IdxMap = {
    organizer: idx.events.organizer$start$subject,
    start: idx.events.start$subject,
    subject: idx.events.subject,
}

class EventEntity<D extends IFindMsgEventDb, T extends IFindMsgEvent, A extends Event> extends DbAccessorBaseComponent<D, T, A> {
    
    tableName = "events";

    lastSyncedKey = "FindMsg_events_last_synced";

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
            log.error("Ignoring event without createdDateTime:", api);
            return null;
        }
        if (!start || !start.dateTime || !end || !end.dateTime) {
            log.error("Ignoring event without start/end:", api);
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
                body = sanitize(api.body.content ?? "");
                text = collapseWhitespace((api.subject ?? "") + " " + stripHtml(body)).toLowerCase();
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
        const recs = (parent as IFindMsgEvent).attendees.map(e => FindMsgAttendee.toDbEntity(e));
        await FindMsgAttendee.putAll(recs);
    };

    protected putAllSubEntity: SubEntytyAllFunction = async (args: ISubEntityFunctionArg[]): Promise<void> => {
        const recs: Array<IFindMsgAttendeeDb> = [];
        args.forEach(arg => {
            const {parent} = arg;
            (parent as IFindMsgEvent).attendees.map(rec => FindMsgAttendee.toDbEntity(rec)).forEach(dbrec => {recs.push(dbrec)});
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

            log.info(`calling API [/me/calendar/events]`);

            const response = await client.api('/me/calendar/events')
            .version('beta')
            .get();
            const fetched = await getAllPages<Event>(client, response);

            if (fetched === null) {
                if (existingIds.length === 0) {
                    throw new SyncError("Could not sync events");
                }
            } else {
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
    
                    // delete teams that where not in the response from the local database.
                    const deletedIds = existingIds.filter(exist => !events.some(t => (t as IFindMsgEvent).id === exist));
                    // delete the channels
                    await Promise.all(deletedIds.map(async dtId => {
                        const deletedAttendee = await db.attendees.where('eventId').equals(dtId).primaryKeys();
                        await db.attendees.bulkDelete(deletedAttendee)
                    }));
                    await db.events.bulkDelete(deletedIds);
                });
    
                this.storeLastSynced(du.now());
    
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

    protected async fetchApiDelta(): Promise<T[]>{
        const res: Array<IFindMsgEvent> = [];
        return res as T[];
    }

    protected async syncSubentity(arg: ISyncFunctionArg, parents: T[]): Promise<boolean>{
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

    async fetch(order: EventOrder, dir: OrderByDirection, offset = 0, limit = 0, filter = ""): Promise<[IFindMsgEvent[], boolean]> {
        const index = order2IdxMap[order];

        const collection = db.events.orderBy(index);

        if (dir === OrderByDirection.descending) collection.reverse();
        if (offset > 0) collection.offset(offset);
        if (limit > 0) collection.limit(limit + 1);
        if (filter.trim()) collection.filter(this.createFilter(filter));

        const result = await collection.toArray();

        let hasMore = false;

        if (result.length > 0) {
            if (result.length > limit) {
                result.pop();
                hasMore = true;
            }

            return [result.map(r => this.fromDbEntity(r as D) as IFindMsgEvent), hasMore];
        }

        return [[], hasMore];

    }

}

export const FindMsgEvent = new EventEntity<IFindMsgEventDb, IFindMsgEvent, Event>();
