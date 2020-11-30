import { Channel } from '@microsoft/microsoft-graph-types';
import { IFindMsgTeam } from './IFindMsgTeam';
import { IFindMsgChannel } from './IFindMsgChannel';
import { IFindMsgChannelDb } from './IFindMsgChannelDb';
import { assert } from '../utils';
import { dateToNumber, numberToDate, invalidDate } from "../dateUtils";
import { db } from './Database';
import { warn } from '../logger';


export class FindMsgChannel {
    static fromDbEntity(channel: IFindMsgChannelDb): IFindMsgChannel {
        const { lastDeltaUpdate, lastFullMessageSync, ...rest } = channel;
        return {
            ...rest,
            lastDeltaUpdate: numberToDate(lastDeltaUpdate),
            lastFullMessageSync: numberToDate(lastFullMessageSync),
        };
    }


    static toDbEntity(channel: IFindMsgChannel): IFindMsgChannelDb {
        const { lastDeltaUpdate, lastFullMessageSync, ...rest } = channel;
        return {
            ...rest,
            lastDeltaUpdate: dateToNumber(lastDeltaUpdate),
            lastFullMessageSync: dateToNumber(lastFullMessageSync),
        };
    }


    /**
     * Convert Microsoft Graph Channel into IFindMsgChannel
     * Returns null when membershipType is not standard.
     * @param channel
     * @param teamId
     * @param getSyncDateFromDb whether to look up and populate last sync dates from local store
     */
    static async fromChannel(channel: Channel, teamId: string, getSyncDateFromDb = true): Promise<IFindMsgChannel | null> {
        if (channel.membershipType !== "standard") {
            warn(`Ignoring private channel [${channel.displayName}] (id=${channel.id})`);
            return null;
        }

        const id = assert(channel.id);
        let delta = invalidDate();
        let full = invalidDate();

        if (getSyncDateFromDb) {
            const dbChannel = await db.channels.get(id);
            if (dbChannel) {
                delta = numberToDate(dbChannel.lastDeltaUpdate);
                full = numberToDate(dbChannel.lastFullMessageSync);
            }
        }

        return {
            id: assert(channel.id, nameof(channel.id)),
            displayName: assert(channel.displayName, nameof(channel.displayName)),
            description: channel.description || null,
            webUrl: channel.webUrl ?? "",
            lastDeltaUpdate: delta,
            lastFullMessageSync: full,
            teamId,
        };
    }


    /**
     * Get a channel by id
     * @param channelId
     */
    static async get(channelId: string): Promise<IFindMsgChannel | null> {
        const result = await db.channels.get(channelId);
        return result ? FindMsgChannel.fromDbEntity(result) : null;
    }


    /**
     * Upsert channel
     * @param channel
     */
    static async put(channel: IFindMsgChannel): Promise<void> {
        await db.channels.put(FindMsgChannel.toDbEntity(channel));
    }


    /**
     * Get all channels / Get all channels of a team
     * @param teamOrTeamId
     */
    static async getAll(teamOrTeamId?: string | IFindMsgTeam): Promise<IFindMsgChannel[]> {
        let result = db.channels.toCollection();

        if (teamOrTeamId) {
            const tid = typeof teamOrTeamId === "string" ? teamOrTeamId : teamOrTeamId.id;
            result = result.filter(t => t.teamId === tid);
        }

        return (await result.toArray()).map(FindMsgChannel.fromDbEntity);
    }


    /**
     * Return the oldest sync date from all channels
     */
    static async getOldestSync(): Promise<Date> {
        let result = Infinity;
        await db.channels.each(c => result = Math.min(result, Math.max(c.lastDeltaUpdate, c.lastFullMessageSync)));
        if (result === Infinity) result = -1;
        return numberToDate(result);
    }
}
