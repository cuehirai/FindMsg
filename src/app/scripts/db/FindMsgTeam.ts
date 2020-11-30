import { Team } from '@microsoft/microsoft-graph-types';
import { IFindMsgTeam } from './IFindMsgTeam';
import { IFindMsgTeamDb } from './IFindMsgTeamDb';
import { IFindMsgChannel } from './IFindMsgChannel';
import { assert } from '../utils';
import { dateToNumber, numberToDate } from "../dateUtils";
import { db } from './Database';
import { FindMsgChannel } from "./FindMsgChannel";


export class FindMsgTeam {
    /**
     * Convert from storage entity to
     * @param team
     */
    static fromDbEntity(team: IFindMsgTeamDb): IFindMsgTeam {
        const { lastChannelListSync, ...rest } = team;
        return {
            ...rest,
            lastChannelListSync: numberToDate(lastChannelListSync),
        };
    }


    /**
     * Convert to storage entity
     * @param team
     */
    static toDbEntity(team: IFindMsgTeam): IFindMsgTeamDb {
        const { lastChannelListSync, ...rest } = team;
        return {
            ...rest,
            lastChannelListSync: dateToNumber(lastChannelListSync),
        };
    }


    /**
     * Convert from Microsoft Graph API Team
     * @param team
     * @param getSyncDateFromDb
     */
    static async fromTeam(team: Team, getSyncDateFromDb = true): Promise<IFindMsgTeam> {

        const id = assert(team.id, nameof(team.id));
        let syncDate = -1;

        if (getSyncDateFromDb) {
            const dbTeam = await db.teams.get(id);
            if (dbTeam) {
                syncDate = dbTeam.lastChannelListSync;
            }
        }

        return {
            id,
            displayName: assert(team.displayName, nameof(team.displayName)),
            description: team.description ?? null,
            webUrl: team.webUrl ?? null,
            lastChannelListSync: numberToDate(syncDate),
        };
    }


    /**
     * Get team from DB by IDs
     * @param id
     */
    static async get(id: string): Promise<IFindMsgTeam | null> {
        const result = await db.teams.get(id);
        return result ? FindMsgTeam.fromDbEntity(result) : null;
    }



    /**
     * Get all teams from DB
     */
    static async getAll(): Promise<IFindMsgTeam[]> {
        return (await db.teams.toArray()).map(FindMsgTeam.fromDbEntity);
    }



    /**
     * Store team in DB (overwrite existing)
     * @param team
     */
    static async put(team: IFindMsgTeam): Promise<void> {
        await db.teams.put(FindMsgTeam.toDbEntity(team));
    }


    /**
     * Get all channels of team from DB
     * @param teamOrId
     */
    static async getChannels(teamOrId: IFindMsgTeam | string): Promise<IFindMsgChannel[]> {
        const key = typeof teamOrId === "string" ? teamOrId : teamOrId.id;
        const result = await db.channels.where('teamId').equals(key).toArray();
        return result.map(FindMsgChannel.fromDbEntity);
    }


    /**
     * Replace all channels of the team in the database
     * Delete channels (and messages) not in the channels parameter
     * @param teamOrId
     * @param channels
     */
    static putChannels(teamOrId: IFindMsgTeam | string, channels: IFindMsgChannel[]): Promise<void> {
        const teamId = typeof teamOrId === "string" ? teamOrId : teamOrId.id;

        if (channels.some(ch => ch.teamId !== teamId)) {
            throw new Error("team id doesn't match channel's team id.");
        }

        return db.transaction('rw', db.channels, db.channelMessages, async () => {
            // save channels
            const ids = await Promise.all(channels.map(ch => db.channels.put(FindMsgChannel.toDbEntity(ch))));

            // get ids of deleted channels
            const deletedChannelIds = await db.channels.where('teamId').equals(teamId).and(ch => ids.indexOf(ch.id) === -1).primaryKeys();

            // delete those channels
            await db.channels.bulkDelete(deletedChannelIds);

            // delete associated messages
            await Promise.all(deletedChannelIds.map(cid => db.channelMessages.where('channelId').equals(cid).delete()));
        });
    }
}
