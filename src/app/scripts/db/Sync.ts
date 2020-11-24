import Dexie from 'dexie';
import { Entity, ChatMessage, Team, Channel, Chat, ConversationMember } from '@microsoft/microsoft-graph-types-beta';
import { Client, ResponseType } from '@microsoft/microsoft-graph-client';
import { IFindMsgTeam } from './IFindMsgTeam';
import { IFindMsgChannel } from './IFindMsgChannel';
import { IFindMsgChannelMessage } from './IFindMsgChannelMessage';
import * as log from '../logger';
import { throwFn, nop, progressFn, filterNull, OperationCancelled, assert, hashBlob, isGraphHostedContentUrl } from '../utils';
import { getAllPages } from "../graph/getAllPages";
import { FullPageIterator } from '../graph/FullPageIterator';
import { FindMsgTeam } from './FindMsgTeam';
import { FindMsgChannel } from './FindMsgChannel';
import { FindMsgChannelMessage } from './FindMsgChannelMessage';
import { db, idx } from './Database';
import * as du from "../dateUtils";
import { AI } from '../appInsights';

import { IFindMsgChat } from './IFindMsgChat';
import { FindMsgChat } from './FindMsgChat';
import { FindMsgChatMessage } from './FindMsgChatMessage';
import { FindMsgChatMember } from './FindMsgChatMember';
import { IFindMsgChatMessage } from './IFindMsgChatMessage';
import { GraphImage } from '../graphImage';
import { IFindMsgImageDb } from './IFindMsgImageDb';

/*
Sync design notes:
==================

We are killed every time the user navigates, so need to plan for interruption.
The flow is designed to accept some inconsitency / stale data.
This in turn allows us to break up the sync into smaller parts
that each are reasonably fast or can be resumed.

Team list sync
--------------
User is assumed to have only a small number of teams --> reasonable fast (worst case 1000).

- check teams sync time
- if older than X then sync teams (add new teams, delete removed teams, delete channels and messages of removed teams)
- set teams sync time to now

Channel list sync (for each Team, assumed to be 'fast', at most 200 channels per team)
-----------------
Done for each team separately.
Team is assumed to have only a few channels --> fast (worst case 200).

- check channel list sync time
- if older than Y then sync channel list (add new channels, delete removed channels, delete messages of removed channels)
- set sync time to now

Top level message full sync
---------------------------
This is the most critical operation, because it might take a very long time
and the API does not seem to allow reliable resume when interrupted.
A channel can have a lot of messages and one message can be up to 28kb big.

- if messages were never synced
- iterate over ALL message (potentially A LOT) and dump them into the database (record answers sync time as null)
- set the sync timestamp to now
- set incremental sync timestamp to now

Top level message incremental sync
----------------------------------
Somewhat critical. May take a long time if there are a lot of messages since the last sync.

This has some crazy semantics:
- The delta endpoint does NOT include new replies to top level messages.
- Messges with new replies will appear in the response even if their createdTime is before the filter time (lastModified is still null)
  --> can detect new/changed replies with this
- Last modified time is only there for actually edited messages
  --> no info about new/changed replies --> must sync
- deletedTime is there for deleted messages (body is blank)
  --> no info new/changed replies --> must sync

Operation:
- if incremental sync timestamp older than Z
- fetch messages/delta with delta timestamp set to last incremental sync timestamp (minus 5 minutes or so to accound for server time offset)
- CUD received messages in to database
- sync replies for received messages according to above semantics
- set incremental sync timestamp to now

Message answer sync
-------------------
May take a long time as a whole.
Done for each answer separately so a lot of requests are needed.
Each single message is fast, assuming it does not have a large number of replies.
Can easily be resumed with top level message granularity.

- for each top level message (parent ID == null)
- fetch message replies and dump into DB
- record answer fetch timestamp in parent

On Cancellation
---------------
Sync operation allows for cooperative cancellation using a token.
*/

/*
On Sync Resume
--------------
There are some stages in the sync operation that can not be resumed,
because they depend on opaque state of the Microsoft Graph API.
This state is encoded into the "@odata.nextLink" returned from the API.

Since these links are valid for an undocumented amount of time, in theory
a resume could be attempted in certain circumstances.

Could try in an unload handler to cancel a running sync, save a @nextLink if any.
On restore, could read the state and try to keep going from there.

This would require a rewrite of the sync logic, because currently state is not exposed.
*/


const teamSyncKey = "FindMsg_teams_last_synced";
// const getTeamsLastSynced = (): Date => du.parseISO(localStorage.getItem(teamSyncKey) ?? "");
// const setTeamsLastSynced = (ts: Date) => localStorage.setItem(teamSyncKey, du.formatISO(ts));
const getTeamsLastSynced = async (): Promise<Date> => await db.getLastSync(teamSyncKey);
const setTeamsLastSynced = async (ts: Date) => await db.storeLastSync(teamSyncKey, ts);

const topLevelMessagesSyncKey = "FindMsg_toplevel_messages_last_synced";
// export const getTopLevelMessagesLastSynced = (): Date => du.parseISO(localStorage.getItem(topLevelMessagesSyncKey) ?? "");
// const setTopLevelMessagesLastSynced = (ts: Date) => localStorage.setItem(topLevelMessagesSyncKey, du.formatISO(ts));
export const getTopLevelMessagesLastSynced = async (): Promise<Date> => await db.getLastSync(topLevelMessagesSyncKey);
const setTopLevelMessagesLastSynced = async (ts: Date, doExport: boolean) => await db.storeLastSync(topLevelMessagesSyncKey, ts, doExport);

// chat部
const chatSyncKey = "FindMsg_chats_last_synced";
// const getChatsLastSynced = (): Date => du.parseISO(localStorage.getItem(chatSyncKey) ?? "");
// const setChatsLastSynced = (ts: Date) => localStorage.setItem(chatSyncKey, du.formatISO(ts));
const getChatsLastSynced = async (): Promise<Date> => await db.getLastSync(chatSyncKey);
const setChatsLastSynced = async (ts: Date) => await db.storeLastSync(topLevelMessagesSyncKey, ts);

export interface ISyncProgressTranslation {
    teamList: string;
    channelList: (teamName: string) => string;
    topLevelMessages: (channelName: string, numSynced: number) => string;
    replies: (channelName: string, numSynced: number) => string;
    syncProblem: string;
    chatList: string;
    chatMessages: (chatId: string, numSynced: number) => string;
}


class SyncError extends Error {
    private constructor(message: string) {
        super(message);
    }

    public static readonly TeamList = () => { throw new SyncError("Could not sync teamlist") }

    public static readonly ChatList = () => { throw new SyncError("Could not sync chatlist") }
}


export class Sync {
    /**
     * Automatically sync teams, channels and messages
     * @param client MsGraph client to use for request
     * @param includeReplies Whether to sync channel message replies
     * @param checkCancel a function that throws OperationCancelledError when cancellation is requested
     * @param progress a function to accept sync progress reports
     */
    @log.traceAsync(true)
    static async autoSyncAll(client: Client, includeReplies: boolean, checkCancel: throwFn = nop, progress: progressFn = nop, { teamList, channelList, topLevelMessages, replies }: ISyncProgressTranslation, doExport: boolean): Promise<boolean> {
        let success = true;

        if (await this.isTeamListStale()) {
            progress(teamList);
            success = success && await Sync.syncTeamList(client);
        }

        const teams = await FindMsgTeam.getAll();
        for (const team of teams.filter(Sync.isChannelListStale)) {
            checkCancel();
            progress(channelList(team.displayName));
            success = success && await Sync.syncChannelList(client, team);
        }

        const channels = await FindMsgChannel.getAll();
        for (const channel of channels.filter(Sync.shouldSyncMessages)) {
            checkCancel();
            let synced = 0;
            success = success && await Sync.syncTopLevelMessages(client, channel, checkCancel, n => progress(topLevelMessages(channel.displayName, synced += n)));
        }
        await setTopLevelMessagesLastSynced(await FindMsgChannel.getOldestSync(), doExport);

        if (includeReplies) {
            const batchSize = 20;
            let total = 0;
            for (const channel of channels) {
                try {
                    progress(`Syncing message replies of channel [${channel.displayName}]... `);

                    const reportCount = (n: number) => progress(replies(channel.displayName, total += n));
                    let synced = 0;
                    do {
                        checkCancel();
                        synced = await Sync.syncMessageRepliesBatch(client, channel, batchSize, checkCancel, reportCount);
                    } while (synced === batchSize);
                } catch (error) {
                    AI.trackException({
                        exception: error,
                        properties: {
                            operation: nameof(Sync.syncMessageRepliesBatch),
                            channelId: channel.id,
                            channelName: channel.displayName,
                        }
                    })
                    success = false;
                }
            }
        }

        return success;
    }


    /**
     * Sync top-level messages of a channel, syncing teams and channels on the way if needed.
     * This is the sync function for the channel topics tab
     * @param client
     * @param teamId
     * @param channelId
     * @param checkCancel
     * @param progress
     */
    @log.traceAsync(true)
    static async channelTopLevelMessages(client: Client, teamId: string, channelId: string, checkCancel: throwFn = nop, progress: progressFn = nop, { teamList, channelList, topLevelMessages }: ISyncProgressTranslation): Promise<boolean> {
        let success = true;

        let team = await FindMsgTeam.get(teamId);
        if (!team || await this.isTeamListStale()) {
            progress(teamList);
            success = success && await Sync.syncTeamList(client);
            team = await FindMsgTeam.get(teamId);
            if (!team) throw new Error(`Team not found: [${channelId}]`);
        }

        let channel = await FindMsgChannel.get(channelId);
        if (!channel || Sync.isChannelListStale(team)) {
            checkCancel();
            progress(channelList(team.displayName));
            success = success && await Sync.syncChannelList(client, team);
            channel = await FindMsgChannel.get(channelId);
            if (!channel) throw new Error(`Channel not found: [${channelId}]`);
        }

        if (Sync.shouldSyncMessages(channel)) {
            checkCancel();
            let accumulator = 0;
            const channelName = channel.displayName;
            const report = (n: number) => progress(topLevelMessages(channelName, accumulator += n));
            success = success && await Sync.syncTopLevelMessages(client, channel, checkCancel, report);
        }

        return success;
    }


    /**
     * Fetch all teams the user is part of
     * The user can be part of 1000 teams at most, so this should be reasonably fast.
     * @param client
     */
    @log.traceAsync(true)
    private static async syncTeamList(client: Client): Promise<boolean> {
        const existingTeamIds = await db.teams.toCollection().primaryKeys();
        const fetchedTeams = await Sync.fetchTeamList(client);

        if (fetchedTeams === null) {
            // Abort if we have no Teams at all
            return existingTeamIds.length === 0 ? SyncError.TeamList() : false;
        } else {
            await db.transaction("rw", db.teams, db.channels, db.channelMessages, async () => {
                const teams = await Promise.all(fetchedTeams.map(ft => FindMsgTeam.fromTeam(ft, true)));
                await Promise.all(teams.map(t => FindMsgTeam.put(t)));

                // delete teams that where not in the response from the local database.
                const deletedTeamIds = existingTeamIds.filter(etId => !teams.some(t => t.id === etId));
                // delete the channels
                await Promise.all(deletedTeamIds.map(dtId => FindMsgTeam.putChannels(dtId, [])));
                await db.teams.bulkDelete(deletedTeamIds);
            });

            await setTeamsLastSynced(du.now());
            return true;
        }
    }


    private static async fetchTeamList(client: Client): Promise<Team[] | null> {
        try {
            const response = await client.api('/me/joinedTeams').version('beta').get();
            const fetchedTeams = await getAllPages<Team>(client, response);

            log.info(`API returned [${fetchedTeams.length}] teams`);

            return fetchedTeams;
        } catch (error) {
            AI.trackException({
                exception: error,
                properties: {
                    operation: nameof(Sync.fetchTeamList),
                }
            });
            return null;
        }
    }


    /**
     * Fetch all channels of the specified team.
     * Writes the supplied team to the DB with fetch timestamp.
     * A team can have a maximum of 200 channels, so this is assumed to be reasonably fast.
     * @param client
     * @param team
     */
    @log.traceAsync()
    private static async syncChannelList(client: Client, team: IFindMsgTeam): Promise<boolean> {
        const fetchedChannels = await Sync.fetchChannelList(client, team);

        if (fetchedChannels === null) {
            return false;
        } else {
            log.info(`API returned ${fetchedChannels.length} channels for team [${team.displayName}]`);
            const timestamp = du.now();

            await db.transaction("rw", db.teams, db.channels, db.channelMessages, async () => {
                const tchannels = await Promise.all(fetchedChannels.map(ch => FindMsgChannel.fromChannel(ch, team.id, true)));
                await FindMsgTeam.putChannels(team, tchannels.filter(filterNull));
                team.lastChannelListSync = timestamp;
                await FindMsgTeam.put(team);
            });

            return true;
        }
    }


    private static async fetchChannelList(client: Client, team: IFindMsgTeam): Promise<Channel[] | null> {
        try {
            const response = await client.api(`/teams/${team.id}/channels`).version('beta').get();
            return await getAllPages<Channel>(client, response);
        } catch (error) {
            AI.trackException({
                exception: error,
                properties: {
                    operation: nameof(Sync.fetchChannelList),
                    teamId: team.id,
                    teamName: team.displayName,
                }
            });
            return null;
        }
    }


    /**
     * Sync channel top level messages.
     * Uses full sync or delta sync based on last sync date
     * @param client
     * @param channel
     */
    @log.traceAsync()
    private static async syncTopLevelMessages(client: Client, channel: IFindMsgChannel, checkCancel: throwFn, report: (n: number) => void): Promise<boolean> {
        const { lastFullMessageSync: full, lastDeltaUpdate: delta } = channel;
        const neverSynced = !du.isValid(full);
        const lastSync = du.isValid(full) && du.isValid(delta) ? du.max([full, delta]) : full;

        // can only get messages in the last 8 months via delta endpoint, but add margin of error
        const canUseDelta = du.isValid(lastSync) && du.isAfter(lastSync, du.subMonths(du.now(), 7));

        const needFullSync = neverSynced || !canUseDelta;

        log.info(`full: [${du.isValid(full) ? full.toISOString() : "invalid"}], delta: [${du.isValid(delta) ? delta.toISOString() : "invalid"}], neverSynced: [${neverSynced}], canUseDelta: [${canUseDelta}]`);

        try {
            if (needFullSync) {
                log.info(`Starting full sync for channel [${channel.displayName}]`);
                await Sync.syncMessagesFull(client, channel, checkCancel, report);
            }
            else {
                log.info(`Starting incremental sync for channel [${channel.displayName}]`);
                await Sync.syncMessagesDelta(client, channel, checkCancel, report);
            }

            return true;
        } catch (error) {
            if (error instanceof OperationCancelled) throw error;
            AI.trackException({
                exception: error,
                properties: {
                    operation: nameof(Sync.syncTopLevelMessages),
                    teamId: channel.teamId,
                    channelId: channel.id,
                    channelName: channel.displayName,
                }
            });
            return false;
        }
    }


    /**
     * Get the date the channel was last synced.
     * Returns invalid Date if never synced.
     * @param channelId
     */
    static async getChannelLastSynced(channelId: string): Promise<Date> {
        const channel = await FindMsgChannel.get(channelId);

        if (!channel) return du.invalidDate();

        const { lastFullMessageSync: full, lastDeltaUpdate: delta } = channel;

        if (du.isValid(full) && du.isValid(delta)) return du.max([full, delta]);
        if (du.isValid(delta)) return delta;

        return full;
    }


    /**
     * Delta sync messages of the channel
     * @param client
     * @param channel
     */
    @log.traceAsync(true)
    private static async syncMessagesDelta(client: Client, channel: IFindMsgChannel, checkCancel: throwFn, report: (n: number) => void): Promise<void> {
        const { lastDeltaUpdate: delta, lastFullMessageSync: full } = channel;
        const last = du.isValid(delta) ? delta : full;

        if (!du.isValid(last)) {
            throw new Error("last delta sync invalid");
        }
        if (du.isBefore(last, du.subDays(du.subMonths(du.now(), 7), 1))) {
            throw new Error("last delta sync too old");
        }

        const cutOffTime = du.subMinutes(last, 5);
        const now = du.now();

        const response = await client.api(`/teams/${channel.teamId}/channels/${channel.id}/messages/delta`)
            .version('beta')
            .top(50) // max supported is 50
            .filter(`lastModifiedDateTime gt ${cutOffTime.toISOString()}`)
            .get();

        const it = new FullPageIterator<ChatMessage>(client, response, async (batch) => {
            const msgs = await Sync.saveDeltaMessageBatch(batch, channel, report);
            await Sync.getHostedImages(client, msgs);
            return true;
        });

        await it.iterate(checkCancel);

        channel.lastDeltaUpdate = now;
        await FindMsgChannel.put(channel);
    }


    /**
     * Sync top level messages of the channel.
     * This could take a long time for old and/or busy channels and the semantics of the API
     * force us to start from the beginning when interrupted.
     * @param client
     * @param channel
     */
    @log.traceAsync(true)
    private static async syncMessagesFull(client: Client, channel: IFindMsgChannel, checkCancel: throwFn, report: (n: number) => void): Promise<void> {
        const cutOffTime = du.now();

        const response = await client.api(`/teams/${channel.teamId}/channels/${channel.id}/messages`)
            .version('beta')
            .top(100) // maximum of 100 for this resource. Will still likely only return 50.
            .get();

        // delete all the old messages once the first request to graph succeeds
        const count = await db.channelMessages
            .where(idx.messages.$channelId$id)
            .between([channel.id, Dexie.minKey], [channel.id, Dexie.maxKey], true, true)
            .delete();

        log.info(`Deleted ${count} messages for channel [${channel.displayName}]`);

        const it = new FullPageIterator<ChatMessage>(client, response, async (batch) => {
            const msgs = await Sync.saveMessageBatch(batch, channel, report)
            await Sync.getHostedImages(client, msgs);
            return true;
        });
        await it.iterate(checkCancel);

        channel.lastFullMessageSync = cutOffTime;
        await FindMsgChannel.put(channel);
    }


    /**
     * Use PUT to save all supplied messages to the DB.
     * Returns true to signal iterator to continue.
     * @param batch
     * @param channel
     */
    @log.traceAsync()
    private static async saveMessageBatch(batch: ChatMessage[], channel: IFindMsgChannel, report: (n: number) => void): Promise<IFindMsgChannelMessage[]> {
        const msgs = await db.transaction("rw", db.channelMessages, db.users, async () => FindMsgChannelMessage.putAll(batch, channel));
        report(batch.length);
        log.info(`Wrote ${batch.length} messages to db`);
        return msgs;
    }


    /**
     * Update batch of messages in
     * @param batch
     * @param channel
     */
    @log.traceAsync()
    private static async saveDeltaMessageBatch(batch: ChatMessage[], channel: IFindMsgChannel, report: (n: number) => void): Promise<IFindMsgChannelMessage[]> {
        // Need to update the message AND the replies.
        // To keep this non-interruptible part of the the sync short, only delete the replies
        // of the incoming messages to avoid inconsistencies down the road.
        await db.transaction("rw", db.channelMessages, db.users, () =>
            Promise.all(batch.map(async (m): Promise<void> => {
                const deleted = await db.channelMessages
                    .where(idx.messages.channelId$replyToId$synced)
                    .between([channel.id, m.id, Dexie.minKey], [channel.id, m.id, Dexie.maxKey], true, true)
                    .delete();
                log.info(`Deleted ${deleted} replies of message [${m.id}]`);
            }))
        );

        return await this.saveMessageBatch(batch, channel, report);
    }


    /**
     * Download replies for top level messages of the channel that never had their replies synced.
     * This might take a long time but can be interrupted and resumed.
     * @param client
     * @param channel
     * @param batchSize number of messages to process in one call
     * @returns number of messages acutally processed (if same as batchSize, there may be more to process)
     */
    @log.traceAsync()
    private static async syncMessageRepliesBatch(client: Client, channel: IFindMsgChannel, batchSize: number, checkCancel: throwFn = nop, report: (n: number) => void = nop): Promise<number> {
        // get all top level messages where replies have been never synced or synced more than 7 months ago
        const now: Date = du.now();
        const messages: IFindMsgChannelMessage[] = await FindMsgChannelMessage.getTopLevelMessagesToSync(channel, batchSize);

        log.info(`Syncing replies for ${messages.length} messages`);

        /* Note: This looks like great use case for the graph batching endpoint, however Microsoft documentation indicated,
         * that each individual request inside the batch is counted against the throttling limit.
         * Not sure if the limit of "GET channel message 5 rps/(app*tenant)	100 rps/app" applies to this query.
         * So while we could save a few requests, we can not save any time (considering multiple people syncing at the same time).
         */
        for (const m of messages) {
            const response = await client.api(`/teams/${channel.teamId}/channels/${channel.id}/messages/${m.id}/replies`)
                .version('beta')
                .top(100) // exceeding 100 gives http 400
                .get();

            // this assumes, that a message has a reasonably small number of replies
            const replies = await getAllPages<ChatMessage>(client, response, checkCancel);
            report(replies.length);

            log.info(`API returned ${replies.length} replies for message [${m.id}]`);

            await db.transaction("rw", db.channelMessages, db.users, async () => {
                m.synced = now;
                await FindMsgChannelMessage.put(m);
                await FindMsgChannelMessage.putAll(replies, channel);
            });
        }

        // indicate how many messages where actually processed
        // it this is the same as batchSize, it indicates to the caller that there may be more
        return messages.length;
    }


    /**
     * Based on the last sync time, check if the team list is considered stale and should be refreshed.
     */
    private static async isTeamListStale(): Promise<boolean> {
        const synced = await getTeamsLastSynced();
        const cutOff = du.subHours(du.now(), 12);

        if (!du.isValid(synced)) return true;
        if (du.isBefore(synced, cutOff)) return true;

        return false;
    }


    /**
     * Based on the last sync time, check if the channel list is considered stale and should be refreshed.
     * @param team
     */
    private static isChannelListStale(team: IFindMsgTeam): boolean {
        // 12 hours is to ensure that the channel list is checked roughly once a day
        const cutOff = du.subHours(du.now(), 12);
        const last = team.lastChannelListSync;
        const isStale = !du.isValid(last) || du.isBefore(last, cutOff);
        log.info(`Team [${team.displayName}]: channel list last updated [${du.isValid(last) ? last.toISOString() : "invalid"}]. stale: [${isStale}]`);
        return isStale;
    }


    /**
     * Based on last sync time, check if the messages of the channel should be synced
     * @param channel
     */
    private static shouldSyncMessages(channel: IFindMsgChannel): boolean {
        const { lastDeltaUpdate: delta, lastFullMessageSync: full } = channel;
        // Do not sync more than once every 3 minutes
        const cutoff = du.subMinutes(du.now(), 3);
        let reason = "none";
        let shouldSync = false;

        if (!du.isValid(full)) {
            reason = "full sync date invalid";
            shouldSync = true;
        } else if (!du.isValid(delta) && du.isBefore(full, cutoff)) {
            reason = "delta invalid & full stale";
            shouldSync = true;
        } else if (du.isBefore(delta, cutoff)) {
            reason = "delta stale";
            shouldSync = true;
        }

        log.info(`${nameof(this.shouldSyncMessages)} [${channel.displayName}]: delta=[${delta}] full=[${full}] sync=[${shouldSync}] reason=[${reason}]`);

        return shouldSync;
    }


    /**
     * Delete channels that do not have associated teams
     * Interrupted sync might conceivably leave orphaned channels.
     */
    @log.traceAsync()
    static async deleteOrphanedChannels(): Promise<void> {
        const teams = await FindMsgTeam.getAll();
        const channels = await FindMsgChannel.getAll();

        for (const channel of channels) {
            // if the channel has a parent team, then skip
            if (teams.some(t => t.id === channel.teamId))
                continue;

            // channel is orphaned --> delete it an associated messages
            await db.transaction("rw", db.channels, db.channelMessages, db.users, async () => {
                await db.channels.delete(channel.id);
                await db.channelMessages.where(idx.messages.$channelId$id).between([channel.id, Dexie.minKey], [channel.id, Dexie.maxKey], true, true).delete();
            });
        }
    }


    /**
     * Delete messages that do not have associated channels
     * Interrupted sync might conceivably leave orphaned messages
     */
    @log.traceAsync()
    static async deleteOrphanedMessages(): Promise<void> {
        const channelIds = (await FindMsgChannel.getAll()).map(c => c.id);

        await db.transaction("rw", db.channelMessages, db.users, async () => {
            await db.channelMessages.filter(msg => channelIds.indexOf(msg.id) === -1).delete();
        });
    }

    /**
     * Automatically sync chats
     * @param client
     * @param checkCancel
     */
    @log.traceAsync(true)
    static async autoSyncChatAll(client: Client, checkCancel: throwFn = nop, progress: progressFn = nop, { chatList, replies }: ISyncProgressTranslation): Promise<boolean> {
        let success = true;

        if (await this.isChatListStale()) {
            progress(chatList);
            success = success && await Sync.syncChatList(client);
        }

        const chats = await FindMsgChat.getAll();
        let total = 0;
        for (const chat of chats) {
            try {
                checkCancel();
                progress(`Syncing chat messages... `);
                const reportCount = (n: number) => progress(replies(chat.id, total += n));
                await Sync.syncChatMembers(client, chat);
                await Sync.syncChatMesssages(client, chat, checkCancel, reportCount);
            } catch (error) {
                AI.trackException({
                    exception: error,
                    properties: {
                        operation: nameof(Sync.syncChatMesssages),
                        chatId: chat.id,
                        chatName: chat.topic,
                    }
                })
                success = false;
            }
        }

        return success;
    }


    /**
     * Based on the last sync time, check if the chat list is considered stale and should be refreshed.
     */
    private static async isChatListStale(): Promise<boolean> {
        const synced = await getChatsLastSynced();
        const cutOff = du.subMinutes(du.now(), 3);

        if (!du.isValid(synced)) return true;
        if (du.isBefore(synced, cutOff)) return true;

        return false;
    }


    /**
     * Fetch all chats the user is part of
     * The user can be part of 1000 chats at most, so this should be reasonably fast.
     * @param client
     */
    @log.traceAsync(true)
    private static async syncChatList(client: Client): Promise<boolean> {
        const existingChatIds = await db.chats.toCollection().primaryKeys();
        const fetchedChats = await Sync.fetchChatList(client);

        if (fetchedChats === null) {
            // Abort if we have no Chats at all
            return existingChatIds.length === 0 ? SyncError.ChatList() : false;
        } else {
            await db.transaction("rw", db.chats, db.channels, db.channelMessages, async () => {
                const chats = await Promise.all(fetchedChats.map(FindMsgChat.fromChat));
                await Promise.all(chats.map(t => FindMsgChat.put(t)));
            });

            await setChatsLastSynced(du.now());
            return true;
        }
    }


    private static async fetchChatList(client: Client): Promise<Chat[] | null> {
        try {
            const response = await client.api('/me/chats').version('beta').get();
            const fetchedChats = await getAllPages<Chat>(client, response);

            log.info(`API returned [${fetchedChats.length}] chats`);

            return fetchedChats;
        } catch (error) {
            AI.trackException({
                exception: error,
                properties: {
                    operation: nameof(Sync.fetchChatList),
                }
            });
            return null;
        }
    }


    @log.traceAsync()
    private static async syncChatMembers(client: Client, chat: IFindMsgChat): Promise<void> {
        const user: Entity = await client.api("/me").get();
        const selfId = assert(user?.id);

        // Note: Bots do not appear in the member list.
        const response = await client.api(`/me/chats/${chat.id}/members`)
            .version('beta')
            .get();

        // delete all the old messages once the first request to graph succeeds
        const delCount = await db.chatMembers
            .where(idx.chatMembers.$chatId$id)
            .between([chat.id, Dexie.minKey], [chat.id, Dexie.maxKey], true, true)
            .delete();

        log.info(`Deleted ${delCount} members for chat [${chat.id}]`);

        const it = new FullPageIterator<ConversationMember>(client, response, batch => Sync.saveChatMembersBatch(batch, chat, selfId));
        await it.iterate();
    }


    private static async saveChatMembersBatch(batch: ConversationMember[], chat: IFindMsgChat, selfId: string): Promise<true> {
        // Do not save logged in user as member
        const counterParts = batch.filter(cp => cp.id !== selfId);
        await db.transaction("rw", db.chatMembers, db.users, async () => FindMsgChatMember.putAll(counterParts, chat));
        log.info(`Wrote ${batch.length} messages to db`);
        return true;
    }


    /**
     * Sync top level messages of the channel.
     * This could take a long time for old and/or busy channels and the semantics of the API
     * force us to start from the beginning when interrupted.
     * @param client
     * @param chat
     */
    @log.traceAsync(true)
    private static async syncChatMesssages(client: Client, chat: IFindMsgChat, checkCancel: throwFn, report: (n: number) => void = nop): Promise<void> {
        const response = await client.api(`/me/chats/${chat.id}/messages`)
            .version('beta')
            .top(50) // maximum of 50 for this resource. Will still likely only return 50.
            .get();

        // delete all the old messages once the first request to graph succeeds
        const delCount = await db.chatMessages
            .where(idx.chatMessages.$chatId$id)
            .between([chat.id, Dexie.minKey], [chat.id, Dexie.maxKey], true, true)
            .delete();

        log.info(`Deleted ${delCount} messages for chat [${chat.id}]`);

        const it = new FullPageIterator<ChatMessage>(client, response, async (batch) => {
            const msgs = await Sync.saveChatMessageBatch(batch, chat, report);
            await Sync.getHostedImages(client, msgs);
            return true;
        });
        await it.iterate(checkCancel);

        await FindMsgChat.put(chat);
    }


    /**
     * Use PUT to save all supplied messages to the DB.
     * Returns true to signal iterator to continue.
     * @param batch
     * @param chat
     */
    @log.traceAsync()
    private static async saveChatMessageBatch(batch: ChatMessage[], chat: IFindMsgChat, report: (n: number) => void): Promise<IFindMsgChatMessage[]> {
        const msgs = await db.transaction("rw", db.chatMessages, db.users, async () => FindMsgChatMessage.putAll(batch, chat));
        report(batch.length);
        log.info(`Wrote ${batch.length} messages to db`);
        return msgs;
    }


    /**
     * Scan messages for images and replace them with cached versions
     *
     * Scan each message for images whose URL points to MsGraph hosted content.
     * Download such images and store them in the database.
     * Replace the <img> element in the body with a custom <graph-img> element.
     *
     * @param client
     * @param messages
     */
    private static async getHostedImages(client: Client, messages: (IFindMsgChatMessage | IFindMsgChannelMessage)[]) {
        log.info(`▼▼▼ getHostedImages START ▼▼▼`);
        const msgMaps: Array<IMsgMap> = [];

        const onloadCallBack = async (reader: FileReader, imgrec: IImageMap, parent: IMsgMap) => {
            // log.info(`★★★★★★★★★★ reader.onLoad start for id: [${imgrec.id}] ★★★★★★★★★★`);
            const result = reader.result;
            if (result && typeof result == 'string') {
                imgrec.dataUrl = result;
            }

            await db.images.put({
                id:imgrec.id, 
                data: imgrec.data, 
                srcUrl: imgrec.srcUrl,
                fetched: imgrec.fetched,
                dataUrl: imgrec.dataUrl,
            });

            const gimg = document.createElement("graph-image") as GraphImage;
            gimg.src = imgrec.id;
            imgrec.image.replaceWith(gimg);

            imgrec.done = true;

            let chkMsg = true;
            for (let j = 0; j < parent.images.length; j++) {
                const imgRec = parent.images[j];
                // log.info(`★★★★★★★★★★ checking images[${j}] => done? [${imgRec.done}] ★★★★★★★★★★`);
                if (!imgRec.done) {
                    // 一つでも未処理のイメージがあれば未完了とする
                    chkMsg = false;
                }
            }
            
            if (chkMsg) {
                await process(parent);
            }
        }

        const process = async (msgmap: IMsgMap) => {
            // msgmap.hasImage && log.info(`★★★★★★★★★★ getHostedImages body(before): [${msgmap.msg.body}] ★★★★★★★★★★`);
            msgmap.msg.body = msgmap.tmpl.innerHTML;
            // msgmap.hasImage && log.info(`★★★★★★★★★★ getHostedImages body(after): [${msgmap.msg.body}] ★★★★★★★★★★`);

            if ('chatId' in msgmap.msg) {
                await FindMsgChatMessage.put(msgmap.msg);
            } else {
                await FindMsgChannelMessage.put(msgmap.msg);
            }

            msgmap.completed = true;
        };

        const check = () => {
            const checker = setInterval( function() {
                // log.info(`★★★★★★★★★★ check process start ★★★★★★★★★★`);
                let done = true;
                // すべてのメッセージが処理済みかどうかをチェック
                for (let i = 0; i < msgMaps.length; i++) {
                    const rec = msgMaps[i];
                    // log.info(`★★★★★★★★★★ checking msgMaps[${i}] => completed? [${rec.completed}] ★★★★★★★★★★`);
                    if (!rec.completed) {
                        done = false;
                        break;
                    }

                }
                // log.info(`★★★★★★★★★★ check process end... done? [${done}] ★★★★★★★★★★`);
                if (done) {
                    clearInterval(checker);
                }
            }, 10);
        };

        // let withImage = 0;
        // let withoutImage = 0;
        for (const msg of messages.filter(m => m.type === "html")) {
            const tmpl = document.createElement("template");
            tmpl.innerHTML = msg.body;

            const msgmap: IMsgMap = {
                msg: msg,
                tmpl: tmpl,
                hasImage: false,
                completed: false,
                images: [],
            };
            msgMaps.push(msgmap);

            for (const image of Array.from(tmpl.content.querySelectorAll("img"))) {
                const srcUrl = image.src;

                if (isGraphHostedContentUrl(srcUrl)) {
                    try {

                        const data: Blob = await client.api(srcUrl).responseType(ResponseType.BLOB).get();
                        const id = await hashBlob(data);
                        
                        msgmap.hasImage = true;
                        const imgRec: IImageMap = {
                            id: id,
                            data: data,
                            srcUrl: srcUrl,
                            fetched: new Date().getTime(),
                            dataUrl: "",
                            image: image,
                            done: false,
                        };
                        msgmap.images.push(imgRec);
                        
                    } catch (error) {
                        log.error(error);
                        AI.trackException({ error });
                    }
                }
            }
            // if (msgmap.hasImage) {
            //     withImage += 1;
            // } else {
            //     withoutImage += 1;
            // }

            // for (const image of Array.from(tmpl.content.querySelectorAll("img"))) {
            //     const srcUrl = image.src;

            //     if (isGraphHostedContentUrl(srcUrl)) {
            //         try {
            //             // Note: download as Blob instead of ArrayBuffer because Blob contains the mime type
            //             const data: Blob = await client.api(srcUrl).responseType(ResponseType.BLOB).get();
            //             const id = await hashBlob(data);

            //             await db.images.put({
            //                 id, data, srcUrl,
            //                 fetched: new Date().getTime(),
            //             });

            //             const gimg = document.createElement("graph-image") as GraphImage;
            //             gimg.src = id;
            //             image.replaceWith(gimg);
            //         } catch (error) {
            //             log.error(error);
            //             AI.trackException({ error });
            //         }
            //     }
            // }

            // msg.body = tmpl.innerHTML;

            // if ('chatId' in msg) {
            //     await FindMsgChatMessage.put(msg);
            // } else {
            //     await FindMsgChannelMessage.put(msg);
            // }
        }

        // log.info(`★★★★★★★★★★ HTML Message count: [${msgMaps.length}] withImage: [${withImage}] withoutImage:[${withoutImage}] ★★★★★★★★★★`);

        check();

        const processImage = async (rec: IMsgMap) => {
            if (!rec.hasImage) {
                await process(rec);
            } else {
                rec.images.forEach(imgrec => {
                    // log.info(`★★★★★★★★★★ Image processing start for id: [${imgrec.id}] ★★★★★★★★★★`);
                    const reader = new FileReader;
                    reader.onload = async () => {
                        await onloadCallBack(reader, imgrec, rec);
                    };

                    reader.readAsDataURL(imgrec.data);
                    // log.info(`★★★★★★★★★★ Image processing end for id: [${imgrec.id}] ★★★★★★★★★★`);
                })
            }
        };

        for (let i = 0; i < msgMaps.length; i++) {
            const rec = msgMaps[i];
            // log.info(`★★★★★★★★★★ msgMaps[${i}] rec.completed: [${rec.completed}] rec.hasImage: [${rec.hasImage}] ★★★★★★★★★★`);
            await processImage(rec);
        }

        log.info(`▲▲▲ getHostedImages END ▲▲▲`);
    }
}

interface IImageMap extends IFindMsgImageDb {
    image: HTMLImageElement;
    done: boolean;
}

interface IMsgMap {
    msg: IFindMsgChatMessage | IFindMsgChannelMessage;
    tmpl: HTMLTemplateElement;
    hasImage: boolean;
    completed: boolean;
    images: IImageMap[];
}