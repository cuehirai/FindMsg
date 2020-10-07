export interface IFindMsgTeamDb {
    /** team id */
    id: string;

    /** team name */
    displayName: string;

    /** team description */
    description: string | null;

    /** internal team url */
    webUrl: string | null;

    /** timestamp of last channel list update */
    lastChannelListSync: number;
}
