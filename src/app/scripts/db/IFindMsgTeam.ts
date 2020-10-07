export interface IFindMsgTeam {
    /** team id */
    id: string;

    /** team name */
    displayName: string;

    /** team description */
    description: string | null;

    /** internal team url */
    webUrl: string | null;

    /** timestamp of last channel list update */
    lastChannelListSync: Date;
}
