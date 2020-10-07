export interface IFindMsgChannelDb {
    /** channel id */
    id: string;

    /** channel name */
    displayName: string;

    /** channel description */
    description: string | null;

    /** internal channel url */
    webUrl: string;

    /** id of the team to which the channel belongs */
    teamId: string;

    /** timestamp of last delta message sync */
    lastDeltaUpdate: number;

    /** timestamp of last full sync */
    lastFullMessageSync: number;
}
