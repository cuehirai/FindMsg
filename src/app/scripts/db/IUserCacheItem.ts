export interface IUserCacheItem {
    /** user display name */
    displayName: string;

    /** timestamp of last update */
    lastUpdated: Date;

    /** if this item was updated since it was last written to persistent storage */
    writtenSinceLoad: boolean;
}
