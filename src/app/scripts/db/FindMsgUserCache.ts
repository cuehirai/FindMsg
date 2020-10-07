import { db } from './Database';
import { IFindMsgUserDb } from "./IFindMsgUserDb";
import { IUserCacheItem } from './IUserCacheItem';
import { IFindMsgUser } from './IFindMsgUser';
import { isBefore, dateToNumber, numberToDate } from '../dateUtils';


/**
 * Provides fast userId => displayName lookup
 */
export class FindMsgUserCache {

    private static instance = new FindMsgUserCache();
    private cache = new Map<string, IUserCacheItem>();
    private cacheInitDone: Promise<void>;
    private changed = 0;


    private constructor() {
        this.cacheInitDone = db.users.each(user => this.cache.set(user.id, {
            displayName: user.displayName,
            lastUpdated: numberToDate(user.lastUpdated),
            writtenSinceLoad: false,
        }));
    }


    /**
     * Get the singleton instance of
     */
    static async getInstance(): Promise<FindMsgUserCache> {
        await this.instance.cacheInitDone;
        return this.instance;
    }


    /**
     * Writes the current state of the cache to IndexedDB
     */
    //@traceAsync
    async persistCache(): Promise<void> {
        if (this.changed === 0) {
            return;
        }

        const updatedUsers: IFindMsgUserDb[] = [];

        for (const [id, info] of this.cache.entries()) {
            if (info.writtenSinceLoad) {
                updatedUsers.push({
                    id,
                    displayName: info.displayName,
                    lastUpdated: dateToNumber(info.lastUpdated),
                });
            }
        }

        await db.transaction("rw", db.users, () => db.users.bulkPut(updatedUsers));
        this.changed = 0;
    }


    /**
     * Resolves an userId to a user's displayName.
     * Returns null if userId is not found in cache.
     * @param userId
     */
    getDisplayName(userId: string): string | null {
        return this.cache.get(userId)?.displayName ?? null;
    }


    /**
     * Update the name of the user in the cache.
     * Updates only if the user is not yet in the cache or if the stored name is older than timestamp
     * @param userId
     * @param displayName
     * @param timestamp
     */
    updateUserName(userId: string, displayName: string, timestamp: Date): void {
        const item = this.cache.get(userId);

        if (item) {
            if (isBefore(item.lastUpdated, timestamp)) {
                if (displayName !== item.displayName) {
                    item.displayName = displayName;
                    item.lastUpdated = timestamp;
                    item.writtenSinceLoad = true;
                    this.changed++;
                }
            }
        }
        else {
            this.cache.set(userId, {
                displayName,
                lastUpdated: timestamp,
                writtenSinceLoad: true,
            });
            this.changed++;
        }
    }


    /**
     * Get a list of known users
     */
    getKnownUsers = (): Promise<IFindMsgUser[]> => db.users.toCollection().sortBy(nameof<IFindMsgUserDb>(u => u.displayName), l => l.map(u => ({
        ...u,
        lastUpdated: numberToDate(u.lastUpdated)
    })));
}
