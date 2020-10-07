import { Client, PageCollection } from '@microsoft/microsoft-graph-client';
import { assert, throwFn, nop } from '../utils';
import * as log from '../logger';

export declare type FullPageIteratorCallback<T> = (data: T[]) => Promise<boolean>;


/**
 * Version of @microsoft/microsoft-graph-client {PageIterator} that calls the callback once for every result page returned from the API
 */
export class FullPageIterator<T> {
    client: Client;
    collection: T[];
    nextLink: string | undefined;
    callback: FullPageIteratorCallback<T>;
    complete: boolean;
    page: number;


    constructor(client: Client, pageCollection: PageCollection, callback: FullPageIteratorCallback<T>) {
        this.client = client;
        this.collection = pageCollection.value;
        this.nextLink = pageCollection["@odata.nextLink"];
        this.callback = callback;
        this.complete = false;
        this.page = 1;
    }


    public async iterate(checkCancel: throwFn = nop): Promise<void> {
        log.info(`Page 1 has ${this.collection.length} items`);
        let advance = await this.callback(this.collection);
        while (advance) {
            if (this.nextLink === undefined) {
                advance = false;
            }
            else {
                checkCancel();
                await this.fetchAndUpdateNextPageData();
                log.info(`Page ${this.page} has ${this.collection.length} items`);
                advance = this.collection.length === 0 ? true : await this.callback(this.collection);
            }
        }
        if (this.nextLink === undefined) {
            this.complete = true;
        }
    }


    private async fetchAndUpdateNextPageData(): Promise<void> {
        try {
            const next = assert(this.nextLink, nameof(this.nextLink));
            const response: PageCollection = await this.client.api(next).get();
            this.collection = response.value;
            this.nextLink = response["@odata.nextLink"];
            ++this.page;
        }
        catch (error) {
            log.error(error);
            throw error;
        }
    }
}
