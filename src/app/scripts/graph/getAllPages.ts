import { Client, PageCollection, PageIterator } from '@microsoft/microsoft-graph-client';
import { throwFn, nop } from '../utils';

/**
 * Fetch the data from all pages of the microsoft graph response
 * @param client
 * @param response
 */
export async function getAllPages<T>(client: Client, response: PageCollection, checkCancel: throwFn = nop): Promise<T[]> {
    const items: T[] = [];
    const it = new PageIterator(client, response, (r: T) => {
        items.push(r);
        checkCancel();
        return true;
    });

    await it.iterate();

    return items;
}
