import Dexie from 'dexie';
import { info, traceAsync, warn } from '../logger';
import { isUnknownObject } from '../utils';

type Aggregator = { keys: number, values: number };

function toMb(n: number): string {
    return (n < 0 ? "0" : (n / (1024 * 1024)).toFixed(2)) + 'MB'
}

export class DbStatAggregator {
    private enc: TextEncoder;
    private count: number;
    private agg: Aggregator;

    constructor() {
        this.enc = new TextEncoder();
        this.count = 0;
        this.agg = { keys: 0, values: 0 };
    }

    /**
     * Gather and ouput statistics about the database.
     * Reported key and value sizes are only rough estimates.
     * The numbers are generally too high for chrome (because of some optimizations in chrome)
     * and too low for firefox compared to the numbers reported navigator.storage.estimate()
     * @param db
     */
    @traceAsync()
    async analyzeDb(db: Dexie): Promise<void> {
        if (navigator.storage && navigator.storage.estimate) {
            const { usage, quota } = await navigator.storage.estimate();
            const used = usage ? toMb(usage) : "UNKNOWN";
            const quot = quota ? toMb(quota) : "UNKNOWN";
            info(`Browser reports ${used} of ${quot} used.`);
        } else {
            warn(`Browser does not support used storage estimation.`);
        }

        info(`Analyzing DB ${db.name}...`);
        info("---");

        for (const table of db.tables) {
            await this.analyzeTable(table);
        }

        info(`Total database size: ${this.count} entries, keySize ${toMb(this.agg.keys)}, valueSize ${toMb(this.agg.values)}`);
        info("---");
        info(`Done`);
    }


    async analyzeTable(table: Dexie.Table): Promise<void> {
        const ex = this.getKeyExtractor(table);
        let count = 0;
        const agg = {
            keys: 0,
            values: 0
        };

        await table.each(o => {
            ++count;
            this.estimateObjectStorageSize(o, this.enc, 10, agg);
            this.estimateObjectStorageSize(ex(o), this.enc, 10, agg);
        });

        info(`Table ${table.name}: ${count} entries, keySize ${toMb(agg.keys)}, valueSize ${toMb(agg.values)}`);

        this.count += count;
        this.agg.keys += agg.keys;
        this.agg.values += agg.values;
    }


    /**
     * Creates a function that extracts the key/index properties from an object.
     * Does not account for null/undefined values not being indexed.
     */
    private getKeyExtractor(table: Dexie.Table): (o: { [key: string]: unknown; }) => unknown[] {
        const idx = table.schema.indexes.concat(table.schema.primKey).flatMap(i => i.keyPath ?? []);
        return (o) => idx.map(i => o[i]);
    }

    /**
     * Very inaccurate estimation of the storage size of an object
     * based on various more or less unfounded assumptions.
     * @param o The object to estimate
     * @param enc a TextEncoder instance
     * @param agg an aggregator object instance
     * @param d a maximum recursion depth
     */
    private estimateObjectStorageSize(o: unknown, enc: TextEncoder, d: number, agg: Aggregator = { keys: 0, values: 0 }): Aggregator {
        if (d <= 0) {
            warn("Recursion depth exceeded");
            return agg;
        }

        if (!isUnknownObject(o)) {
            warn("Unsupported object");
            return agg;
        }

        Object.entries(o).forEach(entry => this.estimateValueSize(entry, enc, d, agg));

        return agg;
    }


    private estimateValueSize([k, v]: [string, unknown], enc: TextEncoder, d: number, agg: Aggregator = { keys: 0, values: 0 }): void {
        agg.keys += enc.encode(k).length;

        agg.values += 1; // assume 1 byte for type storage

        switch (typeof v) {
            case "string":
                agg.values += 2 + enc.encode(v).length;
                break;

            case "number":
                agg.values += 8;
                break;

            case "boolean":
            case "undefined":
                agg.values += 1;
                break;

            case "object":
                if (v === null)
                    agg.values += 1;
                else if (v instanceof Array)
                    v.forEach(ve => this.estimateValueSize(ve, enc, d - 1, agg));
                else if (v instanceof ArrayBuffer)
                    agg.values += v.byteLength;
                else if (v instanceof Blob)
                    agg.values += v.size;
                else if (v instanceof Date)
                    agg.values += enc.encode(v.toISOString()).length;
                else
                    this.estimateObjectStorageSize(v, enc, d - 1, agg);
                break;

            case "function":
                // Silently ignore functions
                break;

            default:
                warn(`Ignoring unsupported type ${typeof v}`);
                break;
        }
    }
}
