import { AppConfig } from "../../../../config/AppConfig";
import { DbAccessorBaseComponent, ISubEntityFunctionArg, SubEntityFunction, SubEntytyAllFunction } from "../db-accessor-class-base";
import { IImageDb } from "./IImageDb";
import * as log from '../../logger';
import { b64toBlob, blob2dataUrl } from "../../utils";
import { db } from "../Database";

class ImageEntity<D extends IImageDb, T extends IImageDb, A extends IImageDb> extends DbAccessorBaseComponent<D, T, A> {
    tableName = "images";
    lastSyncedKey = `${AppConfig.AppInfo.name}_images_last_synced`;
    isDeltaSyncAvailable = false;
    parseApi = async (api: A): Promise<T | null> => {
        const res: IImageDb = api;
        return res as T;
    }
    fromDbEntity(dbrec: D): T {
        const res: IImageDb = dbrec;        
        return res as T;
    }
    toDbEntity(teamsrec: T): D {
        const res: IImageDb = teamsrec;
        return res as D;
    }
    protected getSubEntity: SubEntityFunction = async (): Promise<void> => {/* 実装なし */};
    protected putSubEntity: SubEntityFunction = async (): Promise<void> => {/* 実装なし */};
    protected putAllSubEntity: SubEntytyAllFunction = async (): Promise<void> => {/* 実装なし */};
    protected createSubEntityArg(parent: T | null, parentDb: D | null): ISubEntityFunctionArg {
        return {parent, parentDb};
    }
    protected fetchApiAll(): Promise<T[]> {
        // 敢えて実装しません。使用しないでください
        throw new Error("Method not implemented.");
    }
    protected fetchApiDelta(): Promise<T[]> {
        // 敢えて実装しません。使用しないでください
        throw new Error("Method not implemented.");
    }
    protected syncSubentity(): Promise<boolean> {
        // 敢えて実装しません。使用しないでください
        throw new Error("Method not implemented.");
    }

    async get(id: string): Promise<T | null> {
        log.info(`▼▼▼ ` + this.tableName + `.get START (id=` + id +`) ▼▼▼`);
        const result = await db.images.get(id);
        const res = result? this.fromDbEntity(result as D): null;
        if (res) {
            const b64 = (): string => {
                let concat = res.dataUrl?? "";
                res.dataChunk.forEach((chunk) => concat += chunk)
                return concat;
            }
            res.data = b64toBlob(b64());
        }
        log.info(`▲▲▲ ` + this.tableName + `.get END (id=` + id +`) ▲▲▲`);
        return res;        
    }

    async getAll(): Promise<T[]> {
        log.info(`▼▼▼ ` + this.tableName + `.getAll START ▼▼▼`);
        const results = await db.images.toArray();
        const res: Array<T> = [];
        results.forEach(async dbrec => {
            const teamsrec = await this.get(dbrec.id);
            if (teamsrec) {
                res.push(teamsrec as T);
            }
        });
        log.info(`▲▲▲ ` + this.tableName + `.getAll END ▲▲▲`);
        return res;
    }

    async put(teamsrec: T | null): Promise<void> {
        log.info(`▼▼▼ ` + this.tableName + `.put START (id=` + teamsrec?.id +`) ▼▼▼`);
        const dbrec = teamsrec ? this.toDbEntity(teamsrec) : null;
        if (teamsrec && dbrec) {
            const value = await blob2dataUrl(teamsrec.data);
            if (value.length <= 409600) {
                dbrec.dataUrl = value;
            } else {
                dbrec.dataChunk = [];
                let rest = value;
                while(rest.length > 0) {
                    const len = rest.length;
                    if (len < 409600) {
                        dbrec.dataChunk.push(rest);
                        rest = "";
                    } else {
                        dbrec.dataChunk.push(rest.substring(0, 409600));
                        rest = rest.substring(409600)
                    }
                }                    
            }
            await db.images.put(dbrec);
        }       
        log.info(`▲▲▲ ` + this.tableName + `.put END (id=` + teamsrec?.id +`) ▲▲▲`);
    }

    async putAll(teamsrecs: T[]): Promise<void> {
        log.info(`▼▼▼ ` + this.tableName + `.putAll START ▼▼▼`);
        teamsrecs.forEach(async (teamsrec) => await this.put(teamsrec));
        log.info(`▲▲▲ ` + this.tableName + `.putAll END ▲▲▲`);
    }
}

export const ImageTable = new ImageEntity;