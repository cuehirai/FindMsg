import { Client } from "@microsoft/microsoft-graph-client";
import { Drive, DriveItem, UploadSession } from "@microsoft/microsoft-graph-types";
import { getAllPages } from "./graph/getAllPages";
import * as log from './logger';

export interface DriveItemExtended extends DriveItem {
    downloadUrl : string | null;
}

interface IFileChooser {
    name: string;
    isRegExp: boolean;
}

/**
 * 対象ファイル判定クラス
 */
export class FileChooser {
    private names: IFileChooser[] = [];
    /**
     * 対象ファイル名追加
     * @param name ファイル名またはパターン
     * @param isRegExp パターンを指定した場合はtrue
     */
    public add(name: string, isRegExp: boolean): FileChooser {
        const item: IFileChooser = {name, isRegExp};
        this.names.push(item);
        return this;
    }
    /**
     * 対象ファイルかどうかをテスト※対象ファイル条件に何も登録していない場合は無条件に対象と判定します
     * @param filename テストするファイル名
     */
    public test(filename: string | null | undefined): boolean {
        let res = false;
        if (filename) {
            if (this.names.length > 0) {
                for (let i = 0; i < this.names.length; i++) {
                    const item = this.names[i];
                    if (item.isRegExp) {
                        const reg = new RegExp(item.name);
                        if (reg.test(filename)) {
                            log.info(`@@@@@ filename [${filename} matches pattarn [${item.name}] @@@@@]`);
                            res = true;
                            break;
                        }
                    } else {
                        if (item.name === filename) {
                            log.info(`@@@@@ filename [${filename} matches name [${item.name}] @@@@@]`);
                            res = true;
                            break;
                        }
                    }
                }
                if (!res) {
                    log.info(`@@@@@ filename [${filename}] does not match any pattern or name @@@@@]`);
                }
            } else {
                res = true;
            }    
        }
        return res;
    }
}

const reformPathString = (path?: string | null | undefined) => {
    return path? (path.startsWith("root/"))? path.substr(5) : path : "root"
}

/** ファイル操作ユーティリティクラス */
class Util {

    /**
     * ログインユーザのDriveリソースを取得します。
     * @param client MicrosoftGraphクライアント
     */
    public async getDrive(client: Client): Promise<Drive | null> {
        let res: Drive | null = null;
        try {
            const address = `me/drive`;
            log.info(`★★★ requesting api(get): [${address}] ★★★`);
            res = await client.api(address).get();
        } catch (e) {
            log.warn(`getDrive failed: error = [${e}]`);
        }
        return res;
    }

    /**
     * 指定フォルダにあるドライブ項目(フォルダ/ファイル)をすべて取得します。
     * @param client MicrosoftGraphクライアント
     * @param parentPath 走査したいフォルダまでのパス(パス区切り文字は「/」)※パスを省略するとroot直下を調べます。またパスを指定する際はrootは省略可能です。
     * @param createFolderIfNotExist trueを指定するとパス内に存在しないフォルダがある場合にフォルダを作成します。
     */
    public async getItems(client: Client, parentPath?: string | null | undefined, createFolderIfNotExist?: boolean, chooser?: FileChooser): Promise<DriveItem[]> {
        const target = chooser?? new FileChooser;
        // 目的のフォルダをrootからのパスで指定するが、rootは省略可能
        const parent = reformPathString(parentPath);
        // log.info(`▼▼▼ getItems START parentPath: [${parent}] createIfNotExist: [${createFolderIfNotExist?? false}] ▼▼▼`);
        const res: DriveItem[] = [];
        if (parent === "root") {
            const address = `me/drive/root/children`;
            log.info(`★★★ requesting api(get): [${address}] ★★★`);
            const api = await client.api(address).get();
            const fetched = await getAllPages<DriveItem>(client, api);
            fetched.forEach(rec => {
                if (target.test(rec.name)){
                    res.push(rec);
                }
            });
            // log.info(`★★★ Found [${res.length}] items in root folder ★★★`);
        } else {
            const paths = parent.split('/');
            let folder: DriveItem | null = null;
            for (let i = 0; i < paths.length; i++ ) {
                const found = await this.findFolder(client, paths[i], folder, createFolderIfNotExist);
                if (!found) {
                    folder = null;
                    break;
                } else {
                    folder = found;
                }
            }
            if (folder && folder.id) {
                const address = `/me/drive/items/${folder.id}/children`;
                // log.info(`★★★ requesting api(get): [${address}] ★★★`);
                const api = await client.api(address).get();
                const fetched = await getAllPages<DriveItem>(client, api);
                fetched.forEach(rec => {
                    if (target.test(rec.name)){
                        res.push(rec);
                    }
                });
                // log.info(`★★★ Found [${res.length}] items in the folder [${folder.name}] ★★★`);
            }
        }
        // log.info(`▲▲▲ getItems END parentPath: [${parent}] createIfNotExist: [${createFolderIfNotExist?? false}] ▲▲▲`);
        return res;
    }

    /**
     * 指定フォルダを取得します。
     * @param client MicrosoftGraphクライアント
     * @param folderPath 取得したいフォルダのパス(パス区切り文字は「/」)※パスを省略するとrootを取得します。またパスを指定する際はrootは省略可能です。
     * @param createFolderIfNotExist trueを指定するとパス内に存在しないフォルダがある場合にフォルダを作成します。
     */
    public async getFolder(client: Client, folderPath: string, createFolderIfNotExist?: boolean): Promise<DriveItem | null> {
        // log.info(`▼▼▼ getFolder START folderPath: [${folderPath}] createIfNotExist: [${createFolderIfNotExist?? false}] ▼▼▼`);
        let res: DriveItem | null = null;
        const path = reformPathString(folderPath);

        if (path.indexOf("/") < 0) {
            if (path === "root") {
                const address = `/me/drive/root`;
                log.info(`★★★ requesting api(get): [${address}] ★★★`);
                res = await client.api(address).get();
                // log.info(`★★★ Found the root folder ID: [${res?.id}] ★★★`);
            } else {
                // root直下のフォルダの場合
                res = await this.findFolder(client, path, null, createFolderIfNotExist);
            }
        } else {
            // 親階層のフォルダを特定した上で対象のフォルダを取得する
            const parent = path.substr(0, path.lastIndexOf("/"));
            const find = path.substr(parent.length + 1);
            // log.info(`Calling getFolder recursive; parent: [${parent}] find: [${find}]`);
            const parentFolder = await this.getFolder(client, parent, createFolderIfNotExist);
            if (parentFolder) {
                res = await this.findFolder(client, find, parentFolder, createFolderIfNotExist);
            }
        }

        // log.info(`▲▲▲ getFolder END folderPath: [${folderPath}] createIfNotExist: [${createFolderIfNotExist?? false}] resultName: [${res? res.name : "(not found)"}] resultId: [${res? res.id : "(not found)"}] ▲▲▲`);
        return res;
    }

    /**
     * 指定ファイルを取得します。ファイルが見つからない場合はnullを返却します。
     * @param client MicrosoftGraphクライアント
     * @param find ファイル名
     * @param parentPath フォルダのパス(パス区切り文字は「/」)※パスを省略するとrootを取得します。またパスを指定する際はrootは省略可能です。
     * @param createFolderIfNotExist trueを指定するとパス内に存在しないフォルダがある場合にフォルダを作成します。
     */
    public async getFile(client: Client, find: string, parentPath?: string | null | undefined, createFolderIfNotExist?: boolean): Promise<DriveItem | null> {
        // log.info(`▼▼▼ getFile START parentPath: [${parentPath}] createIfNotExist: [${createFolderIfNotExist?? false}] ▼▼▼`);
        let res: DriveItem | null = null;

        const parent = reformPathString(parentPath);
        const items = await this.getItems(client, parent, createFolderIfNotExist);
        // log.info(`★★★ [${items.length}] items found in [${parent}] ★★★`);
        for (let i = 0; i < items.length; i++) {
            // log.info(`★★★ file [${i}] in [${parent}]: name=[${items[i].name}] id=[${items[i].id}] ★★★`);
            if (items[i].name === find) {
                if (items[i].file) {
                    res = items[i];
                    // log.info(`★★★ Target found. name=[${res.name}] id=[${res.id}] ★★★`);
                    break;
                }
            }
        }

        // log.info(`▲▲▲ getFile END parentPath: [${parentPath}] createIfNotExist: [${createFolderIfNotExist?? false}] ▲▲▲`);
        return res;
    }

    /**
     * 指定ファイルに文字列データを書き込みます。
     * @param client MicrosoftGraphクライアント
     * @param fileName ファイル名
     * @param content ファイルに書き込むファイル
     * @param parentPath フォルダのパス(パス区切り文字は「/」)※パスを省略するとrootを取得します。またパスを指定する際はrootは省略可能です。★存在しないフォルダは作成されます。
     * @param overwrite trueを指定すると指定ファイル名がすでに存在する場合に上書きします。
     */
    public async writeFile(client: Client, fileName: string, content: string, parentPath?: string | null | undefined, overwrite?: boolean): Promise<boolean> {
        // log.info(`▼▼▼ writeFile START parentPath: [${parentPath}] fileName: [${fileName}] overwrite:[${overwrite}] ▼▼▼`);
        let res = false;

        let ow = overwrite?? false;

        try {
            const parent = reformPathString(parentPath);
            let file = await this.getFile(client, fileName, parentPath, true);
            if (!file) {
                // ファイルが存在しない場合は先にファイルの「枠」だけ作成
                if (parent === "root") {
                    const address = `/me/drive/root:/${fileName}/content`;
                    log.info(`★★★ requesting api(put): [${address}] ★★★`);
                    file = await client.api(address).put("");
                    if (file) {
                        ow = true;
                        // log.info(`★★★ File [${fileName}] created in the root folder... New ID: [${file.id}] ★★★`);
                    } else {
                        log.warn(`★★★ File [${fileName}] failed to create in the root folder ★★★`);
                    }
                } else {
                    const folder = await this.getFolder(client, parent, true);
                    if (folder && folder.id) {
                        const address = `/me/drive/items/${folder.id}:/${fileName}:/content`;
                        log.info(`★★★ requesting api(put): [${address}] ★★★`);
                        file = await client.api(address).put("");
                        if (file) {
                            ow = true;
                            // log.info(`★★★ File [${fileName}] created in the folder [${folder.name}]... New ID: [${file.id}] ★★★`);
                        } else {
                            log.warn(`★★★ File [${fileName}] failed to create in the folder [${folder.name}] ★★★`);
                        }
                    }
                }
            }

            let item: DriveItem | null = null;
            if (file && file.id) {
                if (ow) {
                    const buf = Buffer.from(content);
                    const wholesize = buf.length;
                    if (wholesize < (1024 * 1024 * 4)) {
                        // 4MB未満のファイルは直接アップロード
                        const address = `/me/drive/items/${file.id}/content`;
                        log.info(`★★★ requesting api(put): [${address}] ★★★`);
                        item = await client.api(address).put(content);
                        if (item) {
                            res = true;
                            // log.info(`★★★ File [${item.name}] ID: [${item.id}] overwritten ★★★`);    
                        } else {
                            log.warn(`★★★ File [${fileName}] failed to write ★★★`);
                        }
                    } else {
                        // 4MB以上のファイルはアップロードセッションで分割してアップロード
                        const address1 = `/me/drive/items/${file.id}/createUploadSession`;
                        log.info(`★★★ requesting api(post): [${address1}] ★★★`);
                        const session: UploadSession | null | undefined = await client.api(address1).post("");
                        if (session && session.uploadUrl) {
                            const url = session.uploadUrl;
                            // log.info(`Upload URL: [${url}] in response: [${JSON.stringify(session)}]`);

                            const max = (320 * 1024) * 10;
                            let start = 0;
                            while (start < wholesize) {
                                const rest = wholesize - start;
                                let len = max;
                                if (rest < max) {
                                    len = rest;
                                }
                                const end = start + len;
                                const send = buf.slice(start, end);
                                // log.info(`★★★ Sending ${len} (actual ${send.length}) bytes of ${start}-${end - 1}/${wholesize} ★★★ `);
                                const req = new XMLHttpRequest();
                                req.open("PUT", url, false);
                                req.setRequestHeader("Content-Range", `bytes ${start}-${end - 1}/${wholesize}`);
                                req.send(send);
                                // log.info(`★★★  Responce: ${JSON.stringify(req.response)} ★★★ `);
                                start = end;
                            }
                            res = true;
                        }
                    }
                } else {
                    // log.warn(`★★★ File [${fileName}] exists and was not overwritten ★★★`);
                }
            }
        } catch (err) {
            log.error(`Error in writing file [${fileName}]: [${JSON.stringify(err)}]`);
        }

        // log.info(`▲▲▲ writeFile END parentPath: [${parentPath}] fileName: [${fileName}] ▲▲▲`);

        return res;
    }

    /**
     * 指定テキストファイルを読み込みます。ファイルが見つからない場合はnullを返却します。
     * @param client MicrosoftGraphクライアント
     * @param fileName ファイル名
     * @param parentPath フォルダのパス(パス区切り文字は「/」)※パスを省略するとrootを取得します。またパスを指定する際はrootは省略可能です。★存在しないフォルダを作成しません。
     */
    public async readFile(client: Client, fileName: string, parentPath?: string | null | undefined): Promise<string | null> {
        // log.info(`▼▼▼ readFile START parentPath: [${parentPath}] fileName: [${fileName}] ▼▼▼`);

        let res: string | null = null;
        const file = await this.getFile(client, fileName, parentPath);
        if (file) {
            const fileEx: DriveItemExtended =JSON.parse(JSON.stringify(file).replace("@microsoft.graph.downloadUrl", "downloadUrl"));
            if (fileEx.downloadUrl) {
                // log.info(`★★★ readFile file found id: [${fileEx.id}] downloadUrl: [${fileEx.downloadUrl}] ★★★`);
                const downloadUrl = fileEx.downloadUrl;
                // log.info(`★★★ readFile file downloadUrl: [${downloadUrl}] ★★★`);
    
                const request = new XMLHttpRequest();
                request.open("GET", downloadUrl, false);
                request.send();
                res = request.response;
                // log.info(`★★★ File [${fileName}] found and fetched... ID: [${file.id}] ★★★`);
            }
        } else {
            // log.info(`★★★ File [${fileName}] not found ★★★`);
        }
        // log.info(`▲▲▲ readFile END parentPath: [${parentPath}] fileName: [${fileName}] ▲▲▲`);
        return res;
    }

    /**
     * 指定ファイルを削除します。
     * @param client MicrosoftGraphクライアント
     * @param fileName ファイル名
     * @param parentPath フォルダのパス(パス区切り文字は「/」)※パスを省略するとrootを取得します。またパスを指定する際はrootは省略可能です。★存在しないフォルダを作成しません。
     */
    public async deleteFile(client: Client, fileName: string, parentPath?: string | null | undefined): Promise<boolean> {
        // log.info(`▼▼▼ deleteFile START parentPath: [${parentPath}] fileName: [${fileName}] ▼▼▼`);

        let res = true;
        const file = await this.getFile(client, fileName, parentPath);
        if (file) {
            try {
                const address = `/me/drive/items/${file.id}`;
                log.info(`★★★ requesting api(delete): [${address}] ★★★`);
                await client.api(address).delete();
            } catch (e) {
                res = false;
                log.error(`deleteFile failed: [${e}]`);
            }
        } else {
            // log.info(`★★★ File [${fileName}] not found ★★★`);
        }

        // log.info(`▲▲▲ deleteFile END parentPath: [${parentPath}] fileName: [${fileName}] ▲▲▲`);
        return res;
    }

    /**
     * 指定したフォルダを削除します。
     * @param client MicrosoftGraphクライアント
     * @param folderPath 削除したいフォルダのパス(パス区切り文字は「/」)※パスを省略するとrootを取得します。またパスを指定する際はrootは省略可能です。
     */
    public async deleteFolder(client: Client, folderPath: string): Promise<boolean> {
        // log.info(`▼▼▼ deleteFolder START folderName: [${folderPath}] ▼▼▼`);

        let res = true;
        const folder = await this.getFolder(client, folderPath);
        if (folder) {
            try {
                const address = `/me/drive/items/${folder.id}`;
                log.info(`★★★ requesting api(delete): [${address}] ★★★`);
                await client.api(address).delete();    
            } catch (e) {
                res = false;
                log.error(`deleteFolder failed: [${e}]`);
            }
        } else {
            // log.info(`★★★ folder [${folderPath}] not found ★★★`);
        }

        // log.info(`▲▲▲ deleteFolder END folderName: [${folderPath}] ▲▲▲`);
        return res;
    }

    /**
     * 指定したフォルダ内のアイテムをバックアップします※もとのファイルをバックアップフォルダに移動します。もとのファイルはなくなってしまうのでご注意ください。
     * @param client MicrosoftGraphクライアント
     * @param folderPath バックアップしたいフォルダのパス(パス区切り文字は「/」)※パスを省略するとrootを取得します。またパスを指定する際はrootは省略可能です。
     * @param backupFolder バックアップフォルダ名※バックアップフォルダはバックアップ対象フォルダに作成します。
     */
    public async backupFolder(client: Client, folderPath: string, backupFolder: string, chooser?: FileChooser): Promise<boolean> {
        // log.info(`▼▼▼ backupFolder START folderName: [${folderPath}] backupFolder: [${backupFolder}] ▼▼▼`);
        let res = true;
        const target = chooser?? new FileChooser;
        const items = await this.getItems(client, folderPath, true, target);
        // バックアップフォルダ（存在しなければ作成）の中のアイテムをすべて削除
        const bkPath = `${folderPath}/${backupFolder}`
        const oldBk = await this.getItems(client, bkPath, true)
        await Promise.all(oldBk.map(async (item) => {
            item.name && this.deleteFile(client, item.name, bkPath)
        }))

        // バックアップ対象フォルダ内の全アイテムをバックアップフォルダに移動
        const bkFolder = await this.getFolder(client, bkPath, true);
        bkFolder && await Promise.all(items.map(async (item) => {
            // ただしバックアップフォルダ自体は移動してはいけない
            if (item.id != bkFolder.id) {
                const driveItem = {
                    parentReference: {
                        id: bkFolder.id
                    },
                    name: item.name,
                }
                try {
                    const address = `/me/drive/items/${item.id}`;
                    log.info(`★★★ requesting api(update): [${address}] ★★★`);
                    await client.api(address).update(driveItem);                    
                } catch (e) {
                    res = false;
                    log.error(`backupFolder failed to move file [${item.name}]: [${e}]`);    
                }
            }
        }))

        // log.info(`▲▲▲ backupFolder END folderName: [${folderPath}] backupFolder: [${backupFolder}] ▲▲▲`);
        return res;
    }

    /**
     * バックアップフォルダ内のファイルからフォルダ内を復旧します。バックアップフォルダは空になります。
     * @param client MicrosoftGraphクライアント
     * @param folderPath リストアしたいフォルダのパス(パス区切り文字は「/」)※パスを省略するとrootを取得します。またパスを指定する際はrootは省略可能です。
     * @param backupFolder バックアップフォルダ名※バックアップフォルダはリストア対象フォルダ内にあることが前提です。存在しなければ作成します。
     */
    public async restoreFromBackup(client: Client, folderPath: string, backupFolder: string, chooser?: FileChooser): Promise<boolean> {
        let res = true;
        const target = chooser?? new FileChooser;
        // バックアップフォルダ（存在しなければ作成）にあるアイテム
        const bkPath = `${folderPath}/${backupFolder}`
        const bkFolder = await this.getFolder(client, bkPath, true);
        const oldBk = await this.getItems(client, bkPath, true)
        oldBk.forEach(rec => {
            rec.name && target.add(rec.name, false);
        })

        if (bkFolder) {
            // リストア対象フォルダ内のアイテムはすべて削除
            const items = await this.getItems(client, folderPath, true);
            await Promise.all(items.map(async(item) => {
                if (item.id !== bkFolder.id && target.test(item.name)) {
                    item.name && await this.deleteFile(client, item.name, folderPath);
                }
            }))
            // バックアップフォルダ内の全ファイルをリストア対象フォルダに移動
            const folder = await this.getFolder(client, folderPath, true);
            folder && await Promise.all(oldBk.map(async (item) => {
                const driveItem = {
                    parentReference: {
                        id: folder.id
                    },
                    name: item.name,
                }
                try {
                    const address = `/me/drive/items/${item.id}`;
                    log.info(`★★★ requesting api(update): [${address}] ★★★`);
                    await client.api(address).update(driveItem);                    
                } catch (e) {
                    res = false;
                    log.error(`restoreFromBackup failed to move file [${item.name}]: [${e}]`);    
                }
            }))    
        }
        return res;
    }

    /**
     * 親フォルダ内から、指定した名前のフォルダを取得します。
     * @param client MicrosoftGraphクライアント
     * @param find フォルダ名
     * @param parent 親フォルダまでのパス(パス区切り文字は「/」)※パスを省略するとrootを取得します。またパスを指定する際はrootは省略可能です。
     * @param createFolderIfNotExist trueを指定するとパス内に存在しないフォルダがある場合にフォルダを作成します。
     */
    private async findFolder(client: Client, find: string, parent?: DriveItem | null | undefined, createFolderIfNotExist?: boolean): Promise<DriveItem | null> {
        // log.info(`▼▼▼ findFolder START parent: [${parent? parent.name : "root"}] find: [${find}] createFolderIfNotExist: [${createFolderIfNotExist}] ▼▼▼`);
        let res: DriveItem | null = null;

        const children: DriveItem[] = [];
        let parentId: string | null = null;
        if (parent) {
            if (parent.id) {
                parentId = parent.id;
                const address = `/me/drive/items/${parent.id}/children`;
                log.info(`★★★ requesting api(get): [${address}] ★★★`);
                const api = await client.api(address).get()
                const fetched = await getAllPages<DriveItem>(client, api);
                fetched.forEach(rec => {children.push(rec)});
                // log.info(`★★★ Found [${children.length}] items in the folder ID: [${parent.id}] ★★★`);
            }
        } else {
            const address = `/me/drive/root/children`;
            log.info(`★★★ requesting api(get): [${address}] ★★★`);
            const api = await client.api(address).get();
            const fetched = await getAllPages<DriveItem>(client, api);
            fetched.forEach(rec => {children.push(rec)});
            // log.info(`★★★ Found [${children.length}] items in the root folder ★★★`);
        }

        for (let i = 0; i < children.length; i++) {
            const item = children[i];
            if (item.folder && item.name === find) {
                res = item;
                // log.info(`★★★ Folder [${res.name}] found... ID: [${res.id}] ★★★`);
                break;
            }
        }

        if (!res && (createFolderIfNotExist?? false)) {
            const driveItem = {
                name: find,
                folder: { },
            };
            if (parent) {
                const address = `/me/drive/items/${parentId}/children`;
                log.info(`★★★ requesting api(post): [${address}] ★★★`);
                res = await client.api(address).post(driveItem);
                if (res) {
                    // log.info(`★★★ Folder [${res.name}] created... ID: [${res.id}] ★★★`);
                }
            } else {
                const address = `/me/drive/root/children`;
                log.info(`★★★ requesting api(post): [${address}] ★★★`);
                res = await client.api(address).post(driveItem);
                if (res) {
                    // log.info(`★★★ Folder [${res.name}] created... ID: [${res.id}] ★★★`);
                }
            }
        }
        // log.info(`▲▲▲ findFolder END parent: [${parent? parent.name : "root"}] find: [${find}] createFolderIfNotExist: [${createFolderIfNotExist}] ▲▲▲`);

        return res;
    }
}

/**
 * OneDrive上のフォルダやファイルを操作するユーティリティです。
 */
export const FileUtil = new Util();