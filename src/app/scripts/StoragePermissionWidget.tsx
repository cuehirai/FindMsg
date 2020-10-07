/* eslint-disable react/prop-types */
import * as React from 'react';
import { storage, delay } from './utils';
import { Alert, Link, Text } from "./ui";
import { error, info } from './logger';


export interface IStoragePermissionWidgetTranslation {
    grantTitle: string;
    grantMessage: string;
    linkInside: string;
    linkOutside: string;
}


export interface StoragePermissionWidgetProps {
    t: IStoragePermissionWidgetTranslation,
    granted: () => void,
}


const ask = (granted: () => void, t: IStoragePermissionWidgetTranslation): () => Promise<void> => storage.needNewWindow ?
    async () => {
        info("try to open window");
        const u = new URL(window.location.origin + "/storage.html");
        u.searchParams.append("t", t.grantTitle);
        u.searchParams.append("m", t.grantMessage);
        u.searchParams.append("b", t.linkInside);

        const w = window.open(u.toString(), "PermissionRequestWindow");

        if (w) {
            while (!w.closed) {
                await delay(500);
                const persisted = await navigator.storage.persisted();
                if (persisted) {
                    granted();
                    w.close();
                    break;
                }
            }
        } else {
            error("Window.open() failed");
        }
    }
    :
    async () => {
        info("try to ask for permission");
        if (await navigator.storage.persist()) {
            granted();
        }
    };


/**
 * Display controls and status of the sync process
 * @param props
 */
export const StoragePermissionWidget: React.FunctionComponent<StoragePermissionWidgetProps> = ({ t, granted }) => <Alert>
    <Text content={t.grantMessage} />
    {" "}
    <Link onClick={ask(granted, t)}>{storage.needNewWindow ? t.linkOutside : t.linkInside}</Link>
</Alert>;
