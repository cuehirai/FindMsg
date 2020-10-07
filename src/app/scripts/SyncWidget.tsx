import * as React from 'react';
import * as du from "./dateUtils";
import { cancelFn } from './utils';
import { Button, Loader, Text, Flex } from "./ui";


export interface SyncState {
    /** Last time the channel messages were synced */
    lastSynced: Date,

    /** If sync if currently in progress */
    syncing: boolean,

    /** If cancellation was requested */
    syncCancelled: boolean,

    /** Latest message from the sync process */
    syncStatus: string,
}


export interface SyncControl {
    /** Call to request cancellation of ongoing sync */
    syncCancel: cancelFn;

    /**  Call to request sync to start now */
    syncStart?: () => void;
}


export interface ISyncWidgetTranslation {
    lastSynced: (d: Date) => string;
    syncNowButton: string;
    syncing: string;
    cancel: string;
    cancelWait: string;
}


export interface SyncWidgetProps extends SyncState, Required<SyncControl> {
    t: ISyncWidgetTranslation
    loading: boolean;
}


/**
 * Display controls and status of the sync process
 * @param props
 */
export const SyncWidget: React.FunctionComponent<SyncWidgetProps> = ({ lastSynced, syncing, syncCancelled, syncCancel, syncStart, syncStatus, t, loading }) => {
    // only display the sync now button if more than 5 minutes since the last sync
    const cutoff = du.subMinutes(new Date, 5);
    const [displaySyncNowButton, setDisplaySyncNowButton] = React.useState(!du.isValid(lastSynced) || du.isBefore(lastSynced, cutoff));

    /*
    Widget states:

    - idle, recently synced:
        { Last sync: YMDHM }

    - idle:
        { Last sync: YMDHM [sync now] }

    - syncing:
        { <Spinner> syncing [cancel] }
        { statusMessage              }

    - syncing, cancel requested:
        { <Spinner> cancelling }
    */

    if (syncing) {
        if (displaySyncNowButton) setDisplaySyncNowButton(false);
        if (syncCancelled) {
            return <Loader labelPosition="start" label={<Text>{syncStatus}</Text>} />
        } else {
            return <Loader labelPosition="start" label={
                <Flex column hAlign="end">
                    <Flex vAlign="center">
                        <Text>{t.syncing}</Text>
                        <Button size="smallest" onClick={syncCancel} content={t.cancel} />
                    </Flex>
                    <Text>{syncStatus}</Text>
                </Flex>
            } />
        }
    } else {
        if (displaySyncNowButton) {
            return <Flex vAlign="center">
                <Text>{t.lastSynced(lastSynced)}</Text>
                <Button onClick={syncStart} content={t.syncNowButton} disabled={loading} />
            </Flex>
        } else {
            setTimeout(() => setDisplaySyncNowButton(true), du.differenceInMilliseconds(lastSynced, cutoff));
            return <Text>{t.lastSynced(lastSynced)}</Text>;
        }
    }
};
