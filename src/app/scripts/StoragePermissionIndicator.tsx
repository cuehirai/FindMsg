/* eslint-disable react/prop-types */
import * as React from 'react';
import { storage, delay } from './utils';
import { AcceptIcon, BanIcon } from "./ui";

export const StoragePermissionIndicator: React.FunctionComponent<{ loading: boolean }> = ({ loading }) => {
    if (loading) return null;

    const [granted, setGranted] = React.useState(storage.granted);

    if (!granted) {
        delay(10000).then(() => setGranted(storage.granted()));
        return <BanIcon title="Persistent storage permission denied" />;
    }

    return <AcceptIcon title="Persistent storage permission granted" />;
};

