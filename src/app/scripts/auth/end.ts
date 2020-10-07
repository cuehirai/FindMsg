import * as teams from "@microsoft/teams-js";
import { isUnknownObject } from '../utils';
import { client } from './client';
import { AI } from '../appInsights';

async function processResponse() {
    try {
        const response = await client.handleRedirectPromise();

        if (response === null) {
            teams.authentication.notifyFailure("no redirect resonse");
            AI.trackEvent({ name: "AuthResponseNull" });
        } else {
            teams.authentication.notifySuccess(JSON.stringify(response));
        }
    } catch (error) {
        AI.trackException({ error });
        if (isUnknownObject(error) && 'name' in error && 'message' in error) {
            teams.authentication.notifyFailure(`${error.name}: ${error.message}`);
        } else {
            teams.authentication.notifyFailure("Unexpected error");
        }
    }
}

teams.initialize(processResponse);
