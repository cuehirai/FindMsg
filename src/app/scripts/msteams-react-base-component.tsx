// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.

// Heavily modified by Christian Giessner

import * as React from "react";
import { render } from "react-dom";
import { teamsTheme, teamsDarkTheme, teamsHighContrastTheme, ThemePrepared } from "./ui";
import * as microsoftTeams from "@microsoft/teams-js";

/** State interface for the Teams Base user interface React component */
export interface ITeamsBaseComponentState {
    /** The Microsoft Teams theme style (Light, Dark, HighContrast) */
    theme: ThemePrepared<Record<string, unknown>>;
}

/** Base implementation of the React based interface for the Microsoft Teams app */
export default class TeamsBaseComponent<P, S extends ITeamsBaseComponentState> extends React.Component<P, S> {

    constructor(props: P) {
        super(props);
        this.clearMsalStorage();
    }


    /**
     * Static method to render the component
     * @param element DOM element to render the control in
     * @param props Properties
     */
    // eslint-disable-next-line react/require-render-return
    public static render<P>(element: HTMLElement, props: P): void {
        render(React.createElement(this, props), element);
    }

    /**
     * Returns true if hosted in Teams
     * @param timeout timeout in milliseconds, default = 1000
     * @returns a `Promise<boolean>`
     */
    protected inTeams = (timeout = 2000): Promise<boolean> => {
        return new Promise((resolve, reject) => {
            try {
                microsoftTeams.initialize(() => {
                    resolve(true);
                });
                setTimeout(() => {
                    resolve(false);
                }, timeout);
            } catch (e) {
                reject(e);
            }
        });
    }


    protected getTheme = (themeStr?: string): ThemePrepared<Record<string, unknown>> => {
        switch (themeStr) {
            case "dark": return teamsDarkTheme;
            case "contrast": return teamsHighContrastTheme;
            default: return teamsTheme;
        }
    }

    /** Updates the theme */
    protected updateTheme = (themeStr?: string): void => {
        this.setState({ theme: this.getTheme(themeStr) });
    }


    /** Returns the value of a query variable */
    protected getQueryVariable = (variable: string): string | undefined => {
        const query = window.location.search.substring(1);
        const vars = query.split("&");
        for (const varPairs of vars) {
            const pair = varPairs.split("=");
            if (decodeURIComponent(pair[0]) === variable) {
                return decodeURIComponent(pair[1]);
            }
        }
        return undefined;
    }


    /**
     * Clear MSAL entries in localstorage
     * Try to get rid of "interaction_in_progress" errors
     */
    private clearMsalStorage() {
        let index = localStorage.length;
        while (--index >= 0) {
            const key = localStorage.key(index);
            if (key?.startsWith("msal.")) {
                localStorage.removeItem(key);
            }
        }
    }
}
