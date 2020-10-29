import * as React from "react";

import TeamsBaseComponent, { ITeamsBaseComponentState } from "../msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

import { ComponentEventHandler, Alert, AlertProps, Button, Dialog, Divider, Dropdown, DropdownItemProps, Flex, Input, InputProps, Provider, Segment, Text, Page, DropdownProps } from "../ui";

import { Client, AuthProviderCallback } from "@microsoft/microsoft-graph-client";
import { getAuthTokenSilent, loginPopup, AuthError, haveUserInfo } from "../auth/auth";
import * as log from '../logger';
import { assert1, cancellation, progressFn, nop, OperationCancelled } from "../utils";
import { Sync, IFindMsgChannelMessage, FindMsgChannelMessage, Direction, MessageOrder, FindMsgTeam, FindMsgChannel } from '../db';
import { MessageTable } from "./MessageTable";
import { SyncWidget, SyncControl, SyncState } from "../SyncWidget";
import { invalidDate } from "../dateUtils";
import * as strings from '../i18n/messages';
import { IMessageTranslation } from "../i18n/IMessageTranslation";
import { AI } from '../appInsights';
import { getTopLevelMessagesLastSynced } from "../db/Sync";


declare type DropdownItemPropsKey = DropdownItemProps & { key: string };

const initialDisplayCount = 50;
const loadMoreCount = 25;

export interface IFindMsgTopicsTabState extends ITeamsBaseComponentState, SyncState, SyncControl {
    /** If the page is currently loading */
    loading: boolean;

    error: string;
    warning: string;

    searchResult: ISearchResult;
    teamsInfo: ITeamsInfo;
    filterInput: string;
    filterString: string;

    dropdownDisabled: boolean;
    teamIdx: number;
    channelIdx: number;

    /**
     * Auth is called from the internals of graph client.
     * This callback allows the waiting graph client to continue.
     * This is stored as a way to allow a "second chance" login,
     * when automatic login fails.
     */
    authInProgress: AuthProviderCallback;
    authResult: AuthError | null;
    loginRequired: boolean;

    t: IMessageTranslation;
}


interface ITeamsInfo {
    locale: string | null;
    groupId: string;
    channelId: string;
    loginHint: string;
    teamName: string;
    channelName: string;
    teamOptions: DropdownItemPropsKey[];
    channelOptions: DropdownItemPropsKey[][];
}


interface ISearchResult {
    messages: IFindMsgChannelMessage[];
    hasMore: boolean;
    order: MessageOrder;
    dir: Direction;
}


export interface ITopicsTabTranslation {
    pageTitle: string;
    // team: string;
    // channel: string;
    // loadMore: string;
    // allTeams: string;
    // allChannels: string;
}


export class FindMsgTopicsTab extends TeamsBaseComponent<never, IFindMsgTopicsTabState> {

    constructor(props: never) {
        super(props);

        const cid = this.getQueryVariable("cid");
        const gid = this.getQueryVariable("gid");
        const l = this.getQueryVariable("l");

        this.state = {
            loading: true,
            syncing: false,
            syncStatus: "",
            syncCancel: nop,
            syncCancelled: false,
            lastSynced: invalidDate(),
            error: "",
            warning: "",

            authInProgress: nop,
            authResult: null,
            loginRequired: false,

            searchResult: {
                messages: [],
                hasMore: false,
                order: MessageOrder.touched,
                dir: Direction.descending,
            },

            filterInput: "",
            filterString: "",

            dropdownDisabled: !!(cid && gid),
            teamIdx: 0,
            channelIdx: 0,

            teamsInfo: {
                channelId: cid ?? "",
                groupId: gid ?? "",
                locale: l ?? null,
                loginHint: this.getQueryVariable("hint") ?? "",
                channelName: "",
                teamName: "",
                channelOptions: [[]],
                teamOptions: [],
            },

            t: strings.get(l),
            theme: this.getTheme(this.getQueryVariable("theme")),
        }
        this.msGraphClient = Client.init({ authProvider: this.authProvider });
    }


    private msGraphClient: Client;
    private getContext = (): Promise<microsoftTeams.Context> => new Promise(microsoftTeams.getContext);
    private filterTimeout = 0;


    public async componentDidMount(): Promise<void> {
        this.updateTheme(this.getQueryVariable("theme"));

        if (await this.inTeams(2000)) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);

            const context = await this.getContext();
            log.info("context", context);
            this.updateTheme(context.theme);
            microsoftTeams.appInitialization.notifySuccess();
            this.initInfo(context);
        } else {
            this.initInfo();
        }
    }


    private initInfo = async (context?: microsoftTeams.Context): Promise<void> => {
        const loginHint = context?.loginHint ?? this.state.teamsInfo.loginHint;
        let groupId = context?.groupId ?? this.state.teamsInfo.groupId;
        let channelId = context?.channelId ?? this.state.teamsInfo.channelId;
        const locale = context?.locale ?? this.state.teamsInfo.locale;
        const teamName = context?.teamName ?? "";
        const channelName = context?.channelName ?? "";
        const t = strings.get(locale);

        document.title = t.topics.pageTitle;

        microsoftTeams.setFrameContext({
            contentUrl: location.href,
            websiteUrl: location.href,
        });

        // add lastSynced for top level messages
        const lastSynced = channelId ? await Sync.getChannelLastSynced(channelId) : getTopLevelMessagesLastSynced();

        let teamIdx: number;
        let channelIdx: number;
        let teamOptions: DropdownItemPropsKey[];
        let channelOptions: DropdownItemPropsKey[][];

        if (this.state.dropdownDisabled) {
            // We are in a channel. teamId, channelId, teamName, channelName are provided by teams and can not be changed.
            teamIdx = 0;
            channelIdx = 0;

            teamOptions = [{
                header: teamName || (await FindMsgTeam.get(groupId))?.displayName || "(unknown)",
                key: groupId,
                selected: true,
            }];

            channelOptions = [[{
                header: channelName || (await FindMsgChannel.get(channelId))?.displayName || "(unknown)",
                key: channelId,
                selected: true,
            }]];

        } else {
            // We are standalone. Get all the data from the local store.
            const allChannels = await FindMsgChannel.getAll();
            const teams = (await FindMsgTeam.getAll()).map(t => ({ ...t, channels: allChannels.filter(c => c.teamId === t.id) }));

            const allTeamsOption: DropdownItemPropsKey = {
                header: t.common.allTeams,
                key: "",
                selected: false,
            };

            const allChannelsOption: DropdownItemPropsKey = {
                header: t.common.allChannels,
                key: "",
                selected: false,
            };

            teamOptions = teams.map(t => ({
                header: t.displayName,
                key: t.id,
                selected: false,
            }));

            // find the currently selected index including the the "All Teams" entry we add below
            teamIdx = 1 + teamOptions.findIndex(t => t.key === groupId, 0);
            teamOptions.unshift(allTeamsOption);
            teamOptions[teamIdx].selected = true;

            // group channels by team and construct dropdown options
            channelOptions = teams.map(t => allChannels.reduce<DropdownItemPropsKey[]>((filtered, { displayName, id, teamId }) => {
                if (teamId === t.id) filtered.push({
                    header: displayName,
                    key: id,
                    selected: id === channelId,
                });
                return filtered;
            }, []));
            channelOptions.forEach(element => {
                element.unshift(allChannelsOption)
            });

            // add the "All Channels" options as the only element of "All Teams"
            channelOptions.unshift([allChannelsOption]);

            channelIdx = Math.max(0, channelOptions[teamIdx].findIndex(c => c.key === channelId));
            channelOptions[teamIdx][channelIdx].selected = true;

            groupId = teamOptions[teamIdx].key;
            channelId = channelOptions[teamIdx][channelIdx].key;
        }

        this.setState({
            teamsInfo: {
                groupId, channelId,
                teamOptions, channelOptions,
                teamName, channelName,
                loginHint, locale,
            },
            teamIdx, channelIdx,
            lastSynced,
            loginRequired: !haveUserInfo(loginHint),
            t,
        }, this.getMessages);
    }


    private onFilterChanged: ComponentEventHandler<InputProps & { value: string; }> = (_: unknown, data): void => {
        this.setState({ filterInput: data?.value ?? "" }, () => {
            window.clearTimeout(this.filterTimeout);
            this.filterTimeout = window.setTimeout(() => {
                const { order, dir } = this.state.searchResult;
                this.getMessages(order, dir);
            }, 250);
        });
    }


    private onTeamChanged = (_: unknown, data: DropdownProps): void => {
        const selected = data.value as DropdownItemPropsKey;
        const { teamsInfo } = this.state;
        const { teamOptions, channelOptions } = teamsInfo;

        const newIdx = teamOptions.findIndex(t => t.key === selected.key);
        const newOpts = teamOptions.map((to, i) => ({ ...to, selected: i === newIdx }));

        const newChannelIdx = Math.max(0, channelOptions[newIdx].findIndex(co => co.selected));
        channelOptions[newIdx][newChannelIdx].selected = true;

        this.setState({
            teamIdx: newIdx,
            // if no channel is selected, select the 2nd entry (first channel entry), but fall back to the 1st entry ("All Channels")
            channelIdx: newChannelIdx,
            teamsInfo: {
                ...teamsInfo,
                teamOptions: newOpts,
                channelId: channelOptions[newIdx][newChannelIdx].key,
                groupId: teamOptions[newIdx].key,
            }
        }, this.getMessages);
    }


    private onChannelChanged = (_: unknown, data: DropdownProps): void => {
        const selected = data.value as DropdownItemPropsKey;
        const { teamIdx, channelIdx, teamsInfo } = this.state;
        const { channelOptions } = teamsInfo;

        const newIdx = channelOptions[teamIdx].findIndex(t => t.key === selected.key);
        channelOptions[teamIdx] = channelOptions[teamIdx].map((co, i) => (i === channelIdx || i === newIdx) ? { ...co, selected: i === newIdx } : co);

        this.setState({
            channelIdx: newIdx,
            teamsInfo: {
                ...teamsInfo,
                channelId: channelOptions[teamIdx][newIdx].key,
            }
        }, this.getMessages);
    }


    public render(): JSX.Element {
        const {
            theme,
            loading,
            error,
            warning,

            syncing, syncStatus, syncCancelled, lastSynced,

            authResult,
            loginRequired,

            filterInput,
            filterString,
            teamIdx,
            channelIdx,
            dropdownDisabled,

            teamsInfo: {
                teamOptions, channelOptions
            },

            searchResult: { messages, order, dir, hasMore },
            t: {
                dateTimeFormat, common, sync,
                table, footer, auth, filter,
                unknownUserDisplayName,
            }
        } = this.state;

        return (
            <Provider theme={theme}>
                <Page>
                    <Dialog
                        open={loginRequired && authResult === null}
                        header={auth.loginDialogHeader}
                        confirmButton={auth.loginButtonText}
                        onConfirm={this.login}
                        content={auth.loginMessage}
                    />

                    <Dialog
                        open={authResult !== null}
                        header={authResult?.isRecoverable ? auth.loginDialogHeader : auth.unkownError}
                        confirmButton={authResult?.isRecoverable ? (authResult.adminConsentRequired ? auth.adminLoginButtonText : auth.loginButtonText) : undefined}
                        onConfirm={this.login}
                        content={authResult?.message}
                    />

                    <Segment>
                        <Flex gap="gap.large">
                            <Flex.Item shrink={2}>
                                <Flex gap="gap.small" wrap>
                                    <Dropdown disabled={dropdownDisabled} items={teamOptions} value={teamOptions[teamIdx]} onChange={this.onTeamChanged} />
                                    <Dropdown disabled={dropdownDisabled} items={channelOptions[teamIdx]} value={channelOptions[teamIdx][channelIdx]} onChange={this.onChannelChanged} />
                                </Flex>
                            </Flex.Item>

                            <Flex.Item grow shrink>
                                <Flex gap="gap.small">
                                    <Input
                                        type="text"
                                        label={filter}
                                        labelPosition="inline"
                                        value={filterInput}
                                        onChange={this.onFilterChanged}
                                    />

                                    <Flex.Item grow>
                                        <div />
                                    </Flex.Item>

                                    <Flex.Item align="start">
                                        <SyncWidget
                                            t={sync}
                                            syncStart={this.syncMessages}
                                            syncCancel={this.cancelSync}
                                            syncCancelled={syncCancelled}
                                            syncStatus={syncStatus}
                                            syncing={syncing}
                                            lastSynced={lastSynced}
                                            loading={loading}
                                        />
                                    </Flex.Item>
                                </Flex>
                            </Flex.Item>
                        </Flex>
                    </Segment>

                    {error && <Alert
                        content={error}
                        dismissible
                        variables={{ urgent: true }}
                        onVisibleChange={this.errorVisibilityChanged}
                    />}

                    {warning && <Alert
                        content={warning}
                        dismissible
                        // variables={{ urgent: true }}
                        onVisibleChange={this.warningVisibilityChanged}
                    />}

                    <MessageTable t={table} dateFormat={dateTimeFormat} messages={messages} dir={dir} order={order} sort={this.getMessages} loading={loading} filter={filterString} unknownUserDisplayName={unknownUserDisplayName} />

                    {hasMore && <Button onClick={this.loadMoreMessages} content={common.loadMore} />}

                    <div style={{ flex: 1 }} />
                    <Divider />
                    <Text size="smaller" content={footer} />
                </Page>
            </Provider >
        );
    }


    private setError = (error: unknown, userMessage: string): void => {
        this.setState({ error: userMessage });
        log.error(error);
    }

    private errorVisibilityChanged: ComponentEventHandler<AlertProps> = (_, data) => !data?.visible && this.setState({ error: "" });
    private warningVisibilityChanged: ComponentEventHandler<AlertProps> = (_, data) => !data?.visible && this.setState({ warning: "" });


    private authProvider = async (done: AuthProviderCallback) => {
        try {
            const token = await getAuthTokenSilent(this.state.t.auth, this.state.teamsInfo.loginHint);
            this.setState({ authInProgress: nop, authResult: null, loginRequired: false });
            done(null, token);
        } catch (error) {
            AI.trackException({ exception: error });
            log.error(error);
            this.setState({ authInProgress: done, authResult: error });
        }
    }


    private login = async () => {
        const {
            authInProgress: done,
            authResult,
            teamsInfo: { loginHint },
            t: { auth }
        } = this.state;

        try {
            const result = await loginPopup(auth, loginHint, authResult?.adminConsentRequired);
            this.setState({ authInProgress: nop, authResult: null, loginRequired: false });
            done(null, result.accessToken);
        } catch (error) {
            AI.trackException({ exception: error });
            log.error(error);
            this.setState({ authInProgress: nop, authResult: error });
            done(error, null);
        }
    };


    /**
     * Report sync status callback
     * @param syncStatus
     */
    private reportProgress: progressFn = syncStatus => this.setState({ syncStatus });


    /**
     * Cancel an ongoing sync
     */
    private cancelSync = () => {
        this.state.syncCancel();
        this.setState({ syncCancelled: true });
    };


    /**
     * Sync new messages from microsoft graph
     */
    private syncMessages = async () => {
        let { lastSynced } = this.state;

        try {
            const {
                teamsInfo: { groupId, channelId },
                t: { syncProgress }
            } = this.state;
            const [cancel, throwfn] = cancellation();

            this.setState({ syncing: true, syncCancel: cancel, syncCancelled: false, error: "", warning: "" });
            let syncResult: boolean;
            if (groupId && channelId) {
                syncResult = await Sync.channelTopLevelMessages(this.msGraphClient, groupId, channelId, throwfn, this.reportProgress, syncProgress);
            } else {
                syncResult = await Sync.autoSyncAll(this.msGraphClient, false, throwfn, this.reportProgress, syncProgress);
            }

            if (syncResult) {
                lastSynced = (await Sync.getChannelLastSynced(channelId));
            } else {
                AI.trackEvent({ name: "syncProblem" });
                this.setState({ warning: syncProgress.syncProblem });
            }

            await this.initInfo();
        } catch (error) {
            if (error instanceof OperationCancelled) {
                log.info("sync cancelled");
            } else {
                AI.trackException({ exception: error });
                this.setError(error, this.state.t.error.syncFailed);
            }
        } finally {
            this.setState({ syncing: false, lastSynced });
        }
    }


    private getMessages = async (order: MessageOrder = MessageOrder.touched, dir: Direction = Direction.descending): Promise<void> => {
        const {
            teamsInfo: { channelId, channelOptions },
            filterInput,
        } = this.state;

        this.setState({ loading: true });

        try {
            const { teamIdx, channelIdx } = this.state;
            const channelIds = new Set<string>();
            log.info(`teamIdx= ${teamIdx}`);
            log.info(`channelIdx= ${channelIdx}`);
            if (teamIdx > 0) {
                if (channelIdx == 0) {
                    channelOptions[teamIdx].forEach(c => {
                        log.info(`key= ${c.key} name= ${c.header}`);
                        channelIds.add(c.key);
                    });
                }
            }

            assert1(channelId, nameof(channelId));
            const [messages, hasMore] = await FindMsgChannelMessage.getTopLevelMessagesWithSubject(channelId, channelIds, order, dir, 0, initialDisplayCount, filterInput);
            this.setState({
                filterString: filterInput,
                searchResult: { hasMore, messages, dir, order }
            });
        } catch (error) {
            AI.trackException({ exception: error });
            this.setError(error, this.state.t.error.indexedDbReadFailed);
        } finally {
            this.setState({ loading: false });
        }
    };


    private loadMoreMessages = async () => {
        const {
            searchResult: { messages, order, dir },
            teamsInfo: { channelId, channelOptions },
            filterInput,
        } = this.state;

        try {
            const { teamIdx, channelIdx } = this.state;
            const channelIds = new Set<string>();
            log.info(`teamIdx= ${teamIdx}`);
            log.info(`channelIdx= ${channelIdx}`);
            if (teamIdx > 0) {
                if (channelIdx == 0) {
                    channelOptions[teamIdx].forEach(c => {
                        log.info(`key= ${c.key} name= ${c.header}`);
                        channelIds.add(c.key);
                    });
                }
            }
            assert1(channelId);
            this.setState({ loading: true });

            const [newMessages, hasMore] = await FindMsgChannelMessage.getTopLevelMessagesWithSubject(channelId, channelIds, order, dir, messages.length, loadMoreCount, filterInput);

            this.setState({
                searchResult: {
                    messages: [...messages, ...newMessages],
                    hasMore, order, dir
                }
            });
        } catch (error) {
            AI.trackException({ exception: error });
            this.setError(error, this.state.t.error.indexedDbReadFailed);
        } finally {
            this.setState({ loading: false });
        }
    }
}
