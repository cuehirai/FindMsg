import * as React from "react";
import * as microsoftChats from "@microsoft/teams-js";
import { getAuthTokenSilent, loginPopup, AuthError, haveUserInfo, userId } from "../auth/auth";
import * as log from '../logger';
import ChatsBaseComponent, { ITeamsBaseComponentState } from "../msteams-react-base-component";

import { Alert, AlertProps, Button, DatePicker, Dialog, Divider, Dropdown, DropdownProps, DropdownItemProps, Flex, Header, Input, InputProps, Loader, Provider, RadioGroup, RadioGroupItemProps, Segment, Text, ShorthandCollection, ComponentEventHandler } from "../ui";

import { Client, AuthProviderCallback } from "@microsoft/microsoft-graph-client";
import { cancellation, nop, progressFn, OperationCancelled, cancelFn, cancellationNoThrow, assertT, storage } from "../utils";
import { Sync, FindMsgChat, IFindMsgChatMessage, FindMsgChatMessage, FindMsgUserCache, IFindMsgChat } from "../db";
import { ChatSelect } from "./ChatMemberSelect";
import { SearchResultView } from "./SearchResult";
import * as du from "../dateUtils";
import { SyncState, SyncControl, SyncWidget } from "../SyncWidget";
import * as strings from '../i18n/messages';
import { IFindMsgTranslation } from "../i18n/IFindMsgTranslation";
import { Page } from '../ui';
import { StoragePermissionWidget } from "../StoragePermissionWidget";
import { StoragePermissionIndicator } from "../StoragePermissionIndicator";
import { AI } from '../appInsights';
import { db, idx } from "../db/Database";
import Dexie from "dexie";


declare type SearchUserItem = DropdownItemProps & { key: string };
export declare type IFindMsgChatEx = IFindMsgChat & { singleCounterpart: string | null; }


interface IChatCache {
    chats: IFindMsgChatEx[];
}

enum DateRange { AllTime, PastWeek, PastMonth, PastYear, Custom }


interface ISearchInfo {
    searchChat: string;
    searchTime: DateRange;
    searchTimeFrom: Date;
    searchTimeTo: Date;
    searchUsers: SearchUserItem[];

    searchResults: IFindMsgChatMessage[];
    searching: boolean;
    searchCancel: cancelFn;
}


interface ITeamssInfo {
    entityId: string | null;
    subEntityId: string | null;
    loginHint: string;
}


export interface IFindMsgSearchChatState extends ITeamsBaseComponentState, IChatCache, ISearchInfo, SyncState, SyncControl {
    checkState: Map<string, boolean>;
    checkAll: boolean;
    searchUserOptions: ShorthandCollection<DropdownItemProps>;

    loading: boolean;

    error: string;
    warning: string;

    // function to cancel running sync (cooperative)
    cancel: (() => void) | null;

    authInProgress: AuthProviderCallback;
    authResult: AuthError | null;
    loginRequired: boolean;

    askForStoragePermission: boolean;

    teamsInfo: ITeamssInfo;
    t: IFindMsgTranslation;
}


export interface IChatSearchTabTranslation {
    pageTitle: string;
    header: string;
    search: string;
    searching: string;
    allChats: string;
    from: string;
    to: string;
    messagesFound: (shown: number, total: number) => string;
    cancel: string;

    searchTimeAll: string;
    searchTimePastWeek: string;
    searchTimePastMonth: string;
    searchTimePastYear: string;
    searchTimeCustom: string;

    searchUsersLabel: string;
    searchUsersPlaceholder: string;
}


const lastSyncedKey = "FindMsgSearchChat_last_synced";
const loadChatLastSynced = (): Date => du.parseISO(localStorage.getItem(lastSyncedKey) ?? "");
const storeChatLastSynced = (m: Date): void => localStorage.setItem(lastSyncedKey, du.formatISO(m));


interface DateRangeRadioGroupItemProps extends RadioGroupItemProps {
    value: DateRange;
}


export class FindMsgSearchChat extends ChatsBaseComponent<never, IFindMsgSearchChatState> {

    constructor(props: never) {
        super(props);

        this.state = {
            chats: [],

            error: "",
            warning: "",

            loading: true,
            cancel: null,
            checkState: new Map<string, boolean>(),
            checkAll: true,
            searchUserOptions: [],

            searchResults: [],
            searchChat: "",
            searchTime: DateRange.AllTime,
            searchTimeFrom: du.invalidDate(),
            searchTimeTo: du.invalidDate(),
            searchCancel: nop,
            searchUsers: [],
            searching: false,

            syncCancel: nop,
            syncCancelled: false,
            syncStatus: "",
            syncing: false,

            authInProgress: nop,
            authResult: null,
            loginRequired: false,

            // There is no easy way to determine a value for this based on the sync logic,
            // so do the expedient thing is to use the last time sync was executed in this tab.
            lastSynced: loadChatLastSynced(),

            teamsInfo: {
                entityId: this.getQueryVariable("eid") ?? "",
                subEntityId: this.getQueryVariable("sid") ?? "",
                loginHint: this.getQueryVariable("hint") ?? "",
            },

            askForStoragePermission: false,

            t: strings.get(this.getQueryVariable("l")),
            theme: this.getTheme(this.getQueryVariable("theme")),
        }

        this.msGraphClient = Client.init({ authProvider: this.authProvider });
    }


    public async componentDidMount(): Promise<void> {
        this.updateTheme(this.getQueryVariable("theme"));

        await this.getDataFromDb();
        let { t } = this.state;

        if (await this.inTeams()) {
            microsoftChats.initialize();
            microsoftChats.registerOnThemeChangeHandler(this.updateTheme);
            const context = await this.getContext();
            this.updateTheme(context.theme);
            microsoftChats.appInitialization.notifySuccess();

            microsoftChats.setFrameContext({
                contentUrl: location.href,
                websiteUrl: location.href,
            });

            t = strings.get(context.locale);

            this.setState({
                t,
                loginRequired: !haveUserInfo(context.loginHint),
                teamsInfo: {
                    entityId: context.entityId,
                    subEntityId: context.subEntityId ?? null,
                    loginHint: context.loginHint ?? "",
                }
            });
        } else {
            this.setState({
                loginRequired: !haveUserInfo(this.state.teamsInfo.loginHint),
                teamsInfo: {
                    entityId: "",
                    subEntityId: null,
                    loginHint: this.state.teamsInfo.loginHint,
                }
            });
        }

        document.title = t.chatSearch.pageTitle;

        this.setState({
            askForStoragePermission: !storage.granted() && storage.askForPermission,
            loading: false
        });
    }


    public render(): JSX.Element {
        const {
            t: {
                footer,
                filter,
                showCollapsed,
                showExpanded,
                sync,
                auth,
                storagePermission,
                unknownUserDisplayName,
                chatSearch: {
                    header,
                    search,
                    searching: searchingMsg,
                    cancel,
                    allChats,
                    from, to,
                    messagesFound,
                    searchUsersLabel,
                    searchUsersPlaceholder,
                },
            },
            askForStoragePermission,
            theme,
            chats,
            loading,
            checkState,
            checkAll,
            searching,
            searchChat,
            searchTime,
            searchTimeFrom,
            searchTimeTo,
            searchUserOptions,
            searchResults,
            searchCancel,
            lastSynced,
            syncing,
            syncCancelled,
            syncStatus,
            authResult,
            loginRequired,
            error,
            warning,
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

                    <Flex space="between">
                        <Header content={header} style={{ marginBlockStart: 0, marginBlockEnd: 0 }} />

                        <Flex.Item align="start">
                            <SyncWidget
                                t={sync}
                                syncStart={this.syncChatMessages}
                                syncCancel={this.cancelSync}
                                syncCancelled={syncCancelled}
                                syncStatus={syncStatus}
                                syncing={syncing}
                                loading={loading}
                                lastSynced={lastSynced} />
                        </Flex.Item>
                    </Flex>

                    {error && <Alert
                        content={error}
                        dismissible
                        variables={{ urgent: true }}
                        onVisibleChange={this.errorVisibilityChanged}
                    />}

                    {warning && <Alert
                        content={warning}
                        dismissible
                        onVisibleChange={this.warningVisibilityChanged}
                    />}

                    {askForStoragePermission && <StoragePermissionWidget granted={this.storagePermissionGranted} t={storagePermission} />}

                    <Segment>
                        <Flex gap="gap.small">
                            <Input onChange={this.searchChatChanged} onKeyDown={this.searchKeyDown} type="text" disabled={loading && !chats.length} value={searchChat} />
                            {!searching && <Button primary onClick={this.search} disabled={loading && !chats.length} content={search} />}
                            {searching && <Loader label={searchingMsg} labelPosition="end" delay={200} />}
                            {searching && <Button content={cancel} onClick={searchCancel} />}
                        </Flex>
                    </Segment>

                    <Segment>
                        <ChatSelect allText={allChats} all={checkAll} chats={chats} checkState={checkState} changed={this.chatCheckChanged} />
                    </Segment>

                    <Segment>
                        <Flex column gap="gap.small">
                            <RadioGroup checkedValue={searchTime} items={this.searchTimeOptions()} onCheckedValueChange={this.searchTimeOptionChanged} />
                            {searchTime === DateRange.Custom && <Flex gap="gap.medium">
                                <DatePicker label={from} value={searchTimeFrom} onSelectDate={this.searchTimeFromChanged} formatDate={this.formatDate} />
                                <DatePicker label={to} value={searchTimeTo} onSelectDate={this.searchTimeToChanged} formatDate={this.formatDate} />
                            </Flex>}
                        </Flex>
                    </Segment>

                    <Segment>
                        <Flex column>
                            <Text content={searchUsersLabel} />
                            <Dropdown
                                multiple clearable search
                                position="above"
                                placeholder={searchUsersPlaceholder}
                                items={searchUserOptions}
                                onChange={this.searchUserChanged}
                            />
                        </Flex>
                    </Segment>

                    <Divider />
                    <SearchResultView filter={filter} countFormat={messagesFound} m2dt={this.formatDate} messages={searchResults} searchChat={searchChat} showCollapsed={showCollapsed} showExpanded={showExpanded} unknownUserDisplayName={unknownUserDisplayName} />

                    <div style={{ flex: 1 }} />
                    <Divider />
                    <Flex space="between">
                        <Text size="smaller" content={footer} />
                        <StoragePermissionIndicator loading={loading} />
                    </Flex>
                </Page>
            </Provider>
        );
    }


    private searchUserChanged = (_e: unknown, data: DropdownProps) => {
        const values = data.value as SearchUserItem[];
        this.setState({ searchUsers: [...values] });
    };


    private errorVisibilityChanged: ComponentEventHandler<AlertProps> = (_e, data) => !data?.visible && this.setState({ error: "" });
    private warningVisibilityChanged: ComponentEventHandler<AlertProps> = (_e, data) => !data?.visible && this.setState({ warning: "" });


    private storagePermissionGranted = () => {
        this.setState({ askForStoragePermission: false });
    }


    private formatDate = (date?: Date | undefined): string => date ? du.format(date, this.state.t.dateFormat) : "";


    private msGraphClient: Client;


    private getDateFrom = (value: DateRange): Date => {
        switch (value) {
            case DateRange.Custom: return du.isValid(this.state.searchTimeFrom) ? this.state.searchTimeFrom : du.subMonths(du.startOfToday(), 1);
            case DateRange.PastWeek: return du.subWeeks(du.startOfToday(), 1);
            case DateRange.PastMonth: return du.subMonths(du.startOfToday(), 1);
            case DateRange.PastYear: return du.subYears(du.startOfToday(), 1);
            case DateRange.AllTime: return du.invalidDate();
            default: return du.invalidDate();
        }
    }


    private getDateTo = (value): Date => {
        switch (value) {
            case DateRange.Custom: return du.isValid(this.state.searchTimeTo) ? this.state.searchTimeTo : du.endOfToday();
            case DateRange.PastWeek: return du.endOfToday();
            case DateRange.PastMonth: return du.endOfToday();
            case DateRange.PastYear: return du.endOfToday();
            case DateRange.AllTime: return du.invalidDate();
            default: return du.invalidDate();
        }
    }


    /**
     * Creates the options for the search time selection RadioGroup
     * This is memoized and recreated only when locale changes.
     */
    private searchTimeOptions: () => ShorthandCollection<DateRangeRadioGroupItemProps> = (() => {
        let lastState: IChatSearchTabTranslation | null = null;
        let lastOptions: ShorthandCollection<DateRangeRadioGroupItemProps> = [];

        return () => {
            if (this.state.t.chatSearch === lastState) {
                return lastOptions;
            }

            const {
                searchTimeAll,
                searchTimePastWeek,
                searchTimePastMonth,
                searchTimePastYear,
                searchTimeCustom,
            } = lastState = this.state.t.chatSearch;

            return lastOptions = [
                {
                    key: "AllTime",
                    value: DateRange.AllTime,
                    label: searchTimeAll,
                },
                {
                    key: "PastWeek",
                    value: DateRange.PastWeek,
                    label: searchTimePastWeek,
                },
                {
                    key: "PastMonth",
                    value: DateRange.PastMonth,
                    label: searchTimePastMonth,
                },
                {
                    key: "PastYear",
                    value: DateRange.PastYear,
                    label: searchTimePastYear,
                },
                {
                    key: "Custom",
                    value: DateRange.Custom,
                    label: searchTimeCustom,
                },
            ];
        };
    })();


    private searchTimeOptionChanged: ComponentEventHandler<RadioGroupItemProps> = (_e, props) => {
        try {
            if (!props) {
                log.error("Invalid date range option");
                return;
            }

            const { value } = props as DateRangeRadioGroupItemProps;

            this.setState({
                searchTime: assertT(value, typeof DateRange.AllTime),
                searchTimeFrom: assertT(this.getDateFrom(value), du.isDate),
                searchTimeTo: assertT(this.getDateTo(value), du.isDate),
            });
        } catch (error) {
            AI.trackException({ exception: error });
            this.setError(error, this.state.t.error.internalError);
        }
    };


    private searchTimeFromChanged = (d: Date | undefined | null): void => {
        if (d) this.setState({ searchTimeFrom: du.startOfDay(d) })
    };


    private searchTimeToChanged = (d: Date | undefined | null): void => {
        if (d) this.setState({ searchTimeTo: du.endOfDay(d) })
    };


    private searchKeyDown = (event: React.KeyboardEvent<HTMLInputElement>): void => {
        if (event.key === "Enter") this.search();
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
    private syncChatMessages = async () => {
        let { lastSynced } = this.state;
        const { syncProgress } = this.state.t;

        try {
            const [cancel, checkCancel] = cancellation();

            this.setState({ syncing: true, syncCancel: cancel, syncCancelled: false, error: "", warning: "" });
            const result = await Sync.autoSyncChatAll(this.msGraphClient, checkCancel, this.reportProgress, syncProgress);
            if (result) {
                lastSynced = du.now();
                storeChatLastSynced(lastSynced);
            } else {
                AI.trackEvent({ name: "syncProblem" });
                this.setState({ warning: syncProgress.syncProblem });
            }
            await this.getDataFromDb();
        } catch (error) {
            if (error instanceof OperationCancelled) {
                log.info("sync messages cancelled");
            } else {
                AI.trackException({ exception: error });
                this.setError(error, this.state.t.error.syncFailed);
            }
        } finally {
            this.setState({ syncing: false, lastSynced });
        }
    }


    private setError = (error: unknown, userMessage: string): void => {
        this.setState({ error: userMessage });
        log.error(error);
    }


    private searchChatChanged: ComponentEventHandler<InputProps & { value: string; }> = (_e, data) => this.setState({ searchChat: data?.value ?? "" });


    private getContext = (): Promise<microsoftChats.Context> => new Promise(microsoftChats.getContext);


    private getDataFromDb = async (): Promise<void> => {
        try {
            const [dbChats, userCache] = await Promise.all([FindMsgChat.getAll(), FindMsgUserCache.getInstance()]);

            const cs = new Map(this.state.checkState);
            dbChats.forEach(c => cs.set(c.id, cs.get(c.id) ?? false));

            const { unknownUserDisplayName } = this.state.t;
            const users = await userCache.getKnownUsers();
            const searchUserOptions = users.map(({ id, displayName }) => ({ key: id, header: displayName || unknownUserDisplayName }));
            const uid = userId(this.state.teamsInfo.loginHint);

            const chats = await Promise.all(dbChats.map(async (c) => {
                let singleCounterpart: string | null = null;

                if (!c.topic) {
                    const member = await db.chatMembers.where(idx.chatMembers.$chatId$id).between([c.id, Dexie.minKey], [c.id, Dexie.maxKey]).first();

                    console.log("no topic");
                    console.log(`uid: ${uid}`);
                    console.dir(member);
                    console.log(uid);

                    if (member) {
                        singleCounterpart = userCache.getDisplayName(member.userId ?? member.id);
                        console.log(`scp: ${singleCounterpart}`);
                    } else if (uid) { // Bots are not listed as chat members by the MsGraph API...
                        const firstCounterpartMessage = await db.chatMessages.where(idx.chatMessages.$chatId$id)
                            .between([c.id, Dexie.minKey], [c.id, Dexie.maxKey])
                            .filter(m => m.authorId !== uid)
                            .first();

                        console.dir(firstCounterpartMessage);
                        if (firstCounterpartMessage) {
                            singleCounterpart = userCache.getDisplayName(firstCounterpartMessage.authorId);
                            console.log(`scp: ${singleCounterpart}`);
                        }
                    }
                }

                return { ...c, singleCounterpart };
            }));

            this.setState({ chats, checkState: cs, searchUserOptions });
        }
        catch (error) {
            AI.trackException({ exception: error });
            this.setError(error, this.state.t.error.indexedDbReadFailed);
        }
    };


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
     * Handle target chat/chat checkbox change
     * @param id Id of the chat or chat to be selected or deselected (undefined toggles the "search all" checkbox)
     */
    private chatCheckChanged = (id?: string) => {
        const { checkState, checkAll } = this.state;

        if (!id) {
            this.setState({ checkAll: !checkAll });
            return;
        }

        const newCheckState = new Map(checkState);
        newCheckState.set(id, !checkState.get(id));

        this.setState({ checkState: newCheckState });
    }


    /**
     * Search for messages
     */
    private search = async () => {
        try {
            const { searchChat, checkState, checkAll, searchTimeFrom, searchTimeTo, searchUsers } = this.state;
            const [cancel, checkCancel] = cancellationNoThrow();
            const userIds = new Set<string>(searchUsers.map((u) => u.key));
            const chats = new Set(checkAll ? null : Array.from(checkState.entries()).filter(([, v]) => v).map(([k]) => k));

            this.setState({ searching: true, searchCancel: cancel });
            const messages = await FindMsgChatMessage.search(searchChat, searchTimeFrom, searchTimeTo, chats, userIds, checkCancel);
            messages.sort(FindMsgChatMessage.compareByTouched);
            log.info(`Found ${messages.length} messages`);
            this.setState({ searchResults: messages });
        } catch (error) {
            AI.trackException({ exception: error });
            this.setError(error, this.state.t.error.searchFailed);
            this.setState({ searchResults: [] });
        } finally {
            this.setState({ searching: false, searchCancel: nop });
        }
    };
}

