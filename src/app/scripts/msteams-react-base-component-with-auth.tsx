import * as React from "react";
import { AuthProviderCallback, Client } from "@microsoft/microsoft-graph-client";
import { AuthError, getAuthTokenSilent, haveUserInfo, loginPopup } from "./auth/auth";
import { IMessageTranslation } from "./i18n/IMessageTranslation";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "./msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

import { assertT, nop, progressFn, storage } from "./utils";
import * as strings from './i18n/messages';
import * as log from './logger';
import { ComponentEventHandler, Alert, AlertProps, Dialog, Divider, Provider, Text, Page, DropdownItemProps, DropdownProps, Flex, Dropdown, RadioGroup, RadioGroupItemProps, ShorthandCollection, DatePicker, InputProps } from "./ui";
import { AI } from "./appInsights";
import * as du from "./dateUtils";
import { SyncControl, SyncState } from "./SyncWidget";
import { invalidDate } from "./dateUtils";
import { FindMsgChannel, FindMsgTeam, FindMsgUserCache } from "./db";
import { ICommonMessage } from "./i18n/ICommonMessage";
import { StoragePermissionIndicator } from "./StoragePermissionIndicator";
import { StoragePermissionWidget } from "./StoragePermissionWidget";

/** ユーザ選択コンボボックスの要素タイプ */
export declare type SearchUserItem = DropdownItemProps & { key: string };
/** コンボボックス選択肢要素のタイプ */
export declare type DropdownItemPropsKey = DropdownItemProps & { key: string };
/** hasMore制御が必要な場合の標準初期表示件数 */
export const initialDisplayCount = 50;
/** hasMore制御が必要な場合の標準追加表示件数 */
export const loadMoreCount = 25;
/** 日付範囲指定種類用の列挙体 */
export enum DateRange { AllTime, PastWeek, PastMonth, PastYear, Custom }

/** 日付範囲指定コンポーネント用プロパティ */
export interface DateRangeRadioGroupItemProps extends RadioGroupItemProps {
    value: DateRange;
}

/** Teamsに関連するステータス管理用 */
export interface ITeamsInfo {
    locale: string | null;
    groupId: string;
    channelId: string;
    entityId: string | null;
    subEntityId: string | null;
    loginHint: string;
    teamName: string;
    channelName: string;
    teamOptions: DropdownItemPropsKey[];
    channelOptions: DropdownItemPropsKey[][];
}

/** クラス固有のステータス（これを継承してクラス固有ステータスを管理してください。通常は検索結果などを含めます。） */
export interface IMyOwnState {
    initialized: boolean;
}

/** 基本のステータス管理用 */
export interface ITeamsAuthComponentState extends ITeamsBaseComponentState, SyncState, SyncControl {
    /** ページローディング中フラグ */
    loading: boolean;

    /** エラーメッセージ */
    error: string;
    /** 警告メッセージ */
    warning: string;

    /** Teamsに関連するステータス */
    teamsInfo: ITeamsInfo;

    /** リストのフィルタ機能を提供する場合のフィルタ指定文字 */
    filterInput: string;
    /** リストのフィルタ機能を提供する場合のフィルタ指定文字(入力値保存用) */
    filterString: string;

    /** プルダウンの使用可否（プルダウンを使用する場合用） */
    dropdownDisabled: boolean;
    /** Teamプルダウンで現在選択されているインデックス(0=すべて) */
    teamIdx: number;
    /** Channelプルダウンで現在選択されているインデックス(0=すべて) */
    channelIdx: number;

    /** 日付範囲の種類（日付範囲を使用する場合用） */
    searchTime: DateRange;
    /** 日付範囲（from） */
    searchTimeFrom: Date;
    /** 日付範囲（to） */
    searchTimeTo: Date;   

    /** ユーザ選択プルダウンで選択されているユーザ */
    searchUsers: SearchUserItem[];
    /** ユーザ選択プルダウンの選択肢 */
    searchUserOptions: ShorthandCollection<DropdownItemProps>;

    /** 認証プロバイダのコールバック */
    authInProgress: AuthProviderCallback;
    /** 認証結果 */
    authResult: AuthError | null;
    /** ログインが必要かどうか */
    loginRequired: boolean;

    /** 言語依存リソース */
    translation: IMessageTranslation;

    /** 永続化ストレージの許可が必要かどうか */
    askForStoragePermission: boolean;

    /** クラス固有のステータス */
    me: IMyOwnState;
}

/**
 * Teams用コンポーネントの基底クラス（msal認証フローサポート）
 */
export abstract class TeamsBaseComponentWithAuth extends TeamsBaseComponent<never, ITeamsAuthComponentState> {

    /** MS Graph クライアント */
    protected msGraphClient: Client;

    protected getContext = (): Promise<microsoftTeams.Context> => new Promise(microsoftTeams.getContext);

    protected filterTimeout = 0;

    constructor(props: never) {
        super(props);
        log.info(`▼▼▼ constructor START ▼▼▼`);

        const cid = this.getQueryVariable("cid");
        const gid = this.getQueryVariable("gid");
        const locale = this.getQueryVariable("l");

        this.state = {
            // ITeamsBaseComponentState
            theme: this.getTheme(this.getQueryVariable("theme")),

            // SyncState
            syncing: false,
            syncStatus: "",
            syncCancel: nop,
            syncCancelled: false,
            lastSynced: invalidDate(),

            // 当クラスで定義したプロパティ
            loading: true,
            error: "",
            warning: "",

            authInProgress: nop,
            authResult: null,
            loginRequired: false,

            filterInput: "",
            filterString: "",

            dropdownDisabled: !!(cid && gid),
            teamIdx: 0,
            channelIdx: 0,

            searchTime: DateRange.AllTime,
            searchTimeFrom: du.invalidDate(),
            searchTimeTo: du.invalidDate(),

            searchUsers: [],
            searchUserOptions: [],

            teamsInfo: {
                channelId: cid ?? "",
                groupId: gid ?? "",
                locale: locale ?? null,
                entityId: this.getQueryVariable("eid") ?? "",
                subEntityId: this.getQueryVariable("sid") ?? "",
                loginHint: this.getQueryVariable("hint") ?? "",
                channelName: "",
                teamName: "",
                channelOptions: [[]],
                teamOptions: [],
            },

            translation: strings.get(locale),
            askForStoragePermission: false,

            // 派生クラスで独自に定義するプロパティ
            me: this.CreateMyState(),
        }
        // CreateMyStateが正しく実装されているかをチェック
        if (!this.state.me.initialized) {
            throw new Error("CreateMyState is not implemented.");
        }
        this.msGraphClient = Client.init({ authProvider: this.authProvider });
        log.info(`▲▲▲ constructor END ▲▲▲`);
    }

    /**
     * このページで永続化ストレージを使用するかどうかを宣言してください。
     */
    protected abstract isUsingStorage: boolean;

    /**
     * このページでチーム・チャネルのコンボボックスを使用するかどうかを宣言してください。
     */
    protected abstract isTeamAndChannelComboIncluded: boolean;

    /**
     * ページタイトルを指定してください。
     */
    protected abstract GetPageTitle(): string;

    /**
     * コンストラクタから呼び出します。
     * オーバーライドしてクラス固有のステータスを設定してください。
     * ※initializedに必ずtrueを設定すること！
     */
    protected abstract CreateMyState(): IMyOwnState

    /**
     * componentDidMountから呼び出します。
     * オーバーライドしてクラス固有の初期化処理を実装してください。
     * @param context 
     */
    protected abstract setMyState(): IMyOwnState

    /**
     * renderから呼び出します。
     * オーバーライドしてコンテンツ上部（検索条件など）の要素を記述してください。
     */
    protected abstract renderContentTop(): JSX.Element

    /**
     * renderから呼び出します。
     * オーバーライドしてコンテンツ（一覧など）の要素を記述してください。
     */
    protected abstract renderContent(): JSX.Element

    /**
     * renderから呼び出します。
     * オーバーライドしてコンテンツの追加要素を記述してください。
     */
    protected abstract renderContentBottom(): JSX.Element

    /**
     * オーバーライドしてinitBaseInfoにおけるsetStateのコールバックを実装してください。
     */
    protected abstract setStateCallBack(): void;
    
    /**
     * オーバーライドしてonFilterChangedにおけるsetStateのコールバックを実装してください。
     */
    protected abstract onFilterChangedCallBack(): void;

    /**
     * オーバーライドしてonSearchUserChangedにおけるsetStateのコールバックを実装してください。
     */
    protected abstract onSearchUserChangedCallBack(): void;

    /***
     * オーバーライドしてonTeamChanged/onChannelChangedにおけるsetStateのコールバックを実装してください。
     */
    protected abstract onTeamOrChannelChangedCallBack(): void;

    /**
     * オーバーライドしてsearchTimeOptionChangedにおけるsetStateのコールバックを実装してください。
     */
    protected abstract onDateRangeChangedCallBack(): void;

    /**
     * Graphからのデータを同期する場合はオーバーライドして同期処理を実装してください。
     * 使用時はSyncWidgetの属性にこのメソッドを設定します。（例：syncStart={this.startSync}）
     */
    protected abstract async startSync(): Promise<void> 

    /**
     * 同期を必要とする場合はオーバーライドしてクラス固有の最終同期日時を返却してください。
     * @param target 
     */
    protected abstract async GetLastSync(target?: string): Promise<Date>
    
    // コンポーネントがマウントされたときに呼び出されます。
    // ここでステータスをセットするとページが再描画されます。
    // クラス固有のステータスを設定するにはsetMyStateをオーバーライドしてください。
    public async componentDidMount(): Promise<void> {
        log.info(`▼▼▼ componentDidMount START ▼▼▼`);
        this.updateTheme(this.getQueryVariable("theme"));

        if (await this.inTeams(2000)) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);

            const context = await this.getContext();
            log.info("context", context);
            this.updateTheme(context.theme);
            microsoftTeams.appInitialization.notifySuccess();
            this.initBaseInfo(context);
        } else {
            this.initBaseInfo();
        }

        if(this.isUsingStorage) {
            this.setState({
                askForStoragePermission: !storage.granted() && storage.askForPermission,
                loading: false
            });
        }
        log.info(`▲▲▲ componentDidMount END ▲▲▲`);
    }

    // コンポーネントを描画します。
    // renderContentTop,renderContent,renderContentBottomをそれぞれオーバーライドしてページを完成させてください。
    public render(): JSX.Element {
        log.info(`▼▼▼ render START ▼▼▼`);
        const {
            loading,
            theme,
            error,
            warning,

            authResult,
            loginRequired,
            askForStoragePermission,

            translation: {
                footer, auth,
                storagePermission,
            }
        } = this.state;

        const contentTop = this.renderContentTop();
        const content = this.renderContent();
        const contentBottom = this.renderContentBottom();

        const res:JSX.Element = (
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

                    {contentTop}

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
                    {askForStoragePermission && <StoragePermissionWidget granted={this.storagePermissionGranted} t={storagePermission} />}

                    {content}
                    {contentBottom}

                    <div style={{ flex: 1 }} />
                    <Divider />
                    {!this.isUsingStorage && <Text size="smaller" content={footer} />}
                    {this.isUsingStorage && <Flex space="between">
                        <Text size="smaller" content={footer} />
                        <StoragePermissionIndicator loading={loading} />
                    </Flex>}
                </Page>
            </Provider >
        );
        log.info(res);
        log.info(`▲▲▲ render END ▲▲▲`);

        return res;
    }

    /**
     * コンポーネントマウント時の初期化処理
     * @param context 
     */
    protected initBaseInfo = async (context?: microsoftTeams.Context): Promise<void> => {
        log.info(`▼▼▼ initBaseInfo START ▼▼▼`);
        const loginHint = context?.loginHint ?? this.state.teamsInfo.loginHint;
        let groupId = context?.groupId ?? this.state.teamsInfo.groupId;
        let channelId = context?.channelId ?? this.state.teamsInfo.channelId;
        const entityId = context?.entityId ?? "";
        const subEntityId= context?.subEntityId ?? null;
        const locale = context?.locale ?? this.state.teamsInfo.locale;
        const teamName = context?.teamName ?? "";
        const channelName = context?.channelName ?? "";
        const translation = strings.get(locale);

        microsoftTeams.setFrameContext({
            contentUrl: location.href,
            websiteUrl: location.href,
        });

        const lastSynced = await this.GetLastSync(channelId);

        let teamIdx: number;
        let channelIdx: number;
        let teamOptions: DropdownItemPropsKey[];
        let channelOptions: DropdownItemPropsKey[][];

        if (!this.isTeamAndChannelComboIncluded || this.state.dropdownDisabled) {
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
                header: translation.common.allTeams,
                key: "",
                selected: false,
            };

            const allChannelsOption: DropdownItemPropsKey = {
                header: translation.common.allChannels,
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

            // 各Teamに属するChannelのリストに（すべてのチャネル）を追加しておく
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
                entityId, subEntityId,
            },
            teamIdx, channelIdx,
            lastSynced,
            loginRequired: !haveUserInfo(loginHint),
            translation: translation,
            me: this.setMyState(),
        }, this.setStateCallBack);

        document.title = this.GetPageTitle();

        log.info(`▲▲▲ initBaseInfo END ▲▲▲`);
    }

    /**
     * (ユーティリティ)ユーザ選択プルダウン
     */
    protected getUserOptions = async (): Promise<void> => {
        try {
            const userCache = await FindMsgUserCache.getInstance();

            const { unknownUserDisplayName } = this.state.translation;
            const users = await userCache.getKnownUsers();
            const searchUserOptions = users.map(({ id, displayName }) => ({ key: id, header: displayName || unknownUserDisplayName }));

            this.setState({ searchUserOptions });
        }
        catch (error) {
            AI.trackException({ exception: error });
            this.setError(error, this.state.translation.error.indexedDbReadFailed);
        }
    }

    /**
     * (ユーティリティ)Team＋channelプルダウン
     */
    protected renderTeamAndChannelPulldown(): JSX.Element {
        log.info(`▼▼▼ renderTeamAndChannelPulldown START ▼▼▼`);
        const {
            teamIdx,
            channelIdx,
            dropdownDisabled,

            teamsInfo: {
                teamOptions, channelOptions
            },
        } = this.state;
        const res:JSX.Element = (
            <Flex gap="gap.small" wrap>
                <Dropdown disabled={dropdownDisabled} items={teamOptions} value={teamOptions[teamIdx]} onChange={this.onTeamChanged} />
                <Dropdown disabled={dropdownDisabled} items={channelOptions[teamIdx]} value={channelOptions[teamIdx][channelIdx]} onChange={this.onChannelChanged} />
            </Flex>
        );
        log.info(`▲▲▲ renderTeamAndChannelPulldown END ▲▲▲`);
        return res;
    }

    /**
     * (ユーティリティ)日付範囲選択オプション
     */
    protected renderTermSelection(): JSX.Element {
        log.info(`▼▼▼ renderTermSelection START ▼▼▼`);
        const {
            searchTime,
        } = this.state;

        const res: JSX.Element = (
            <Flex column gap="gap.small">
                <RadioGroup checkedValue={searchTime} items={this.searchTimeOptions()} onCheckedValueChange={this.searchTimeOptionChanged} />
                {searchTime === DateRange.Custom && this.renderCustomTerm()}
            </Flex>
        );
        log.info(`▲▲▲ renderTermSelection END ▲▲▲`);

        return res;
    }

    /**
     * (ユーティリティ)カスタム日付範囲指定
     */
    protected renderCustomTerm(): JSX.Element {
        log.info(`▼▼▼ renderCustomTerm START ▼▼▼`);
        const {
            searchTimeFrom,
            searchTimeTo,
            translation: {
                common: {
                    from, to,
                },
            },
        } = this.state;

        const res: JSX.Element = (
            <Flex gap="gap.medium">
                <DatePicker label={from} value={searchTimeFrom} onSelectDate={this.searchTimeFromChanged} formatDate={this.formatDate} />
                <DatePicker label={to} value={searchTimeTo} onSelectDate={this.searchTimeToChanged} formatDate={this.formatDate} />
            </Flex>
        );
        log.info(`▲▲▲ renderCustomTerm END ▲▲▲`);

        return res;
    }

    /**
     * (ユーティリティ)日付範囲選択オプション※ロケールが変更された場合は自動的に再生成します
     * 実装例：`<RadioGroup checkedValue={searchTime} items={this.searchTimeOptions()} onCheckedValueChange={this.searchTimeOptionChanged} />`
     */
    protected searchTimeOptions: () => ShorthandCollection<DateRangeRadioGroupItemProps> = (() => {
        log.info(`▼▼▼ searchTimeOptions START ▼▼▼`);
        let lastState: ICommonMessage | null = null;
        let lastOptions: ShorthandCollection<DateRangeRadioGroupItemProps> = [];

        return () => {
            if (this.state.translation.common === lastState) {
                log.info(`▲▲▲ searchTimeOptions END (no changes) ▲▲▲`);
                return lastOptions;
            }

            const {
                searchTimeAll,
                searchTimePastWeek,
                searchTimePastMonth,
                searchTimePastYear,
                searchTimeCustom,
            } = lastState = this.state.translation.common;

            lastOptions = [
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
            log.info(`▲▲▲ searchTimeOptions END ▲▲▲`);
            return lastOptions;
        };
    })();

    /** 永続化ストレージ許可済みをセット */
    protected storagePermissionGranted = (): void => {
        this.setState({ askForStoragePermission: false });
    }

    /**
     * (ユーティリティ)ユーザ選択プルダウンのChangeイベント
     * @param _e 
     * @param data 
     */
    protected onSearchUserChanged = (_e: unknown, data: DropdownProps):void => {
        const values = data.value as SearchUserItem[];
        this.setState({ searchUsers: [...values] }, this.onSearchUserChangedCallBack);
    };

    /**
     * (ユーティリティ)フィルタのChangeイベント
     * 実装例：<Input type="text"label={filter} labelPosition="inline" value={filterInput} onChange={this.onFilterChanged} />
     * @param _ 
     * @param data 
     */
    protected onFilterChanged: ComponentEventHandler<InputProps & { value: string; }> = (_: unknown, data): void => {
        log.info(`▼▼▼ onFilterChanged START ▼▼▼`);
        this.setState({ filterInput: data?.value ?? "" }, () => {
            window.clearTimeout(this.filterTimeout);
            this.filterTimeout = window.setTimeout(() => {
                this.onFilterChangedCallBack();
            }, 250);
        });
        log.info(`▲▲▲ onFilterChanged END ▲▲▲`);
    }

    /**
     * (ユーティリティ)TeamコンボボックスのChangeイベント
     * 実装例：`<Dropdown disabled={dropdownDisabled} items={teamOptions} value={teamOptions[teamIdx]} onChange={this.onTeamChanged} />`
     * @param _ 
     * @param data 
     */
    protected onTeamChanged = (_: unknown, data: DropdownProps): void => {
        log.info(`▼▼▼ onTeamChanged START ▼▼▼`);
        const selected = data.value as DropdownItemPropsKey;
        const { teamsInfo } = this.state;
        const { teamOptions, channelOptions } = teamsInfo;

        const newIdx = teamOptions.findIndex(t => t.key === selected.key);
        const newOpts = teamOptions.map((to, i) => ({ ...to, selected: i === newIdx }));

        const newChannelIdx = Math.max(0, channelOptions[newIdx].findIndex(co => co.selected));
        channelOptions[newIdx][newChannelIdx].selected = true;

        this.setState({
            teamIdx: newIdx,
            channelIdx: newChannelIdx,
            teamsInfo: {
                ...teamsInfo,
                teamOptions: newOpts,
                channelId: channelOptions[newIdx][newChannelIdx].key,
                groupId: teamOptions[newIdx].key,
            }
        }, this.onTeamOrChannelChangedCallBack);
        log.info(`▲▲▲ onTeamChanged END ▲▲▲`);
    }

    /**
     * (ユーティリティ)ChannelコンボボックスのChangeイベント
     * 実装例：<Dropdown disabled={dropdownDisabled} items={channelOptions[teamIdx]} value={channelOptions[teamIdx][channelIdx]} onChange={this.onChannelChanged} />
     * @param _ 
     * @param data 
     */
    protected onChannelChanged = (_: unknown, data: DropdownProps): void => {
        log.info(`▼▼▼ onChannelChanged START ▼▼▼`);
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
        }, this.onTeamOrChannelChangedCallBack);
        log.info(`▲▲▲ onChannelChanged END ▲▲▲`);
    }

    /**
     * (ユーティリティ)日付範囲オプションのchangeイベント
     * 実装例：`<RadioGroup checkedValue={searchTime} items={this.searchTimeOptions()} onCheckedValueChange={this.searchTimeOptionChanged} />`
     * @param _e 
     * @param props 
     */
    protected searchTimeOptionChanged: ComponentEventHandler<RadioGroupItemProps> = (_e, props) => {
        log.info(`▼▼▼ searchTimeOptionChanged START ▼▼▼`);
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
            }, this.onDateRangeChangedCallBack);
        } catch (error) {
            AI.trackException({ exception: error });
            this.setError(error, this.state.translation.error.internalError);
        } finally {
            log.info(`▲▲▲ searchTimeOptionChanged END ▲▲▲`);
        }
    };

    /**
     * (ユーティリティ)日付範囲指定（from）のchangeイベント
     * 実装例：`<DatePicker label={from} value={searchTimeFrom} onSelectDate={this.searchTimeFromChanged} formatDate={this.formatDate} />`
     * @param d 
     */
    protected searchTimeFromChanged = (d: Date | undefined | null): void => {
        log.info(`▼▼▼ searchTimeFromChanged START ▼▼▼`);
        if (d) this.setState({ searchTimeFrom: du.startOfDay(d) }, this.onDateRangeChangedCallBack)
        log.info(`▲▲▲ searchTimeFromChanged END ▲▲▲`);
    };

    /**
     * (ユーティリティ)日付範囲指定（to）のchangeイベント
     * 実装例：`<DatePicker label={to} value={searchTimeTo} onSelectDate={this.searchTimeToChanged} formatDate={this.formatDate} />`
     * @param d 
     */
    protected searchTimeToChanged = (d: Date | undefined | null): void => {
        log.info(`▼▼▼ searchTimeToChanged START ▼▼▼`);
        if (d) this.setState({ searchTimeTo: du.endOfDay(d) }, this.onDateRangeChangedCallBack)
        log.info(`▲▲▲ searchTimeToChanged END ▲▲▲`);
    };

    /**
     * (ユーティリティ)日付書式設定
     * 実装例：`<DatePicker label={from} value={searchTimeFrom} onSelectDate={this.searchTimeFromChanged} formatDate={this.formatDate} />`
     * @param date 
     */
    protected formatDate = (date?: Date | undefined): string => date ? du.format(date, this.state.translation.dateFormat) : "";
    
    /**
     * (ユーティリティ)日付範囲指定（from）の値を取得
     * @param value 
     */
    protected getDateFrom = (value: DateRange): Date => {
        switch (value) {
            case DateRange.Custom: return du.isValid(this.state.searchTimeFrom) ? this.state.searchTimeFrom : du.subMonths(du.startOfToday(), 1);
            case DateRange.PastWeek: return du.subWeeks(du.startOfToday(), 1);
            case DateRange.PastMonth: return du.subMonths(du.startOfToday(), 1);
            case DateRange.PastYear: return du.subYears(du.startOfToday(), 1);
            case DateRange.AllTime: return du.invalidDate();
            default: return du.invalidDate();
        }
    }

    /**
     * (ユーティリティ)日付範囲指定（to）の値を取得
     * @param value 
     */
    protected getDateTo = (value: DateRange): Date => {
        switch (value) {
            case DateRange.Custom: return du.isValid(this.state.searchTimeTo) ? this.state.searchTimeTo : du.endOfToday();
            case DateRange.PastWeek: return du.endOfToday();
            case DateRange.PastMonth: return du.endOfToday();
            case DateRange.PastYear: return du.endOfToday();
            case DateRange.AllTime: return du.invalidDate();
            default: return du.invalidDate();
        }
    }

    /** (ユーティリティ)エラーメッセージセット */
    protected setError = (error: unknown, userMessage: string): void => {
        this.setState({ error: userMessage });
        log.error(error);
    }

    /** (ユーティリティ)ダイアログを閉じた後のエラー状態設定 */
    protected errorVisibilityChanged: ComponentEventHandler<AlertProps> = (_, data) => !data?.visible && this.setState({ error: "" });

    /** (ユーティリティ)ダイアログを閉じた後のウォーニング状態設定 */
    protected warningVisibilityChanged: ComponentEventHandler<AlertProps> = (_, data) => !data?.visible && this.setState({ warning: "" });

    /** (ユーティリティ)認証プロバイダのコールバック */
    protected authProvider = async (done: AuthProviderCallback): Promise<void> => {
        log.info(`▼▼▼ authProvider START ▼▼▼`);
        try {
            const token = await getAuthTokenSilent(this.state.translation.auth, this.state.teamsInfo.loginHint);
            this.setState({ authInProgress: nop, authResult: null, loginRequired: false });
            done(null, token);
        } catch (error) {
            AI.trackException({ exception: error });
            log.error(error);
            this.setState({ authInProgress: done, authResult: error });
        } finally {
            log.info(`▲▲▲ authProvider END ▲▲▲`);
        }
    }

    /** (ユーティリティ)マイクロソフトアカウントにログイン */
    protected login = async (): Promise<void> => {
        log.info(`▼▼▼ login START ▼▼▼`);
        const {
            authInProgress: done,
            authResult,
            teamsInfo: { loginHint },
            translation: { auth }
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
        } finally {
            log.info(`▲▲▲ login END ▲▲▲`);
        }
    };

    /**
     * (ユーティリティ)Report sync status callback
     * @param syncStatus
     */
    protected reportProgress: progressFn = syncStatus => this.setState({ syncStatus });

    /**
     * (ユーティリティ)Cancel an ongoing sync
     */
    protected cancelSync = (): void => {
        this.state.syncCancel();
        this.setState({ syncCancelled: true });
    };
}