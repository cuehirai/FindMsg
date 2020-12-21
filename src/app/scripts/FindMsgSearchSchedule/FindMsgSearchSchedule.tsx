import React from "react";
import * as log from '../logger';
import { DateRange, IMyOwnState, initialDisplayCount, ITeamsAuthComponentState, loadMoreCount, SearchUserItem, TeamsBaseComponentWithAuth } from "../msteams-react-base-component-with-auth";
import { SyncWidget } from "../SyncWidget";
import { Button, Dropdown, Flex, Input, Segment, Text } from "../ui";
import { cancellation, OperationCancelled } from "../utils";
import { AI } from "../appInsights";
import { EventTable } from "./EventTable";
import { EventOrder, FindMsgEvent } from "../db/Event/FindMsgEventEntity";
import { ISyncFunctionArg, OrderByDirection } from "../db/db-accessor-class-base";
import { IFindMsgEvent } from "../db/Event/IFindMsgEvent";
import { AppConfig } from "../../../config/AppConfig";
import * as du from "../dateUtils";

/** スケジュール検索用ロケール依存リソース定義 */
export interface IFindMsgScheduleTranslation {
    pageTitle: string;
    filterByStart: string;
    filterByOrganizer: string;
}

/** 検索結果情報 */
interface ISearchResult {
    events: IFindMsgEvent[];
    hasMore: boolean;
    order: EventOrder;
    dir: OrderByDirection;
    rowCount: number;
}

/** クラス固有のステートプロパティ */
interface IFindMsgSearchScheduleState extends IMyOwnState {
    searchResult: ISearchResult;
}

const stateSaveKey = `${AppConfig.AppInfo.name}SearchScheduleSavedState`

interface ISavedState {
    /** ソートキー */
    order: EventOrder;
    /** ソート方向 */
    dir: OrderByDirection;
    /** 下記条件で絞り込んだ行数 */
    rowCount: number;
    /** リストのフィルタ機能を提供する場合のフィルタ指定文字 */
    filterInput: string;
    /** 日付範囲の種類（日付範囲を使用する場合用） */
    searchTime: DateRange;
    /** 日付範囲（from） */
    searchTimeFrom: Date;
    /** 日付範囲（to） */
    searchTimeTo: Date;   
    /** ユーザ選択プルダウンで選択されているユーザ */
    searchUsers: SearchUserItem[];
}

/**
 * スケジュール検索コンポーネント
 */
export class FindMsgSearchSchedule extends TeamsBaseComponentWithAuth {
    protected isFixedPageSize = false;
    protected showInformation = true;
    protected async setAdditionalState(newstate: ITeamsAuthComponentState, context?: microsoftTeams.Context, inTeams?: boolean): Promise<void> {
        if (context && inTeams? true : false) {
            log.info(`★★★ setAdditionalState is called from componentDidMount; hosted in teams ★★★`);
        }
        const saved = localStorage.getItem(stateSaveKey);
        if (saved) {
            log.info(`★★★ Saved condition fetched from localStorage: [${saved}] ★★★`);
            const filter = JSON.parse(saved);

            newstate.filterInput = String(filter.filterInput);
            newstate.searchTime = Number(filter.searchTime);
            newstate.searchTimeFrom = du.parseISO(String(filter.searchTimeFrom));
            newstate.searchTimeTo = du.parseISO(String(filter.searchTimeTo));
            newstate.searchUsers = filter.searchUsers;
            const me = newstate.me as IFindMsgSearchScheduleState;
            me.searchResult.dir = filter.dir;
            me.searchResult.order = filter.order;
            me.searchResult.rowCount = Number(filter.rowCount);
        }
    }

    protected requireDatabase = true;
    
    protected requireMicrosoftLogin = true;

    protected isTeamAndChannelComboIncluded = false;

    protected isUsingStorage = true;

    protected GetPageTitle(): string {
        return this.state.translation.schedule.pageTitle;
    }
    
    protected startSync = async():Promise<void> => {
        log.info(`▼▼▼ startSync START ▼▼▼`);
        let { lastSynced } = this.state;

        try {
            const {
                translation: { syncProgress }
            } = this.state;
            const [cancel, throwfn] = cancellation();

            this.setState({ syncing: true, syncCancel: cancel, syncCancelled: false, error: "", warning: "" });
            const arg: ISyncFunctionArg = {
                client: this.msGraphClient,
                checkCancel: throwfn,
                progress: this.reportProgress,
                subentity: true,
                translate: this.state.translation,
            };
            const syncResult = await FindMsgEvent.sync(arg);

            if (syncResult) {
                lastSynced = await FindMsgEvent.getLastSynced();
            } else {
                AI.trackEvent({ name: "syncProblem" });
                this.setState({ warning: syncProgress.syncProblem });
            }

            await this.initBaseInfo();
        } catch (error) {
            if (error instanceof OperationCancelled) {
                log.info("sync cancelled");
            } else {
                AI.trackException({ exception: error });
                this.setError(error, this.state.translation.error.syncFailed);
            }
        } finally {
            this.setState({ syncing: false, lastSynced });
        }
        log.info(`▲▲▲ startSync END ▲▲▲`);

    };

    protected async GetLastSync(): Promise<Date> {
        log.info(`■■■ GetLastSync ENTERED ■■■`)
        return await FindMsgEvent.getLastSynced();
    }

    protected CreateMyState(): IFindMsgSearchScheduleState {
        log.info(`■■■ CreateMyState ENTERED ■■■`)
        return {
            initialized: true,
            searchResult: {
                events: [],
                hasMore: false,
                order: EventOrder.start,
                dir: OrderByDirection.descending,
                rowCount: 0,
            },
        };
    }

    protected setMyState(value?: ISearchResult): IMyOwnState {
        log.info(`▼▼▼ setMyState START ▼▼▼`);
        const me = (this.state.me as IFindMsgSearchScheduleState);
        if (value) {
            me.searchResult = {
                events: value.events,
                hasMore: value.hasMore,
                order: value.order,
                dir: value.dir,
                rowCount: value.rowCount,
                };    
        } else {
            me.searchResult = {
                events: [],
                hasMore: false,
                order: EventOrder.start,
                dir: OrderByDirection.descending,
                rowCount: 0,
                };    
        }
        log.info(`▲▲▲ setMyState END ▲▲▲`);
        return me;
    }

    protected renderContentTop(): JSX.Element {
        log.info(`▼▼▼ renderContentTop START ▼▼▼`);
        const {
            loading,
            syncing, syncStatus, syncCancelled, lastSynced,
            filterInput,
            searchUserOptions,
            translation: {
                common: {
                    noSelection,
                },
                schedule: {
                    filterByStart,
                    filterByOrganizer,
                },
                sync, filter,
            }
        } = this.state;
    
        const res:JSX.Element = (
            <div>
                <Segment>
                    <Flex gap="gap.large">
                        <Flex.Item grow shrink>
                            <Flex gap="gap.small">
                                <Segment>
                                    <Input
                                        type="text"
                                        label={filter}
                                        labelPosition="above"
                                        value={filterInput}
                                        onChange={this.onFilterChanged}
                                    />
                                </Segment>

                                <Segment>
                                    <Flex column gap="gap.small">
                                        <Text content={filterByStart} />
                                        {this.renderTermSelection()}
                                    </Flex>
                                </Segment>

                                <Segment>
                                    <Flex column>
                                        <Text content={filterByOrganizer} />
                                        <Dropdown
                                            multiple clearable search
                                            position="above"
                                            placeholder={noSelection}
                                            items={searchUserOptions}
                                            onChange={this.onSearchUserChanged}
                                        />
                                    </Flex>
                                </Segment>

                                <Flex.Item grow>
                                    <div />
                                </Flex.Item>

                                <Flex.Item align="start">
                                    <SyncWidget
                                        t={sync}
                                        syncStart={this.startSync}
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

            </div>
        );
        log.info(`▲▲▲ renderContentTop END ▲▲▲`);
        return res;
    }

    protected renderContent(): JSX.Element {
        log.info(`▼▼▼ renderContent START ▼▼▼`);
        const {
            loading,
            filterString,
            translation: {
                dateFormat, dateTimeFormat, eventTable, unknownUserDisplayName,
            }
        } = this.state;
        const { events, order, dir } = (this.state.me as IFindMsgSearchScheduleState).searchResult;

        const res:JSX.Element = (
            <EventTable translation={eventTable} dateFormat={dateFormat} dateTimeFormat={dateTimeFormat} events={events} dir={dir} order={order} sort={this.getEvents} loading={loading} filter={filterString} unknownUserDisplayName={unknownUserDisplayName} />
        );
        log.info(`▲▲▲ renderContent END ▲▲▲`);
        return res;
    }
    
    protected renderContentBottom(): JSX.Element {
        log.info(`▼▼▼ renderContentBottom START ▼▼▼`);
        const {
            translation: {
                common,
            }
        } = this.state;
        const { hasMore } = (this.state.me as IFindMsgSearchScheduleState).searchResult;
        let res:JSX.Element = <div/>;
        if (hasMore) {
            res = <Button onClick={this.loadMoreEvents} content={common.loadMore} />;
        }
        log.info(`▲▲▲ renderContentBottom END ▲▲▲`);
        return res;
    }

    protected setStateCallBack(): void {
        this.getUserOptions();
        const { order, dir } = (this.state.me as IFindMsgSearchScheduleState).searchResult;
        this.getEvents(order, dir);
    }
    
    protected onFilterChangedCallBack(): void {
        const { order, dir } = (this.state.me as IFindMsgSearchScheduleState).searchResult;
        this.getEvents(order, dir);
    }

    protected onSearchUserChangedCallBack(): void {
        this.getUserOptions();
        const { order, dir } = (this.state.me as IFindMsgSearchScheduleState).searchResult;
        this.getEvents(order, dir);
    }
    protected onTeamOrChannelChangedCallBack(): void {
        //実装なし
    }
    protected onDateRangeChangedCallBack(): void {
        const { order, dir } = (this.state.me as IFindMsgSearchScheduleState).searchResult;
        this.getEvents(order, dir);
    }

    private getEvents = async (order: EventOrder, dir: OrderByDirection): Promise<void> => {
        log.info(`▼▼▼ getEvents START ▼▼▼`);
        const {
            filterInput, searchTime, searchTimeFrom, searchTimeTo, searchUsers,
        } = this.state;
        const {rowCount} = (this.state.me as IFindMsgSearchScheduleState).searchResult;

        this.setState({ loading: true });

        try {
            const maxRows = (initialDisplayCount < rowCount)? rowCount : initialDisplayCount;

            const userIds = new Set<string>(searchUsers.map((u) => u.key));
            const [events, hasMore] = await FindMsgEvent.fetch(order, dir, 0, maxRows, filterInput, searchTimeFrom, searchTimeTo, userIds);
            log.info(` ★★★ fetched [${events.length}] events from DB ★★★`);
            const value: ISearchResult = {
                events: events,
                hasMore: hasMore,
                order: order,
                dir: dir,
                rowCount: events.length,
            };

            const save: ISavedState = {
                order: order,
                dir: dir,
                rowCount: value.events.length,
                filterInput: filterInput,
                searchTime: searchTime,
                searchTimeFrom: searchTimeFrom,
                searchTimeTo: searchTimeTo,
                searchUsers: searchUsers
            };
            const json = JSON.stringify(save);
            // sessionStorageではTeamsアプリでは他のページを表示した時点で無効になってしまうらしい。
            // ※Teams-WebではsessionStorageがTab移動の他、他のアプリを使用したあとでも有効だったので、結局localStorageと変わらない。
            localStorage.setItem(stateSaveKey, json);
            log.info(`★★★ Filter condition saved: [${json}] ★★★`);

            this.setState({
                filterString: filterInput,
                searchTimeFrom: searchTimeFrom,
                searchTimeTo: searchTimeTo,
                searchUsers: searchUsers,
                me: this.setMyState(value),
            });
        } catch (error) {
            AI.trackException({ exception: error });
            this.setError(error, this.state.translation.error.indexedDbReadFailed);
        } finally {
            this.setState({ loading: false });
        }
        log.info(`▲▲▲ getEvents END ▲▲▲`);

    };

    private loadMoreEvents = async () => {
        log.info(`▼▼▼ loadMoreEvents START ▼▼▼`);
        const {
            searchTime, filterInput, searchTimeFrom, searchTimeTo, searchUsers,
        } = this.state;
        const searchResult = (this.state.me as IFindMsgSearchScheduleState).searchResult

        try {
            this.setState({ loading: true });

            const userIds = new Set<string>(searchUsers.map((u) => u.key));
            const [newEvents, hasMore] = await FindMsgEvent.fetch(searchResult.order, searchResult.dir, searchResult.events.length, loadMoreCount, filterInput, searchTimeFrom, searchTimeTo, userIds);

            const me = (this.state.me as IFindMsgSearchScheduleState)
            me.searchResult = {
                events: [...searchResult.events, ...newEvents],
                hasMore: hasMore, 
                order: searchResult.order, 
                dir: searchResult.dir  ,
                rowCount: [...searchResult.events, ...newEvents].length,         
            }

            const save: ISavedState = {
                order: me.searchResult.order,
                dir: me.searchResult.dir,
                rowCount: me.searchResult.rowCount,
                filterInput: filterInput,
                searchTime: searchTime,
                searchTimeFrom: searchTimeFrom,
                searchTimeTo: searchTimeTo,
                searchUsers: searchUsers
            };
            const json = JSON.stringify(save);
            localStorage.setItem(stateSaveKey, json);
            log.info(`★★★ Filter condition saved (more): [${json}] ★★★`);

            this.setState({
                me: me
            });
        } catch (error) {
            AI.trackException({ exception: error });
            this.setError(error, this.state.translation.error.indexedDbReadFailed);
        } finally {
            this.setState({ loading: false });
        }
        log.info(`▲▲▲ loadMoreEvents END ▲▲▲`);
    }

    /**
     * ユーザオプションは主催者（Teamsのユーザとは限らない）なのでオーバーライド
     */
    protected getUserOptions = async (): Promise<void> => {
        try {
            const users = await FindMsgEvent.getOrganizers();
            const searchUserOptions = users.map((rec) => ({ key: rec, header: rec, selected: this.isExist(rec, this.state.searchUsers) }));

            this.setState({ searchUserOptions });
        }
        catch (error) {
            AI.trackException({ exception: error });
            this.setError(error, this.state.translation.error.indexedDbReadFailed);
        }
    }

    private isExist(key: string, arr: SearchUserItem[]): boolean {
        let res = false;
        arr.forEach(rec => {
            if (key === rec.key) {
                res = true;
            }
        });
        return res;
    }

}