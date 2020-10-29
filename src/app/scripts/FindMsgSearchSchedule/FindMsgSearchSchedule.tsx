import React from "react";
import * as log from '../logger';
import { IMyOwnState, initialDisplayCount, loadMoreCount, TeamsBaseComponentWithAuth } from "../msteams-react-base-component-with-auth";
import { SyncWidget } from "../SyncWidget";
import { Button, ComponentEventHandler, Flex, Input, InputProps, Segment } from "../ui";
import { cancellation, OperationCancelled } from "../utils";
import { AI } from "../appInsights";
import { EventTable } from "./EventTable";
import * as strings from '../i18n/messages';
import { EventOrder, FindMsgEvent } from "../db/Event/FindMsgEventEntity";
import { ISyncFunctionArg, OrderByDirection } from "../db/db-accessor-class-base";
import { IFindMsgEvent } from "../db/Event/IFindMsgEvent";

export interface IFindMsgScheduleTranslation {
    pageTitle: string;
}

interface ISearchResult {
    events: IFindMsgEvent[];
    hasMore: boolean;
    order: EventOrder;
    dir: OrderByDirection;
}

interface IFindMsgSearchScheduleState extends IMyOwnState {
    searchResult: ISearchResult;
}

export class FindMsgSearchSchedule extends TeamsBaseComponentWithAuth {
    protected isTeamAndChannelComboIncluded = false;

    protected isUsingStorage = true;

    protected GetPageTitle(locale: string): string {
        const translation = strings.get(locale);
        const res = translation.schedule.pageTitle;
        return res;
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
                lastSynced = FindMsgEvent.getLastSynced();
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
        return FindMsgEvent.getLastSynced();
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
            },
        };
    }

    protected setMyState(): IMyOwnState {
        log.info(`▼▼▼ setMyState START ▼▼▼`);
        const me = (this.state.me as IFindMsgSearchScheduleState);
        me.searchResult = {
            events: [],
            hasMore: false,
            order: EventOrder.start,
            dir: OrderByDirection.descending,
            };
        log.info(`▲▲▲ setMyState END ▲▲▲`);
        return me;
    }

    protected renderContentTop(): JSX.Element {
        log.info(`▼▼▼ renderContentTop START ▼▼▼`);
        const {
            loading,
            syncing, syncStatus, syncCancelled, lastSynced,
            filterInput,
            translation: {
                sync, filter,
            }
        } = this.state;

        const res:JSX.Element = (
            <Segment>
                <Flex gap="gap.large">
                    {/* <Flex.Item shrink={2}>
                        {this.renderTeamAndChannelPulldown}
                    </Flex.Item>
 */}
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
            <EventTable translation={eventTable} dateFormat={dateFormat} dateTimeFormat={dateTimeFormat} events={events} dir={dir} order={order} sort={this.setStateCallBack} loading={loading} filter={filterString} unknownUserDisplayName={unknownUserDisplayName} />
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
            res = <Button onClick={this.loadMoreMessages} content={common.loadMore} />;
        }
        log.info(`▲▲▲ renderContentBottom END ▲▲▲`);
        return res;
    }

    protected setStateCallBack = async (): Promise<void> => {
        this.getEvents();
    };
    
    private onFilterChanged: ComponentEventHandler<InputProps & { value: string; }> = (_: unknown, data): void => {
        log.info(`▼▼▼ onFilterChanged START ▼▼▼`);
        this.setState({ filterInput: data?.value ?? "" }, () => {
            window.clearTimeout(this.filterTimeout);
            this.filterTimeout = window.setTimeout(() => {
                const { order, dir } = (this.state.me as IFindMsgSearchScheduleState).searchResult;
                this.getEvents(order, dir);
            }, 250);
        });
        log.info(`▲▲▲ onFilterChanged END ▲▲▲`);
    }

    private getEvents = async (order: EventOrder = EventOrder.start, dir: OrderByDirection = OrderByDirection.descending): Promise<void> => {
        log.info(`▼▼▼ getEvents START ▼▼▼`);
        const {
            filterInput,
        } = this.state;

        this.setState({ loading: true });

        try {
            const [events, hasMore] = await FindMsgEvent.fetch(order, dir, 0, initialDisplayCount, filterInput);
            log.info(` ★★★ fetched [${events.length}] events from DB ★★★`);
            const me = (this.state.me as IFindMsgSearchScheduleState)
            me.searchResult = { hasMore, events, dir, order }
            this.setState({
                filterString: filterInput,
                me: me,
            });
        } catch (error) {
            AI.trackException({ exception: error });
            this.setError(error, this.state.translation.error.indexedDbReadFailed);
        } finally {
            this.setState({ loading: false });
        }
        log.info(`▲▲▲ getEvents END ▲▲▲`);

    };

    private loadMoreMessages = async () => {
        log.info(`▼▼▼ loadMoreMessages START ▼▼▼`);
        const {
            filterInput,
        } = this.state;
        const searchResult = (this.state.me as IFindMsgSearchScheduleState).searchResult

        try {
            this.setState({ loading: true });

            const [newEvents, hasMore] = await FindMsgEvent.fetch(searchResult.order, searchResult.dir, searchResult.events.length, loadMoreCount, filterInput);

            const me = (this.state.me as IFindMsgSearchScheduleState)
            me.searchResult = {
                events: [...searchResult.events, ...newEvents],
                hasMore: hasMore, 
                order: searchResult.order, 
                dir: searchResult.dir           
            }

            this.setState({
                me: me
            });
        } catch (error) {
            AI.trackException({ exception: error });
            this.setError(error, this.state.translation.error.indexedDbReadFailed);
        } finally {
            this.setState({ loading: false });
        }
        log.info(`▲▲▲ loadMoreMessages END ▲▲▲`);
    }

}