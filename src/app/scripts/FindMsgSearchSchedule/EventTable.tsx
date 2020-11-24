import * as React from "react";
import { Button, Typography } from "@material-ui/core";
import * as msTeams from '@microsoft/teams-js';

import { Link, TriangleUpIcon, TriangleDownIcon, Table, Loader, TableRowProps, TableCellProps, ShorthandValue, ShorthandCollection, ComponentSlotStyle } from "../ui";
import { ChipsArray, ContentElement, HtmlTooltip } from "../ui-jsx";
import { fixMessageLink } from "../utils";
import { format } from "../dateUtils";
import * as log from '../logger'

import { OrderByDirection } from "../db/db-accessor-class-base";
import { EventOrder } from "../db/Event/FindMsgEventEntity";
import { IFindMsgEvent } from "../db/Event/IFindMsgEvent";
import { IFindMsgAttendee } from "../db/Attendee/IFindMsgAttendee";

/**
 * ヘッダのクリックでソートする際のメソッドタイプ
 * ※ソート項目がテーブル固有とするため、検索結果のコンポーネントごとに定義する必要があります。
 */
declare type sortFn = (order: EventOrder, dir: OrderByDirection,) => void;

/**
 * イベント検索結果コンポーネント用のプロパティ
 */
export interface IEventTableProps {
    events: IFindMsgEvent[];
    order: EventOrder;
    dir: OrderByDirection;
    loading: boolean;
    sort: sortFn;
    translation: IEventTableTranslation;
    dateFormat: string;
    dateTimeFormat: string;
    filter: string;
    unknownUserDisplayName: string;
}

/**
 * クリックによりソート可能なヘッダコンポーネント用のプロパティ
 * ※ソート項目がテーブル固有のため共通化できそうでできません。
 */
interface ISortableHeaderProps {
    title: string;
    type: EventOrder,
    order: EventOrder,
    dir: OrderByDirection,
    defaultDir: OrderByDirection,
    sort: sortFn;
}

/**
 * この検索結果コンポーネントで使用するロケール依存リソース
 */
export interface IEventTableTranslation {
    subject: string;
    organizer: string;
    start: string;
    end: string;
    attendees: string;
    body: string;
    allday: string;
    notitle: string;
    noattendee: string;
}

/**
 * クリック可能コンポーネントスタイルでマウスポインタの形状を設定
 */
const clickableStyle: ComponentSlotStyle = {
    cursor: "pointer",
};

/**
 * 行コンポーネントスタイルで行数なしの場合の表示スタイルを設定
 */
const emptyStyle: ComponentSlotStyle = {
    'justify-content': 'center',
};

/**
 * 「行なし」行コンポーネント
 */
const emptyRows: ShorthandCollection<TableRowProps> = [{ key: 'empty', children: <div>No events to display</div>, styles: emptyStyle }];

/**
 * 「loading」行コンポーネント
 */
const loadingRows: ShorthandCollection<TableRowProps> = [{ key: 'loading', children: <Loader />, styles: emptyStyle }];

/**
 * クリックでソート可能なヘッダコンポーネント
 * @param param0 
 */
const SortableHeader: (props: ISortableHeaderProps) => TableCellProps = ({ title, type, order, dir, defaultDir, sort }: ISortableHeaderProps) => {
    let fn: () => void;
    let indicator: JSX.Element | null;

    if (type === order) {
        if (dir === OrderByDirection.ascending) {
            fn = () => sort(type, OrderByDirection.descending);
            indicator = <TriangleUpIcon />;
        } else {
            fn = () => sort(type, OrderByDirection.ascending);
            indicator = <TriangleDownIcon />;
        }
    } else {
        fn = () => sort(type, defaultDir);
        indicator = null;
    }

    return {
        content: <span>{title}{indicator}</span>,
        key: title,
        onClick: fn,
        styles: clickableStyle,
    };
};

/**
 * イベント検索結果表示コンポーネント
 * @param props
 */
export const EventTable: React.FunctionComponent<IEventTableProps> = ({ translation, events, order, dir, loading, sort, dateFormat, dateTimeFormat, filter, unknownUserDisplayName }: IEventTableProps) => {
    let rows: ShorthandCollection<TableRowProps>;
    const { allday, notitle } = translation;

    if (events.length === 0) {
        rows = loading ? loadingRows : emptyRows;
    } else {
        /** 日時に書式を設定 */
        const m2dt: (m: Date) => string = m => format(m, dateTimeFormat);
        /** 件名を生成（件名の設定がない場合に「（件名なし）」を表示） */
        const title: (s: string | null, n: string) => string = (s, n) => {
            let res = s?? "";
            if (res == "") {
                res = n;
            }
            return res;
        };
        /** 開始日時を生成（通常は「開始～終了」、終日イベントの場合は「開始（終日）」を表示 */
        const stContent: (s: Date, e: Date, a: boolean) => string = (s, e, a) => {
            let res: string = m2dt(s) + " ~ " + m2dt(e);
            if (a) {
                res = format(s, dateFormat) + " " + allday;
            }
            return res;
        };
        /** 参加者Tooltipを生成 */
        const organizerWithAttendeeTooltip : (name: string | null, mail: string | null, unknown:string, attendees: IFindMsgAttendee[])
         => JSX.Element = (name, mail, unknown, attendees) => {
            const organizer = name?? mail?? unknown;
            const names = attendees.map(rec => (rec.name));
            const nodata = translation.noattendee;
            return (
              <div>
                <HtmlTooltip
                  title={
                    <React.Fragment>
                      <Typography color="inherit">{translation.attendees}</Typography>
                      {ChipsArray({names, nodata})}
                    </React.Fragment>
                  }
                >
                  <Button>{organizer}</Button>
                </HtmlTooltip>
              </div>
            );
        };
        /** deeplinkのジャンプ先を生成 */
        const deeplink : (eventId: string) => string = (eventId) => {
            const link = `https://teams.microsoft.com/_#/scheduling-form/?eventId=${eventId}&opener=1&providerType=0`;
            return link;
        };

        const EventTableRow: (event: IFindMsgEvent) => TableRowProps = ({ id, subject, organizerName, organizerMail, attendees, start, end, isAllDay, body, type, webLink }) => ({
            key: id,
            items: [
                { key: 's', truncateContent: true, content: <Link onClick={() => msTeams.executeDeepLink(fixMessageLink(deeplink(id)), log.info)} disabled={!webLink}><ContentElement body={title(subject, notitle)} type="text" filter={filter} /></Link> },
                { key: 'o', truncateContent: true, content: organizerWithAttendeeTooltip(organizerName, organizerMail, unknownUserDisplayName, attendees) },
                { key: 't', truncateContent: false, content: stContent(start, end, isAllDay) },
                { key: 'c', truncateContent: true, content: <ContentElement body={body} type={type} filter={filter} tooltip={true}/> }
            ],
        });

        rows = events.map(EventTableRow);
    }

    const header: ShorthandValue<TableRowProps> = {
        header: true,
        items: [
            SortableHeader({ title: translation.subject, type: EventOrder.subject, defaultDir: OrderByDirection.ascending, dir, order, sort }),
            SortableHeader({ title: translation.organizer, type: EventOrder.organizer, defaultDir: OrderByDirection.ascending, dir, order, sort }),
            SortableHeader({ title: translation.start, type: EventOrder.start, defaultDir: OrderByDirection.descending, dir, order, sort }),
            { key: translation.body, content: translation.body, styles: { cursor: "default" } },
        ]
    };

    return <Table className="eventTable" header={header} rows={rows} />;
};


