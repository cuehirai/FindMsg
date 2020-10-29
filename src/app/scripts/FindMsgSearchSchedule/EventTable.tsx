import * as React from "react";

import { Link, TriangleUpIcon, TriangleDownIcon, Table, Loader, TableRowProps, TableCellProps, ShorthandValue, ShorthandCollection, ComponentSlotStyle } from "../ui";

import { IFindMsgEvent } from "../db/Event/IFindMsgEvent";
import * as msTeams from '@microsoft/teams-js';
import { format } from "../dateUtils";
import * as log from '../logger'
import { fixMessageLink } from "../utils";
import { highlightNode, collapse, empty } from "../highlight";
import { stripHtml } from "../purify";
import { EventOrder } from "../db/Event/FindMsgEventEntity";
import { OrderByDirection } from "../db/db-accessor-class-base";
// import { Tooltip } from "office-ui-fabric-react/lib/components/Tooltip/Tooltip";
// import { IFindMsgAttendee } from "../db/Attendee/IFindMsgAttendee";



declare type sortFn = (order: EventOrder, dir: OrderByDirection,) => void;


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

interface ISortableHeaderProps {
    title: string;
    type: EventOrder,
    order: EventOrder,
    dir: OrderByDirection,
    defaultDir: OrderByDirection,
    sort: sortFn;
}


export interface IEventTableTranslation {
    subject: string;
    organizer: string;
    start: string;
    end: string;
    attendees: string;
    body: string;
    allday: string;
    notitle: string;
}


const clickableStyle: ComponentSlotStyle = {
    cursor: "pointer",
};

const emptyStyle: ComponentSlotStyle = {
    'justify-content': 'center',
};


const emptyRows: ShorthandCollection<TableRowProps> = [{ key: 'empty', children: <div>No messages to display</div>, styles: emptyStyle }];
const loadingRows: ShorthandCollection<TableRowProps> = [{ key: 'loading', children: <Loader />, styles: emptyStyle }];


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


interface EventContentProps {
    body: string;
    type: "text" | "html";
    filter: string;
}


const EventContent: React.FunctionComponent<EventContentProps> = ({ type, body, filter }: EventContentProps) => {
    const el = React.useRef<HTMLSpanElement>(null);
    React.useEffect(() => {
        if (!el.current) return;
        empty(el.current);
        const c = document.createElement("span");
        if (type === "text") {
            c.textContent = body;
        } else {
            // there is no more html, but still entities like &nbsp;
            c.innerHTML = stripHtml(body);
        }
        const hasHighlight = highlightNode(c, [filter, ""]);
        if (body.length > 30 && hasHighlight) collapse(c, 20, 6);
        el.current.appendChild(c);
    }, [type, body, filter]);
    return <span ref={el} />
}


/**
 * Table of messages
 * @param props
 */
export const EventTable: React.FunctionComponent<IEventTableProps> = ({ translation, events, order, dir, loading, sort, dateFormat, dateTimeFormat, filter, unknownUserDisplayName }: IEventTableProps) => {
    let rows: ShorthandCollection<TableRowProps>;
    const { allday, notitle } = translation;

    if (events.length === 0) {
        rows = loading ? loadingRows : emptyRows;
    } else {
        const m2dt: (m: Date) => string = m => format(m, dateTimeFormat);
        const title: (s: string | null, n: string) => string = (s, n) => {
            let res = s?? "";
            if (res == "") {
                res = n;
            }
            return res;
        };
        const stContent: (s: Date, e: Date, a: boolean) => string = (s, e, a) => {
            let res: string = m2dt(s) + " ~ " + m2dt(e);
            if (a) {
                res = format(s, dateFormat) + " " + allday;
            }
            return res;
        };
        const organizer: (n: string | null, m: string | null, u:string) => string = (n, m, u) => {
            const res = n?? m?? u;
            return res;
        };
        // const tooltip: (a: IFindMsgAttendee[]) => string = (a) => {
        //     const arr: Array<string> = [];
        //     a.forEach(rec => {
        //         if (rec.name) {
        //             arr.push(rec.name);
        //         }
        //     });
        //     const res = translation.attendees + ": " + arr.join(", ");
        //     log.info(`tooltip for attendee created: [${res}]`)
        //     return res;
        // }
        const EventTableRow: (msg: IFindMsgEvent) => TableRowProps = ({ id, subject, organizerName, organizerMail, start, end, isAllDay, body, type, webLink }) => ({
            key: id,
            items: [
                { key: 's', truncateContent: true, content: <Link onClick={() => msTeams.executeDeepLink(fixMessageLink(webLink), log.info)} disabled={!webLink}><EventContent body={title(subject, notitle)} type="text" filter={filter} /></Link> },
                { key: 'o', truncateContent: true, content: organizer(organizerName, organizerMail, unknownUserDisplayName)},
                { key: 't', truncateContent: false, content: stContent(start, end, isAllDay) },
                { key: 'c', truncateContent: true, content: <EventContent body={body} type={type} filter={filter} /> }
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


