import * as React from "react";

import { Link, TriangleUpIcon, TriangleDownIcon, Table, Loader, TableRowProps, TableCellProps, ShorthandValue, ShorthandCollection, ComponentSlotStyle } from "../ui";

import { MessageOrder, Direction, IFindMsgChannelMessage } from "../db";
import * as msTeams from '@microsoft/teams-js';
import { format, isValid } from "../dateUtils";
import { info } from '../logger'
import { fixMessageLink } from "../utils";
// import { highlightNode, collapse, empty } from "../highlight";
// import { stripHtml } from "../purify";
import { ContentElement, HtmlTooltip, IChannelInfo } from "../ui-jsx";
import { Typography } from "@material-ui/core";


declare type sortFn = (order: MessageOrder, dir: Direction,) => void;


export interface IMessageTableProps {
    messages: IFindMsgChannelMessage[];
    order: MessageOrder;
    dir: Direction;
    loading: boolean;
    sort: sortFn;
    t: IMessageTableTranslation;
    dateFormat: string;
    filter: string;
    unknownUserDisplayName: string;
    channelMap: Map<string, IChannelInfo>;
    teamchannel: (teamname: string, channelname: string) => string;
}


interface ISortableHeaderProps {
    title: string;
    type: MessageOrder,
    order: MessageOrder,
    dir: Direction,
    defaultDir: Direction,
    sort: sortFn;
}


export interface IMessageTableTranslation {
    subject: string;
    author: string;
    dateTime: string;
    body: string;
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
        if (dir === Direction.ascending) {
            fn = () => sort(type, Direction.descending);
            indicator = <TriangleUpIcon />;
        } else {
            fn = () => sort(type, Direction.ascending);
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


// interface MessageContentProps {
//     body: string;
//     type: "text" | "html";
//     filter: string;
// }

interface ITeamChannelTooltipArg {
    authorName: string;
    unknownUserDisplayName: string;
    channelId: string;
    channelMap: Map<string, IChannelInfo>;
    teamchannel: (teamname: string, channelname: string) => string;
}

// const MessageContent: React.FunctionComponent<MessageContentProps> = ({ type, body, filter }: MessageContentProps) => {
//     const el = React.useRef<HTMLSpanElement>(null);
//     React.useEffect(() => {
//         if (!el.current) return;
//         empty(el.current);
//         const c = document.createElement("span");
//         if (type === "text") {
//             c.textContent = body;
//         } else {
//             // there is no more html, but still entities like &nbsp;
//             c.innerHTML = stripHtml(body);
//         }
//         const hasHighlight = highlightNode(c, [filter, ""]);
//         if (body.length > 30 && hasHighlight) collapse(c, 20, 6);
//         el.current.appendChild(c);
//     }, [type, body, filter]);
//     return <span ref={el} />
// }


/**
 * Table of messages
 * @param props
 */
export const MessageTable: React.FunctionComponent<IMessageTableProps> = ({ t, messages, order, dir, loading, sort, dateFormat, filter, unknownUserDisplayName, channelMap, teamchannel }: IMessageTableProps) => {
    let rows: ShorthandCollection<TableRowProps>;

    if (messages.length === 0) {
        rows = loading ? loadingRows : emptyRows;
    } else {
        const m2dt: (m: Date) => string = m => format(m, dateFormat);
        const MessageTableRow: (msg: IFindMsgChannelMessage) => TableRowProps = ({ id, subject, authorName, created, modified, body, type, url, channelId }) => ({
            key: id,
            items: [
                { key: 's', truncateContent: true, content: <Link onClick={() => msTeams.executeDeepLink(fixMessageLink(url), info)} disabled={!url}><ContentElement body={subject ?? ""} type="text" filter={filter} /></Link> },
                // { key: 'a', truncateContent: true, content: authorName || unknownUserDisplayName },
                { key: 'a', truncateContent: true, content: teamchannelTooltip({authorName, unknownUserDisplayName, channelId, channelMap, teamchannel}) },
                { key: 't', truncateContent: false, content: m2dt(isValid(modified) ? modified : created) },
                { key: 'c', truncateContent: true, content: <ContentElement body={body} type={type} filter={filter} tooltip={true} /> }
            ],
        });

        rows = messages.map(MessageTableRow);
    }

    const header: ShorthandValue<TableRowProps> = {
        header: true,
        items: [
            SortableHeader({ title: t.subject, type: MessageOrder.subject, defaultDir: Direction.ascending, dir, order, sort }),
            SortableHeader({ title: t.author, type: MessageOrder.author, defaultDir: Direction.ascending, dir, order, sort }),
            SortableHeader({ title: t.dateTime, type: MessageOrder.touched, defaultDir: Direction.descending, dir, order, sort }),
            { key: t.body, content: t.body, styles: { cursor: "default" } },
        ]
    };

    return <Table className="messageTable" header={header} rows={rows} />;
};

const teamchannelTooltip = (params:ITeamChannelTooltipArg): JSX.Element => {
    const channelInfo = params.channelMap.get(params.channelId);
    const name = channelInfo? params.teamchannel(channelInfo.teamDisplayName, channelInfo.channelDisplayName) : null;
    if (name) {
        return (
            <HtmlTooltip
                title={
                    <React.Fragment>
                    <Typography color="inherit">{name}</Typography>
                    </React.Fragment>
                }
            >
                <div>{params.authorName || params.unknownUserDisplayName}</div>
            </HtmlTooltip>
        );
    } else {
        return (<div>{params.authorName || params.unknownUserDisplayName}</div>);
    }
}
