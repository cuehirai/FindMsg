import React from "react";
import { Chip, Tooltip, withStyles } from "@material-ui/core";
import { highlightNode, collapse, empty } from "./highlight";
import { collapseConsecutiveChar, sanitize, stripHtml } from "./purify";
import { FindMsgChannel, FindMsgTeam } from "./db";
import * as log from './logger';

/**
 * コンテンツコンポーネント用プロパティ
 */
export interface ContentElementProps {
    body: string;
    type: "text" | "html";
    filter: string;
    tooltip?: boolean;
}

/** ChipsArray生成時の引数用プロパティ */
export interface IChipsArrayProp {
    names: (string | null)[];
    nodata: string;
}

/** チャネルIDから遡ってチャネル名、チーム名を取得するマップ用レコード */
export interface IChannelInfo {
    channelId: string;
    teamId: string;
    channelDisplayName: string;
    teamDisplayName: string;
}

/**
 * コンテンツコンポーネント（表示しきれない場合に「...」で省略したり検索にヒットした部分をハイライトしたい要素に使用）
 * @param param0 
 */
export const ContentElement: React.FunctionComponent<ContentElementProps> = ({ type, body, filter, tooltip }: ContentElementProps) => {
    const el = React.useRef<HTMLSpanElement>(null);
    React.useEffect(() => {
        if (!el.current) return;
        empty(el.current);
        const c = document.createElement("span");
        if (type === "text") {
            c.textContent = body;
        } else {
            // there is no more html, but still entities like &nbsp;
            // c.innerHTML = stripHtml(body);
            c.innerHTML = stripHtml(collapseConsecutiveChar(sanitize(body), "_", 3));
        }
        const hasHighlight = highlightNode(c, [filter, ""]);
        if (body.length > 30 && hasHighlight) collapse(c, 20, 6);
        el.current.appendChild(c);
    }, [type, body, filter, tooltip]);
    let res = <span ref={el} />;
    if (tooltip?? false) {
        res = (
            <HtmlTooltip title={
                <React.Fragment>
                    <span dangerouslySetInnerHTML={{__html: body}} />
                </React.Fragment>
            }>
                <span ref={el} />
            </HtmlTooltip>
        );
    }
    return res;
}

/**
 * 名前の配列からChip配列を生成します。
 * 名前配列に要素がない場合はnodataで指定された文字列を返します。
 * @param param0 
 */
export function ChipsArray({ names, nodata }: IChipsArrayProp): JSX.Element {
    const chipData: Array<JSX.Element> = [];
    if (names.length > 0) {
        names.forEach(name => {
            if (name) {
                chipData.push(<Chip label={name} />);
            }
        });
        return (
            <div className="tooltip">
                {chipData}
            </div>
        );
    } else {
        return (
            <div>
                {nodata}
            </div>
        );
    }
}

/**
 * HTMLを表示するTooltipを生成します。
 */
export const HtmlTooltip = withStyles((theme) => ({
    tooltip: {
        backgroundColor: '#f5f5f9',
        color: 'rgba(0, 0, 0, 0.87)',
        maxWidth: 800,
        fontSize: theme.typography.pxToRem(12),
        border: '1px solid #dadde9',
    },
}))(Tooltip)

/** チャネルIDをキーとしてチャネル名やチーム名を取得できるMAPを生成します（UIでもJSXでもないですが・・） */
export const getChannelMap = async (): Promise<Map<string, IChannelInfo>> => {
    const channelMap = new Map<string, IChannelInfo>();
    const channels = await FindMsgChannel.getAll();
    channels.forEach(async rec => {
        const team = await FindMsgTeam.get(rec.teamId);
        if (team) {
            const channelInfo: IChannelInfo = {
                channelId: rec.id,
                teamId: team.id,
                channelDisplayName: rec.displayName,
                teamDisplayName: team.displayName,
            };
            channelMap.set(rec.id, channelInfo);
        }
    })
    return channelMap;
}


export const getInformation = (showInfo?: boolean): {hasInfo: boolean, info: JSX.Element} => {
    let res = {hasInfo: false, info: (<div/>)};
    try {
        if (showInfo?? true) {
            const request = new XMLHttpRequest();
            const downloadUrl = `https://${location.hostname}/information.txt`
            log.info(`requesting for information: url [${downloadUrl}]`);
            request.open("GET", downloadUrl, false);
            request.send();
            if (request.status < 300) {
                const text = request.response;
                log.info(`information found...content size: [${text.length}]`);
                if (text.length > 0) {
                    res = 
                        {hasInfo: true,
                            info: (<div style={{ width: '100%', maxHeight: 100, overflow: "scroll", 
                                    borderWidth: "medium", borderColor: "lightyellow", borderStyle: "solid"}}>
                                    <span dangerouslySetInnerHTML = {{__html: text.toString("utf-8")}} />
                            </div>)
                        };
                }
            } else {
                log.warn(`failed to read information: (${request.status}) [${request.statusText}]`)
            }
        }
    } catch (e) {
        log.warn(`failed to read information: [${e}]`);
    }
    return res;
}