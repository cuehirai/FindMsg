import React from "react";
import { Chip, Tooltip, withStyles } from "@material-ui/core";
import { highlightNode, collapse, empty } from "./highlight";
import { stripHtml } from "./purify";

/**
 * コンテンツコンポーネント用プロパティ
 */
export interface ContentElementProps {
    body: string;
    type: "text" | "html";
    filter: string;
}

export interface IChipsArrayProp {
    names: (string | null)[];
    nodata: string;
}

/**
 * コンテンツコンポーネント（表示しきれない場合に「...」で省略したり検索にヒットした部分をハイライトしたい要素に使用）
 * @param param0 
 */
export const ContentElement: React.FunctionComponent<ContentElementProps> = ({ type, body, filter }: ContentElementProps) => {
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
