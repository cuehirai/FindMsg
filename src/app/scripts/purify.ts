import DOMPurify from 'dompurify';

const stripAll = {
    ALLOWED_TAGS: [],
    KEEP_CONTENT: true,
}


export const whiteSpaceRegex = /[\s]+/g.compile();


export function sanitize(html: string): string {
    return DOMPurify.sanitize(html);
}


export function stripHtml(html: string): string {
    return DOMPurify.sanitize(html, stripAll);
}

/** すべての「1文字の区切り文字※」をスペースに置き換えます。
 * ※空白文字( )、改ページ(\f)、改行(\n)、ラインフィード(\r)、タブ文字(\r)、垂直タブ(\v)、
 * No-break space(\u00a0)、Ogham space mark(\u1680)、Mongolian vowel separator(\u180e)、
 * Xxx Quad(\u2000-\u2001)、Xxx Space(\u2002-\u200a)、Line separator(\u2028)、
 * Paragraph separator(\u2029)、Narrow no-break space(\u202f)、
 * Medium mathematical space(\u205f)、全角スペース(\u3000)、BOM(\ufeff) いずれかの 1文字
 */
export function collapseWhitespace(text: string) : string {
    return text.replace(whiteSpaceRegex, " ");
}

/**
 * 指定文字が指定数以上連続している場合に省略した文字列に置き換えます。
 * 例）target: "_"、number: 3 の時、「____________________」→「__...」
 * @param text 変換したい文字列データ
 * @param target 変換対象の文字
 * @param number 連続する数
 */
export function collapseConsecutiveChar(text: string, target: string, number: number) : string {
    const reg = new RegExp(`${target}{${number},}`, "g");
    const res = text.replace(reg, target + target + "...");
    return res;
}