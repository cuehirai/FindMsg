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


export function collapseWhitespace(text: string) : string {
    return text.replace(whiteSpaceRegex, " ");
}