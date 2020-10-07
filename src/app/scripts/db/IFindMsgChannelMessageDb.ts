export interface IFindMsgChannelMessageDb {
    /** ID of the message */
    id: string;

    /** ID of the parent message (empty string for top level messages) */
    replyToId: string;

    /** timestamp of last reply sync (negative number indicates never) */
    synced: number;

    /** timestamp of creation (negative number means never) */
    created: number;

    /** timestamp of last modification (negative number means never) */
    modified: number;

    /** timestamp of deletion (negative number means never) */
    deleted: number;

    /** max(created, modified, deleted) */
    touched: number;

    /** userId of author */
    author: string;

    /** message subject */
    subject: string | null;

    /** message body */
    body: string;

    /** whether body is text or html */
    type: "text" | "html";

    /** message summary */
    summary: string | null;

    /** internal URL of the message */
    url: string;

    /** Id of the channel the message belongs to */
    channelId: string;

    /** normalized message text for searching. concat (subject, htmlToText(body)) */
    text: string | null;
}
