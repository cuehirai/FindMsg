export interface IFindMsgChannelMessage {
    /** ID of the message */
    id: string;

    /** ID of the parent message (empty string for top level messages) */
    replyToId: string;

    /** timestamp of last reply sync (negative number indicates never) */
    synced: Date;

    /** timestamp of creation (negative number means never) */
    created: Date;

    /** timestamp of last modification (negative number means never) */
    modified: Date;

    /** timestamp of deletion (negative number means never) */
    deleted: Date;

    /** userId of author */
    authorId: string;

    /** displayName of the author */
    authorName: string;

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
