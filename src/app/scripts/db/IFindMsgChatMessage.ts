export interface IFindMsgChatMessage {
    /** ID of the message */
    id: string;

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

    /** message body */
    body: string;

    /** whether body is text or html */
    type: "text" | "html";

    /** Id of the channel the message belongs to */
    chatId: string;

    /** normalized message text for searching. concat (subject, htmlToText(body)) */
    text: string | null;
}
