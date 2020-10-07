export interface IFindMsgChatMessageDb {
    /** ID of the message */
    id: string;

    /** timestamp of creation (negative number means never) */
    created: number;

    /** timestamp of last modification (negative number means never) */
    modified: number;

    /** timestamp of deletion (negative number means never) */
    deleted: number;

    /** userId of author */
    authorId: string;

    /** message body */
    body: string;

    /** whether body is text or html */
    type: "text" | "html";

    /** Id of the channel the message belongs to */
    chatId: string;

    /** normalized message text for searching. concat (subject, htmlToText(body)) */
    text: string | null;
}
