import { IFindMsgAttendee } from "../Attendee/IFindMsgAttendee";
import { ITeamsEntityBase } from "../db-accessor-class-base";

/** Eventのプロパティ */
export interface IFindMsgEvent extends ITeamsEntityBase {
    /** 作成日 */
    created: Date;
    
    /** 変更日 */
    modified: Date;

    /** 開催者名 */
    organizerName: string | null;

    /** 開催者メールアドレス */
    organizerMail: string | null;

    /** 開始日時 */
    start: Date;

    /** 終了日時 */
    end: Date;

    /** タイトル */
    subject: string | null;

    /** 会議詳細 */
    body: string;

    /** 会議詳細(body)のタイプ */
    type: "text" | "html";

    /** 添付有無 */
    hasAttachments: boolean;

    /** 重要度 */
    importance: "low" | "normal" | "high";

    /** 機密度 */
    sensitivity: "normal" | "personal" | "private" | "confidential";

    /** 終日 */
    isAllDay: boolean;

    /** キャンセル */
    isCancelled: boolean;

    /** Eventへのリンク */
    webLink: string;

    /** 検索(絞り込み)用のテキスト※subject ＋ htmlToText(body) */
    text: string | null;

    /** 参加者 */
    attendees: IFindMsgAttendee[];
}