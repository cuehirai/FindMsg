import { IDbEntityBase } from "../db-accessor-class-base";

/** Eventのプロパティ(DB用) */
export interface IFindMsgEventDb extends IDbEntityBase {
    /** 作成日 */
    created: number;
    
    /** 変更日 */
    modified: number;

    /** 開催者名 */
    organizerName: string | null;

    /** 開催者メールアドレス */
    organizerMail: string | null;

    /** 開始日時 */
    start: number;

    /** 終了日時 */
    end: number;

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
}