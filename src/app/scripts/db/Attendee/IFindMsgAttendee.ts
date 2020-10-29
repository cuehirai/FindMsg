import { ITeamsEntityBase } from "../db-accessor-class-base";

/** Event.Attendeeのプロパティ */
export interface IFindMsgAttendee extends ITeamsEntityBase {
    /** イベントのID */
    eventId: string;

    /** 開催者かどうか */
    isOrganizer: boolean;

    /** 参加者名※Teamsメンバーの名前とは限らない */
    name: string | null;

    /** メールアドレス※Teamsのログインヒントであるとは限らない */
    mail: string | null;

    /** 参加者のタイプ */
    type: "required"|"optional"|"resource";

    /** 応答ステータス */
    status: "none"|"organizer"|"tentativelyAccepted"|"accepted"|"declined"|"notResponded";
}