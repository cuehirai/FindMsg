import { ISyncWidgetTranslation } from "../SyncWidget";
import { IMessageTableTranslation } from "../FindMsgTopicsTab/MessageTable";
import { ITopicsTabTranslation, ISearchTabTranslation, IFindMsgTopicsTabConfigTranslation, IChatSearchTabTranslation, IFindMsgIndexTranslation } from "../client";
import { ISyncProgressTranslation } from "../db";
import { IAuthMessages } from "../auth/IAuthMessages";
import { ICommonErrorMessage } from "./ICommonErrorMessage";
import { IStoragePermissionWidgetTranslation } from "../StoragePermissionWidget";
import { ICommonMessage } from "./ICommonMessage";
import { IEntityNames } from "../db/Database";
import { IFindMsgScheduleTranslation } from "../FindMsgSearchSchedule/FindMsgSearchSchedule";
import { IEventTableTranslation } from "../FindMsgSearchSchedule/EventTable";

export interface IMessageTranslation {
    dateFormat: string;
    dateTimeFormat: string;

    footer: string;
    filter: string;
    showCollapsed: string;
    showExpanded: string;
    unknownUserDisplayName: string;

    //共通系メッセージインターフェース
    common: ICommonMessage;
    auth: IAuthMessages;
    sync: ISyncWidgetTranslation;
    syncProgress: ISyncProgressTranslation;
    storagePermission: IStoragePermissionWidgetTranslation;
    error: ICommonErrorMessage;

    //エンティティ名
    entities: IEntityNames;

    //ページ固有メッセージインターフェース
    search: ISearchTabTranslation;
    topics: ITopicsTabTranslation;
    topicsConfig: IFindMsgTopicsTabConfigTranslation;
    table: IMessageTableTranslation;
    chatSearch: IChatSearchTabTranslation;
    schedule: IFindMsgScheduleTranslation;
    eventTable: IEventTableTranslation;
    about: IFindMsgIndexTranslation;

}
