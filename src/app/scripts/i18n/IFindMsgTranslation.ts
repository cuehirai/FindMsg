import { ISyncWidgetTranslation } from "../SyncWidget";
import { IMessageTableTranslation } from "../FindMsgTopicsTab/MessageTable";
import { ITopicsTabTranslation, ISearchTabTranslation, IFindMsgTopicsTabConfigTranslation, IChatSearchTabTranslation } from "../client";
import { ISyncProgressTranslation } from "../db";
import { IAuthMessages } from "../auth/IAuthMessages";
import { IFindMsgError } from "./IFindMsgError";
import { IStoragePermissionWidgetTranslation } from "../StoragePermissionWidget";

export interface IFindMsgTranslation {
    dateFormat: string;
    dateTimeFormat: string;

    footer: string;
    filter: string;
    showCollapsed: string;
    showExpanded: string;
    unknownUserDisplayName: string;

    topics: ITopicsTabTranslation;
    topicsConfig: IFindMsgTopicsTabConfigTranslation;
    search: ISearchTabTranslation;
    sync: ISyncWidgetTranslation;
    syncProgress: ISyncProgressTranslation;
    table: IMessageTableTranslation;
    auth: IAuthMessages;
    storagePermission: IStoragePermissionWidgetTranslation;

    chatSearch: IChatSearchTabTranslation;

    error: IFindMsgError;
}
