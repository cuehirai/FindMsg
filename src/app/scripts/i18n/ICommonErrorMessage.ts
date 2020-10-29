export interface ICommonErrorMessage {
    /** IndexedDBのアクセスに失敗しました */
    indexedDbReadFailed: string;
    /** 同期に失敗しました */
    syncFailed: string;
    /** 検索に失敗しました */
    searchFailed: string;
    /** 予期せぬエラーが発生しました */
    internalError: string;
}