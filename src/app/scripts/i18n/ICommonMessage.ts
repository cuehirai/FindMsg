export interface ICommonMessage {
    /** チーム */
    team: string;
    /** チャネル */
    channel: string;
    /** 続きを表示 */
    loadMore: string;
    /** （全チーム） */
    allTeams: string;
    /** （全チャネル） */
    allChannels: string;
    /** チーム名とチャネル名 */
    teamchannel: (teamname: string, channelname: string) => string;

    /** から */
    from: string;
    /** まで */
    to: string;
    /** 検索結果: {0}} 件({1}件表示中) */
    messagesFound: (shown: number, total: number) => string;
    /** 検索 */
    search: string;
    /** 検索中 */
    searching: string;
    /** 中止 */
    cancel: string;

    /** 期間を限定しない */
    searchTimeAll: string;
    /** 1週間以内 */
    searchTimePastWeek: string;
    /** 1ヶ月以内 */
    searchTimePastMonth: string;
    /** 1年以内 */
    searchTimePastYear: string;
    /** 指定の期間内 */
    searchTimeCustom: string;
    /** (すべて) */
    noSelection: string;

    /** {0}を同期中... */
    syncEntity: (entityName: string) => string;
    /** {0}を同期中...{1} */
    syncEntityWithCount: (entityName: string, count: number) => string;
    /** {0}の{1}を同期中... */
    syncSubEntity: (parentName: string, entityName: string) => string;
    /** {0}の{1}を同期中...{2} */
    syncSubEntityWithCount: (parentName: string, entityName: string, count: number) => string;

    /** 最新データをエクスポートしますか？ この処理は数分かかる可能性があります。 */
    confirmExport: string;
    /** (画像エクスポートを省略すると時間が短縮されます。) */
    confirmExportOption: string;
    /** エクスポート中... ( {0} / {1} ) {2}% 完了 */
    exportProgress: (done: number, all: number, progress: number) => string;
    /** データベースにデータをインポートしますか？ */
    confirmImportForNewUser: string;
    /** One Driveに現在のデータよりも新しいファイルがエクスポートされています。インポートしますか？ */
    confirmImportNewerData: string;
    /** インポート中... ( {0} / {1} ) {2}% 完了 */
    importProgress: (done: number, all: number, progress: number) => string;
    /** アプリの終了やタブの移動をしないでください。 */
    exportImportMessage: string;
    /** 処理を待機しています... */
    standingBy: string;
    /** はい */
    yes: string;
    /** いいえ */
    no: string;
    /** 画像をエクスポート */
    exportImages: string;
    /** OneDriveに十分な空き容量がありません。最低でもあと${short}KBの空きが必要です。※ゴミ箱を空にする、AppData内のファイルの履歴を削除する、などの方法が効果的です。 */
    oneDriveQuotaShorts: (short: number) => string;
}