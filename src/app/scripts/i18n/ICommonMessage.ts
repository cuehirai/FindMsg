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

    /** {0}を同期中... */
    syncEntity: (entityName: string) => string;
    /** {0}を同期中...{1} */
    syncEntityWithCount: (entityName: string, count: number) => string;
    /** {0}の{1}を同期中... */
    syncSubEntity: (parentName: string, entityName: string) => string;
    /** {0}の{1}を同期中...{2} */
    syncSubEntityWithCount: (parentName: string, entityName: string, count: number) => string;
}