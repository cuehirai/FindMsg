/* eslint-disable @typescript-eslint/explicit-module-boundary-types */
import { IMessageTranslation } from "./IMessageTranslation";
import * as du from "../dateUtils";

const dateFormat = "yyyy/MM/dd";
const dateTimeFormat = "yyyy/MM/dd HH:mm";

const appName = "メッセージ検索";

export const messages: IMessageTranslation = {
    dateFormat,
    dateTimeFormat,
    footer: "(C) Copyright Kacoms",

    filter: "絞り込む",
    showCollapsed: "簡易表示",
    showExpanded: "詳細表示",
    unknownUserDisplayName: "（不明）",

    common: {
        team: "チーム",
        channel: "チャネル",
        loadMore: "続きを表示",
        allTeams: "（全チーム）",
        allChannels: "（全チャネル）",
        teamchannel: (teamname, channelname) => `チーム: ${teamname} / チャネル: ${channelname}`,
        from: "から",
        to: "まで",
        messagesFound: (shown, total) => `検索結果: ${total} 件` + (shown === total ? "" : `(${shown}件表示中)`),
        search: "検索",
        searching: "検索中",
        cancel: "中止",
        searchTimeAll: "期間を限定しない",
        searchTimePastWeek: "1週間以内",
        searchTimePastMonth: "1ヶ月以内",
        searchTimePastYear: "1年以内",
        searchTimeCustom: "指定の期間内",
        noSelection: "(すべて)",
        syncEntity: entityName => `[${entityName}]を同期中...`,
        syncEntityWithCount: (entityName: string, count: number) => `[${entityName}]を同期中... ${count}`,
        syncSubEntity: (parentName: string, entityName: string) => `[${parentName}]の[${entityName}]を同期中...`,
        syncSubEntityWithCount: (parentName: string, entityName: string, count: number) => `[${parentName}]の[${entityName}]を同期中... ${count}`,
    },

    entities: {
        teams: "チーム一覧",
        channels: "チャンネル一覧",
        messages: "チャネルメッセージ",
        users: "ユーザ一覧",
        chats: "個人チャット",
        chatMembers: "チャットメンバ",
        images: "画像",
        events: "スケジュール",
        attendees: "参加者",
    },

    auth: {
        loginButtonText: "ログイン",
        adminLoginButtonText: "管理者としてログイン",
        unkownError: "ログインが失敗しました。",
        loginDialogHeader: "ログインしてください",
        loginMessage: "このアプリを使うために、Teamsに使用しているマイクロソフトアカウントでログインしてください。",
        needServerInteraction: "ログインボタンをクリックして、Teamsで使用しているマイクロソフトアカウントでログインしてください。",
        needConsent: "ユーザまたは管理者はこのアプリの使用に承諾していません。",
        serverError: "ログインサーバに接続できません。数分後もう一度試してください。",
    },

    sync: {
        cancel: "中止",
        lastSynced: d => du.isValid(d) ? `最後の同期: ${du.format(d, dateTimeFormat)}` : "同期されていません",
        syncNowButton: "今すぐ同期",
        syncing: "同期中",
        cancelWait: "中止待ち",
    },

    syncProgress: {
        teamList: "チーム一覧を同期",
        channelList: t => `[${t}]のチャンネル一覧を同期`,
        topLevelMessages: (c, n) => `[${c}]の投稿を同期... ${n}`,
        replies: (c, n) => `[${c}]の返信を同期... ${n}`,
        syncProblem: "同期中に問題が発生しました。一部のメッセージを取得できなかった可能性があります。数分待ってから再び同期してください。",
        chatList: "チャット一覧を同期",
        chatMessages: (c, n) => `[${c}]の投稿を同期... ${n}`,
    },

    error: {
        indexedDbReadFailed: "IndexedDBのアクセスに失敗しました",
        searchFailed: "検索に失敗しました",
        syncFailed: "同期に失敗しました",
        internalError: "予期せぬエラーが発生しました",
    },

    storagePermission: {
        grantTitle: "ストレージ権限を許可してください",
        grantMessage: "ストレージの権限が付与されていません。メッセージをこの端末で保存するために、ストレージ権限が必要です。",
        linkInside: "権限を許可する",
        linkOutside: "権限を新しいタブで許可する",
    },

    topics: {
        pageTitle: `${appName} - 表題一覧`,
        // team: "チーム",
        // channel: "チャンネル",
        // loadMore: "続きを表示",
        // allTeams: "（全チーム）",
        // allChannels: "（全チャネル）",
    },

    topicsConfig: {
        loading: "準備中。しばらくお待ちください。",
        errorNoChannelId: "チャネルIDの取得に失敗しました。",
        errorNoGroupId: "チームIDの取得に失敗しました。",
        errorNotInTeams: "Microsoft Teamsとの通信に失敗しました。",
        errorPrivateChannel: "このタブをプライベートチャンネルに追加できません。",
        headerConfigure: "タブ設定",
        labelTabName: "タブ名",
        placeholderTabName: "タブの名前の入力してください",
        defaultTabName: "件名一覧",
    },

    search: {
        pageTitle: `${appName} - チャネルメッセージ検索`,
        header: "チャットメッセージ検索",
        // allTeams: "すべてのチームとチャンネルを検索",
        // from: "から",
        // to: "まで",
        // messagesFound: (shown, total) => `検索結果: ${total} 件` + (shown === total ? "" : `(${shown}件表示中)`),
        // search: "検索",
        // searching: "検索中",
        // cancel: "中止",
        // searchTimeAll: "期間を限定しない",
        // searchTimePastWeek: "1週間以内",
        // searchTimePastMonth: "1ヶ月以内",
        // searchTimePastYear: "1年以内",
        // searchTimeCustom: "指定の期間内",
        searchUsersLabel: "このユーザのメッセージのみ",
        searchUsersPlaceholder: "(すべて)",
    },

    chatSearch: {
        pageTitle: `${appName} - チャットメッセージ検索`,
        header: "個人チャットメッセージ検索",
        allChats: "すべてのチャットを検索",
        // from: "から",
        // to: "まで",
        // messagesFound: (shown, total) => `検索結果: ${total} 件` + (shown === total ? "" : `(${shown}件表示中)`),
        // search: "検索",
        // searching: "検索中",
        // cancel: "中止",
        // searchTimeAll: "期間を限定しない",
        // searchTimePastWeek: "1週間以内",
        // searchTimePastMonth: "1ヶ月以内",
        // searchTimePastYear: "1年以内",
        // searchTimeCustom: "指定の期間内",
        searchUsersLabel: "このユーザのメッセージのみ",
        searchUsersPlaceholder: "(すべて)",
    },

    table: {
        subject: "件名",
        author: "発信者",
        dateTime: "日時",
        body: "本文",
    },

    schedule: {
        pageTitle: `${appName} - スケジュール一覧`,
        filterByStart: "開始日で絞り込む",
        filterByOrganizer: "主催者で絞り込む",
    },

    eventTable: {
        subject: "件名",
        organizer: "主催者",
        start: "開始",
        end: "終了",
        attendees: "参加者",
        body: "本文",
        allday: "(終日)",
        notitle: "(件名なし)",
        noattendee: "(なし)",
        },
}