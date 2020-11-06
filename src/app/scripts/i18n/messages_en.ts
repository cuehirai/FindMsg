/* eslint-disable @typescript-eslint/explicit-module-boundary-types */
import { IMessageTranslation } from "./IMessageTranslation";
import * as du from "../dateUtils";
const dateFormat = "yyyy/MM/dd";
const dateTimeFormat = "yyyy/MM/dd HH:mm";

const appName = "KSearch";

export const messages: IMessageTranslation = {
    dateFormat,
    dateTimeFormat,

    footer: "(C) Copyright Kacoms",

    filter: "Filter:",
    showCollapsed: "See less",
    showExpanded: "See more",
    unknownUserDisplayName: "(unknown)",

    common: {
        team: "team",
        channel: "channel",
        loadMore: "Load more",
        allTeams: "(All Teams)",
        allChannels: "(All Channels)",
        teamchannel: (teamname, channelname) => `team: ${teamname} / channel: ${channelname}`,
        from: "between",
        to: "and",
        messagesFound: (shown, total) => `${total} ${total === 1 ? "message" : "messages"} found ${total === shown ? "" : `- ${shown} displayed.`}`,
        search: "search",
        searching: "searching",
        cancel: "cancel",
        searchTimeAll: "All time",
        searchTimePastWeek: "past week",
        searchTimePastMonth: "past month",
        searchTimePastYear: "past year",
        searchTimeCustom: "custom",
        noSelection: "(All)",
        syncEntity: entityName => `Syncing [${entityName}]...`,
        syncEntityWithCount: (entityName: string, count: number) => `Syncing [${entityName}]... ${count}`,
        syncSubEntity: (parentName: string, entityName: string) => `Syncing [${entityName}] of [${parentName}]...`,
        syncSubEntityWithCount: (parentName: string, entityName: string, count: number) => `Syncing [${entityName}] of [${parentName}]... ${count}`,
    },

    entities: {
        teams: "team list",
        channels: "channel list",
        messages: "messages",
        users: "user list",
        chats: "chat",
        chatMembers: "chat member",
        images: "image",
        events: "schedule",
        attendees: "attendee",
    },

    auth: {
        loginButtonText: "Login",
        adminLoginButtonText: "Login as admin",
        loginDialogHeader: "Please log in",
        loginMessage: "Please log in with the account you use for Microsoft Teams to use this app.",
        needServerInteraction: "Please click the login button to login with your microsoft account.",
        unkownError: "Login failed. Reload the app to try again.",
        needConsent: "The user or administrator has not consented to use the application",
        serverError: "Could not connect to the login server. Please try again in a few minutes.",
    },

    sync: {
        cancel: "cancel",
        cancelWait: "cancelling...",
        lastSynced: d => du.isValid(d) ? `last synced: ${du.format(d, dateTimeFormat)}` : "never synced",
        syncNowButton: "sync now",
        syncing: "syncing",
    },

    syncProgress: {
        teamList: "Sync team list",
        channelList: t => `Syncing channel list of [${t}]`,
        topLevelMessages: (c, n) => `Syncing top level messages of [${c}]... ${n}`,
        replies: (c, n) => `Syncing message replies of [${c}]... ${n}`,
        syncProblem: "A problem occured during sync. Some messages may be missing or outdated. Please try to sync again in a few minutes.",
        chatList: "Sync chat list",
        chatMessages: (c, n) => `Syncing chat messages of [${c}]... ${n}`,
    },

    error: {
        indexedDbReadFailed: "Failed to access IndexedDB",
        searchFailed: "Search failed",
        syncFailed: "Sync failed",
        internalError: "This app has encountered an internal error",
    },

    storagePermission: {
        grantTitle: "Please grant storage permission",
        grantMessage: "This app has not been granted permission to use persistent storage. This permission is needed to ensure that messages can be stored on this computer.",
        linkInside: "Click here to grant permission",
        linkOutside: "Click here to grant permission in new window",
    },

    topics: {
        pageTitle: `${appName} - Topics`,
        // team: "team",
        // channel: "channel",
        // loadMore: "Load more",
        // allTeams: "(All Teams)",
        // allChannels: "(All Channels)",
    },

    topicsConfig: {
        loading: "Loading, please wait.",
        errorNoChannelId: "Error: Could not retrieve channel ID",
        errorNoGroupId: "Error: Could not retrieve team ID",
        errorNotInTeams: "Error: not inside Microsoft Teams",
        errorPrivateChannel: "Error: Can not add this tab to a private channel",
        headerConfigure: "Configure your tab",
        labelTabName: "Tab name",
        placeholderTabName: "Enter a name for the tab here",
        defaultTabName: "Topics",
    },

    search: {
        pageTitle: `${appName} - Channel search`,
        header: "Search channel messages",
        // allTeams: "Search all teams and channels",
        // from: "between",
        // to: "and",
        // messagesFound: (shown, total) => `${total} ${total === 1 ? "message" : "messages"} found ${total === shown ? "" : `- ${shown} displayed.`}`,
        // search: "search",
        // searching: "searching",
        // cancel: "cancel",
        // searchTimeAll: "All time",
        // searchTimePastWeek: "past week",
        // searchTimePastMonth: "past month",
        // searchTimePastYear: "past year",
        // searchTimeCustom: "custom",
        searchUsersLabel: "Display only messages from these users",
        searchUsersPlaceholder: "(All Users)",
    },

    chatSearch: {
        pageTitle: `${appName} - Chat search`,
        header: "Search personal chat messages",
        allChats: "Search all chats",
        // from: "between",
        // to: "and",
        // messagesFound: (shown, total) => `${total} ${total === 1 ? "message" : "messages"} found ${total === shown ? "" : `- ${shown} displayed.`}`,
        // search: "search",
        // searching: "searching",
        // cancel: "cancel",
        // searchTimeAll: "All time",
        // searchTimePastWeek: "past week",
        // searchTimePastMonth: "past month",
        // searchTimePastYear: "past year",
        // searchTimeCustom: "custom",
        searchUsersLabel: "Display only messages from these users",
        searchUsersPlaceholder: "(All Users)",
    },

    table: {
        subject: "subject",
        author: "author",
        dateTime: "created",
        body: "body",
    },

    schedule: {
        pageTitle: `${appName} - Events`,
        filterByStart: "Filter by start",
        filterByOrganizer: "Filter by organizer",
    },

    eventTable: {
        subject: "subject",
        organizer: "organizer",
        start: "start",
        end: "end",
        attendees: "attendees",
        body: "body",
        allday: "(all day)",
        notitle: "(no title)",
        noattendee: "(none)",
        },
}