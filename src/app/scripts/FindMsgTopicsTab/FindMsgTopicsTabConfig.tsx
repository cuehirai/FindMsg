import * as React from "react";

import TeamsBaseComponent, { ITeamsBaseComponentState } from "../msteams-react-base-component";
import * as msTeams from "@microsoft/teams-js";
import { assert } from '../utils';
import { Flex, Header, Input, Provider } from "../ui";
import { AI } from '../appInsights';
import * as strings from '../i18n/messages';
import { IMessageTranslation } from "../i18n/IMessageTranslation";


export interface IFindMsgTopicsTabConfigState extends ITeamsBaseComponentState {
    entityId?: string;
    loading: boolean;
    inTeams: boolean;
    isPrivateChannel: boolean;
    channelId: string;
    groupId: string;
    value: string;
    locale: string | null;
    t: IMessageTranslation;
}


export interface IFindMsgTopicsTabConfigTranslation {
    loading: string;
    errorNotInTeams: string;
    errorPrivateChannel: string;
    errorNoGroupId: string;
    errorNoChannelId: string;
    headerConfigure: string;
    labelTabName: string;
    placeholderTabName: string;
    defaultTabName: string;
}


export class FindMsgTopicsTabConfig extends TeamsBaseComponent<never, IFindMsgTopicsTabConfigState> {

    constructor(props: never) {
        super(props);

        const locale = this.getQueryVariable("l") || null;
        const t = strings.get(locale);

        this.state = {
            loading: true,
            inTeams: false,
            isPrivateChannel: false,
            channelId: "",
            groupId: "",
            value: t.topicsConfig.defaultTabName,
            entityId: this.getQueryVariable("eid"),
            locale,
            theme: this.getTheme(this.getQueryVariable("theme")),
            t,
        }
    }


    public async componentWillMount(): Promise<void> {
        this.updateTheme(this.getQueryVariable("theme"));

        if (await this.inTeams(10000)) {
            msTeams.initialize();
            msTeams.getContext(this.contextReceived);
        } else {
            this.setState({
                loading: false,
                inTeams: false,
            });
        }
    }


    contextReceived = (context: msTeams.Context): void => {
        const t = strings.get(context.locale);

        this.setState({
            channelId: context.channelId ?? "",
            groupId: context.groupId ?? "",
            loading: false,
            inTeams: true,
            isPrivateChannel: context.channelType === "Private",
            locale: context.locale,
            t,
            value: t.topicsConfig.defaultTabName,
        });

        this.updateTheme(context.theme);
        msTeams.appInitialization.notifySuccess();

        msTeams.settings.registerOnSaveHandler(this.saveHandler);

        // IMPORTANT: check if the channel is private
        // In a private channel, groupdId is hidden (there might be a way around this, by letting the user select team/channel to use...), so we can not retrieve messages!
        // Display an error message to tell the user
        // see: https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/access-teams-context
        msTeams.settings.setValidityState(context.channelType === "Regular" && !!context.channelId && !!context.groupId);
    };


    saveHandler = (saveEvent: msTeams.settings.SaveEvent): void => {
        // Calculate host dynamically to enable local debugging
        try {
            const host = "https://" + window.location.host;
            const { channelId, groupId, value } = assert(this.state, nameof<msTeams.Context>());

            const tabUrl = `${host}/FindMsgTopicsTab/?theme={theme}&l={locale}&eid={entityId}&sid={subEntityId}&cid=${channelId}&gid=${groupId}&hint={loginHint}&tid={tid}&uid={userObjectId}`;

            AI.trackEvent({ name: "AddTopicsTab" });
            AI.flushBuffer();

            msTeams.settings.setSettings({
                contentUrl: tabUrl,
                websiteUrl: tabUrl,
                suggestedDisplayName: value,
                entityId: assert(this.state.entityId)
            });
            saveEvent.notifySuccess();
        } catch (error) {
            AI.trackException({ exception: error });
            saveEvent.notifyFailure(error.message);
        }
    };


    // TODO: prettify
    public render(): JSX.Element {
        const { loading, inTeams, isPrivateChannel, theme, value, groupId, channelId, t: { topicsConfig: t } } = this.state;
        let content: JSX.Element | JSX.Element[];

        if (loading) {
            content = <Header content={t.loading} />;
        } else if (!inTeams) {
            content = <Header content={t.errorNotInTeams} />;
        } else if (isPrivateChannel) {
            content = <Header content={t.errorPrivateChannel} />;
        } else if (!groupId) {
            content = <Header content={t.errorNoGroupId} />;
        } else if (!channelId) {
            content = <Header content={t.errorNoChannelId} />;
        } else {
            content = [
                <Header content={t.headerConfigure} key="header" />,
                <Input
                    label={t.labelTabName}
                    key="input"
                    required fluid clearable
                    placeholder={t.placeholderTabName}
                    value={value}
                    onChange={(_, data) => this.setState({ value: data?.value ?? "" })}
                />
            ];
        }

        return (
            <Provider theme={theme}>
                <Flex column fill={true}>
                    {content}
                </Flex>
            </Provider>
        );
    }
}
