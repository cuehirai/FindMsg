import * as React from "react";

import TeamsBaseComponent, { ITeamsBaseComponentState } from "../msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import { Flex, Header, Provider, Text } from "../ui";


export interface IFindMsgTopicsTabRemoveState extends ITeamsBaseComponentState {
    value: string;
}


/**
 * Implementation of FindMsg topics remove page
 */
export class FindMsgTopicsTabRemove extends TeamsBaseComponent<never, IFindMsgTopicsTabRemoveState> {

    public async componentWillMount(): Promise<void> {
        this.updateTheme(this.getQueryVariable("theme"));

        if (await this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.appInitialization.notifySuccess();
        }
    }

    public render(): JSX.Element {
        return (
            <Provider theme={this.state.theme}>
                <Flex fill={true}>
                    <Flex.Item>
                        <div>
                            <Header content="You're about to remove your tab..." />
                            <Text content="You can just add stuff here if you want to clean up when removing the tab. For instance, if you have stored data in an external repository, you can delete or archive it here. If you don't need this remove page you can remove it." />
                        </div>
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }
}
