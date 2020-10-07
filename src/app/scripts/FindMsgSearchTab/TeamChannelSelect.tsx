import { MyTeam } from "./FindMsgSearchTab";
import React from "react";
import { Checkbox, Divider, Flex, Segment, Text } from "../ui";


export interface ITeamSelectProps {
    teams: MyTeam[];
    checkState: Map<string, boolean>;
    all: boolean;
    allText: string;
    changed: (id?: string) => void;
}


interface ITeamChannelSelectProps {
    team: MyTeam;
    checkState: Map<string, boolean>;
    changed: (id: string) => void;
}


const TeamChannelSelect: React.SFC<ITeamChannelSelectProps> = ({ team: { id, displayName, channels }, changed, checkState }: ITeamChannelSelectProps) => {
    const allChannels = channels.every(c => checkState.get(c.id));

    return (
        <Segment>
            <Flex column hAlign="start">
                <Checkbox title={displayName} label={<Text content={displayName} weight="bold" />} checked={allChannels} onChange={() => changed(id)} key={id} />
                <Divider size={1} key="div" />
                {channels.map(c => <Checkbox title={c.displayName} label={c.displayName} checked={checkState.get(c.id)} onChange={() => changed(c.id)} key={c.id} />)}
            </Flex>
        </Segment>
    );
}


/**
 * Widget to select which channels/teams to search
 * @param props
 */
export const TeamSelect: React.SFC<ITeamSelectProps> = ({ all, teams, checkState, changed, allText }: ITeamSelectProps) => (
    <Flex column hAlign="start">
        <Checkbox label={allText} checked={all} onChange={() => changed()} />
        {!all && <Divider />}
        {!all && <Flex>{teams.map(t => <TeamChannelSelect key={t.id} team={t} checkState={checkState} changed={changed} />)}
        </Flex>}
    </Flex>
);
