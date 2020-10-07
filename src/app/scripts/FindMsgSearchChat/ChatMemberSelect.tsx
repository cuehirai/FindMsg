import React from "react";
import { Checkbox, Divider, Flex } from "../ui";
import { IFindMsgChatEx } from "./FindMsgSearchChat";


export interface IChatSelectProps {
    chats: IFindMsgChatEx[];
    checkState: Map<string, boolean>;
    all: boolean;
    allText: string;
    changed: (id?: string) => void;
}


/**
 * Widget to select which channels/chats to search
 * @param props
 */
export const ChatSelect: React.SFC<IChatSelectProps> = ({ all, chats, checkState, changed, allText }: IChatSelectProps) => {
    const one: IFindMsgChatEx[] = [];
    const many: IFindMsgChatEx[] = [];

    chats.forEach(c => c.topic ? many.push(c) : one.push(c));

    return (
        <Flex column hAlign="start">
            <Checkbox label={allText} checked={all} onChange={() => changed()} />
            {!all && <Divider />}
            {!all && <Flex gap="gap.large">
                {one.length > 0 &&
                    <Flex column hAlign="start">
                        {one.map(({ id, singleCounterpart }) => <Checkbox label={singleCounterpart ?? "(unknown)"} dir="ltr" checked={checkState.get(id)} onChange={() => changed(id)} key={id} />)}
                    </Flex>
                }

                {one.length > 0 && many.length > 0 && <Divider vertical />}

                {one.length > 0 &&
                    <Flex column>
                        {many.map(({ id, topic }) => <Checkbox label={topic ?? "(unknown)"} dir="ltr" checked={checkState.get(id)} onChange={() => changed(id)} key={id} />)}
                    </Flex>
                }
            </Flex>
            }
        </Flex>
    );
};
