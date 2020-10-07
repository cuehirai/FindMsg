/* eslint-disable react/prop-types */
import { IFindMsgChatMessage, FindMsgChatMessage } from "../db";
import React from "react";
import { Card, CardHeader, CardFooter, CardBody, Flex, FocusZone, FocusZoneDirection, Input, List, Link, Segment, Text, ComponentEventHandler, InputProps, mergeStyles } from "../ui";
import { collapse, empty, highlightEqual, highlightNode, noHighlight } from '../highlight';
import { isValid } from "../dateUtils";
import { collapseWhitespace } from '../purify';
import * as msTeams from '@microsoft/teams-js';


export interface SearchResultViewProps {
    messages: IFindMsgChatMessage[];
    searchChat: string;
    countFormat: (shown: number, total: number) => string;
    m2dt: (date: Date) => string;
    filter: string;
    showCollapsed: string;
    showExpanded: string;
    unknownUserDisplayName: string;
}

interface SearchResultProps {
    message: IFindMsgChatMessage;
    m2dt: (date: Date) => string;
    highlight: [string, string];
    showCollapsed: string;
    showExpanded: string;
    unknownUserDisplayName: string;
}

interface SearchResultState {
    message: IFindMsgChatMessage | null;
    highlight: Readonly<[string, string]>;
    content: Node | null;
    collapsible: boolean;
    collapsed: boolean;
    collapsedContent: Node | null;
}


const margin = mergeStyles({
    "margin-bottom": "16px"
});


const div = () => document.createElement("div");


/**
 * One single search result, with highlighing and collapsing
 */
class SearchResult extends React.Component<SearchResultProps, SearchResultState> {

    readonly ref = React.createRef<HTMLDivElement>();

    constructor(props: SearchResultProps) {
        super(props);
        this.state = {
            message: null,
            highlight: noHighlight,
            collapsed: true, // WARNING: default state must be collapsed, because of the size measurement of the parent list
            collapsible: false,
            content: null,
            collapsedContent: null,
        };
    }

    static getDerivedStateFromProps(newProps: SearchResultProps, oldState: SearchResultState): SearchResultState {
        const { message, highlight, message: { body, type } } = newProps;
        const { message: oldMessage, highlight: oldHighlight } = oldState;

        if (oldMessage === message && highlightEqual(oldHighlight, highlight)) return oldState;

        /*
           Implementation note:
           Want to preserve text formatting in the expanded state,
           but throw it away in the collapsed state to save screen space.
           Could eiher mark->clone->postprocess or mark twice.
           Both are definitely expensive operations expensive.
           Use the mark twice aproach, because it is less work to implement.
        */

        const content = div();
        if (type === "html") {
            content.innerHTML = body;
        } else {
            content.innerText = body; // will preserve line breaks as <br> and stuff when setting
        }

        const collapsedContent = div();
        collapsedContent.textContent = collapseWhitespace(content.innerText /* will throw away formatting when getting */);

        highlightNode(content, highlight);
        highlightNode(collapsedContent, highlight);

        const collapsible = collapse(collapsedContent) > 120; // only collapse if there is a significant amount

        return {
            message,
            highlight,
            collapsible,
            collapsed: true,
            content,
            collapsedContent: collapsible ? collapsedContent : content, // to free up the memory because content and collapsibleContent are identical anyway
        };
    }


    componentDidMount() { this.inject(); }
    componentDidUpdate() { this.inject(); }
    toggleCollapsed = () => this.setState({
        collapsed: !this.state.collapsed
    });


    inject() {
        const container = this.ref.current;

        if (!container) return;
        empty(container);

        const { collapsible, collapsed, content, collapsedContent } = this.state;
        const newContent = collapsible ? (collapsed ? collapsedContent : content) : content;

        if (container.firstChild === newContent || !newContent) return;

        container.appendChild(newContent);
    }


    render() {
        const { m2dt, showCollapsed, showExpanded, unknownUserDisplayName, message: { authorName, created, modified, id, chatId } } = this.props;
        const { collapsed, collapsible } = this.state;
        const edited = isValid(modified);

        const footer = collapsible ? (<CardFooter>
            <Link onClick={this.toggleCollapsed}>
                <Text content={collapsed ? showExpanded : showCollapsed} />
            </Link>
        </CardFooter>) : null;

        return (
            <Card fluid>
                <CardHeader>
                    <Flex gap="gap.medium" vAlign="center">
                        <Text content={authorName || unknownUserDisplayName} size="small" weight="bold" />
                        <Link onClick={() => msTeams.executeDeepLink(`https://teams.microsoft.com/l/chat/${chatId}/${id}`)}>
                            <Text timestamp content={m2dt(edited ? modified : created)} size="small" />
                        </Link>
                        {edited && <Text content="Edited" size="small" />}
                    </Flex>
                </CardHeader>
                <CardBody>
                    <div ref={this.ref} />
                </CardBody>
                {footer}
            </Card>
        );
    }
}


/**
 * List of search results
 * @param props
 */
export function SearchResultView(props: SearchResultViewProps): JSX.Element | null {
    const { filter, messages, countFormat, m2dt, searchChat, showCollapsed, showExpanded, unknownUserDisplayName } = props;

    if (!messages || messages.length === 0) return null;

    const [items, setItems] = React.useState(messages as IFindMsgChatMessage[]);
    const [filterKey, setFilterKey] = React.useState("");
    const terms: [string, string] = [searchChat.trim(), filterKey.trim()];
    const onFilterChanged: ComponentEventHandler<InputProps & { value: string; }> = (_: unknown, data): void => setFilterKey(data?.value ?? "");

    /* Important caveat: renderCell is NOT a react function component, but a plain function that returns ReactNode. That means, hooks can not be used. */
    const renderCell = (item?: IFindMsgChatMessage | undefined, index?: number): React.ReactNode => {
        //info("Render list item " + index);
        return item && <SearchResult key={index} message={item} m2dt={m2dt} highlight={terms} showCollapsed={showCollapsed} showExpanded={showExpanded} unknownUserDisplayName={unknownUserDisplayName} />;
    };

    React.useEffect(() => setItems(filterKey ? props.messages.filter(FindMsgChatMessage.createFilter(filterKey)) : props.messages), [props.messages, filterKey]);

    return messages.length < 1 ? null : (
        <FocusZone direction={FocusZoneDirection.vertical}>
            <Segment className={margin}>
                <Flex space="between" vAlign="center">
                    <Text key="NumResults" content={countFormat(items.length, messages.length)} />
                    <Input
                        type="text"
                        label={filter}
                        labelPosition="inline"
                        value={filterKey}
                        onChange={onFilterChanged}
                    />
                </Flex>
            </Segment>
            <List
                items={items}
                onRenderCell={renderCell}
                ignoreScrollingState
            />
        </FocusZone>
    );
}
