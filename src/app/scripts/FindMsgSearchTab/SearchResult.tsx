/* eslint-disable react/prop-types */
import { IFindMsgChannelMessage, FindMsgChannelMessage } from "../db";
import React from "react";

import { Card, CardHeader, CardFooter, CardBody, Flex, FocusZone, FocusZoneDirection, Input, List, Link, Segment, Text, ComponentEventHandler, InputProps, mergeStyles } from "../ui";

import * as msTeams from '@microsoft/teams-js';
import { collapse, empty, highlightEqual, highlightNode, noHighlight } from '../highlight';
import { isValid } from "../dateUtils";
import { info } from '../logger';
import { fixMessageLink } from "../utils";
import { collapseWhitespace } from '../purify';
import { HtmlTooltip, IChannelInfo } from "../ui-jsx";
import { Typography } from "@material-ui/core";


export interface SearchResultViewProps {
    messages: IFindMsgChannelMessage[];
    searchTerm: string;
    countFormat: (shown: number, total: number) => string;
    m2dt: (date: Date) => string;
    filter: string;
    showCollapsed: string;
    showExpanded: string;
    unknownUserDisplayName: string;
    channelMap: Map<string, IChannelInfo>;
    teamchannel: (teamname: string, channelname: string) => string;
}

interface SearchResultProps {
    message: IFindMsgChannelMessage;
    m2dt: (date: Date) => string;
    highlight: [string, string];
    showCollapsed: string;
    showExpanded: string;
    unknownUserDisplayName: string;
    channelMap: Map<string, IChannelInfo>;
    teamchannel: (teamname: string, channelname: string) => string;
}

interface SearchResultState {
    message: IFindMsgChannelMessage | null;
    highlight: Readonly<[string, string]>;
    subjectHtml: null | { __html: string };
    content: Node | null;
    collapsible: boolean;
    collapsed: boolean;
    collapsedContent: Node | null;
}

interface ITeamChannelTooltipArg {
    channelId: string;
    edited: boolean;
    modified: Date;
    created: Date;
    url: string;
    m2dt: (date: Date) => string;
    channelMap: Map<string, IChannelInfo>;
    teamchannel: (teamname: string, channelname: string) => string;
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
            subjectHtml: null,
            collapsed: true, // WARNING: default state must be collapsed, because of the size measurement of the parent list
            collapsible: false,
            content: null,
            collapsedContent: null,
        };
    }

    static getDerivedStateFromProps(newProps: SearchResultProps, oldState: SearchResultState): SearchResultState {
        const { message, highlight, message: { body, subject, type } } = newProps;
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

        let subjectHtml: null | { __html: string } = null;

        if (subject) {
            const s = div();
            s.innerText = subject;
            highlightNode(s, highlight);
            subjectHtml = { __html: s.innerHTML };
        }

        highlightNode(content, highlight);
        highlightNode(collapsedContent, highlight);

        const collapsible = collapse(collapsedContent) > 120; // only collapse if there is a significant amount

        return {
            message,
            highlight,
            subjectHtml,
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
        const { m2dt, showCollapsed, showExpanded, unknownUserDisplayName, channelMap, teamchannel, message: { authorName, subject, created, modified, url, channelId } } = this.props;
        const { collapsed, collapsible, subjectHtml } = this.state;
        const edited = isValid(modified);

        const footer = collapsible ? (<CardFooter>
            <Link onClick={this.toggleCollapsed}>
                <Text content={collapsed ? showExpanded : showCollapsed} />
            </Link>
        </CardFooter>) : null;

        return (
            <Card fluid>
                <CardHeader>
                    <Flex column gap="gap.small">
                        <Flex gap="gap.medium" vAlign="center">
                            <Text content={authorName || unknownUserDisplayName} size="small" weight="bold" />
                            {/* <Link onClick={() => msTeams.executeDeepLink(fixMessageLink(url), info)} disabled={!url}>
                                <Text timestamp content={m2dt(edited ? modified : created)} size="small" />
                            </Link> */}
                            {this.teamchannelTooltip({channelId, edited, modified, created, url, m2dt, channelMap, teamchannel})}
                            {edited && <Text content="Edited" size="small" />}
                        </Flex>
                        {subject && (subjectHtml ? <Text dangerouslySetInnerHTML={subjectHtml} size="large" weight="bold" /> : <Text content={subject} size="large" weight="bold" />)}
                    </Flex>
                </CardHeader>
                <CardBody>
                    <div ref={this.ref} />
                </CardBody>
                {footer}
            </Card>
        );
    }

    private teamchannelTooltip(params:ITeamChannelTooltipArg): JSX.Element {
        const channelInfo = params.channelMap.get(params.channelId);
        const name = channelInfo? params.teamchannel(channelInfo.teamDisplayName, channelInfo.channelDisplayName) : null;
        const link: JSX.Element = (
            <Link onClick={() => msTeams.executeDeepLink(fixMessageLink(params.url), info)} disabled={!params.url}>
                <Text timestamp content={params.m2dt(params.edited ? params.modified : params.created)} size="small" />
            </Link>
        );

        if (name) {
            return (
                <HtmlTooltip
                    title={
                        <React.Fragment>
                        <Typography color="inherit">{name}</Typography>
                        </React.Fragment>
                    }
                >
                    {link}
                </HtmlTooltip>
            );
        } else {
            return link;
        }
    }
}


/**
 * List of search results
 * @param props
 */
export function SearchResultView(props: SearchResultViewProps): JSX.Element | null {
    const { filter, messages, countFormat, m2dt, searchTerm, showCollapsed, showExpanded, unknownUserDisplayName, channelMap, teamchannel } = props;

    if (!messages || messages.length === 0) return null;

    const [items, setItems] = React.useState(messages as IFindMsgChannelMessage[]);
    const [filterKey, setFilterKey] = React.useState("");
    const terms: [string, string] = [searchTerm.trim(), filterKey.trim()];
    const onFilterChanged: ComponentEventHandler<InputProps & { value: string; }> = (_: unknown, data): void => setFilterKey(data?.value ?? "");

    /* Important caveat: renderCell is NOT a react function component, but a plain function that returns ReactNode. That means, hooks can not be used. */
    const renderCell = (item?: IFindMsgChannelMessage | undefined, index?: number): React.ReactNode => {
        //info("Render list item " + index);
        return item && <SearchResult key={index} message={item} m2dt={m2dt} highlight={terms} showCollapsed={showCollapsed} showExpanded={showExpanded} unknownUserDisplayName={unknownUserDisplayName} channelMap={channelMap} teamchannel={teamchannel} />;
    };

    React.useEffect(() => setItems(filterKey ? props.messages.filter(FindMsgChannelMessage.createFilter(filterKey)) : props.messages), [props.messages, filterKey]);

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
