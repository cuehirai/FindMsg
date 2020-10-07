// Load application insights
export { AI } from './appInsights';

// Automatically added for the FindMsgTopicsTab tab
export * from "./FindMsgTopicsTab/FindMsgTopicsTab";
export * from "./FindMsgTopicsTab/FindMsgTopicsTabConfig";
export * from "./FindMsgTopicsTab/FindMsgTopicsTabRemove";
// Automatically added for the FindMsgSearchTab tab
export * from "./FindMsgSearchTab/FindMsgSearchTab";
// Automatically added for the FindMsgSearchChat tab
export * from "./FindMsgSearchChat/FindMsgSearchChat";

// only for (manual) debugging
export { db } from "./db/Database";
export { Sync } from "./db/Sync";

// for side-effect of registering the custom element
import './graphImage';

export { teamsTheme, teamsDarkTheme, teamsHighContrastTheme } from './ui';
