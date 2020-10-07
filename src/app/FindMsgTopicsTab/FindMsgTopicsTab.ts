import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/FindMsgTopicsTab/index.html")
@PreventIframe("/FindMsgTopicsTab/config.html")
@PreventIframe("/FindMsgTopicsTab/remove.html")
export class FindMsgTopicsTab {
}
