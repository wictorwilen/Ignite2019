import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/thr3041Tab/index.html")
@PreventIframe("/thr3041Tab/config.html")
@PreventIframe("/thr3041Tab/remove.html")
export class Thr3041Tab {
}
