import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/thr2005Tab/index.html")
@PreventIframe("/thr2005Tab/config.html")
@PreventIframe("/thr2005Tab/remove.html")
export class Thr2005Tab {
}
