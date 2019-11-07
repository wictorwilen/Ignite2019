import * as express from "express";
import * as crypto from "crypto";
import * as AuthenticationContext from "adal-node";
import { JsonDB } from "node-json-db";
import * as cookieParser from "cookie-parser";
import cookieSession = require("cookie-session");


// sample: https://github.com/AzureAD/azure-activedirectory-library-for-nodejs/blob/master/sample/website-sample.js
export const authRouter = (options: any): express.Router => {
    const router = express.Router();

    router.use(cookieParser("The secret"));
    router.use(cookieSession({ secret: "SOMERANDOMSTUFFINHERE" }));

    const templateAuthzUrl = "https://login.windows.net/common/oauth2/authorize?response_type=code&client_id=<client_id>&redirect_uri=<redirect_uri>&state=<state>&resource=<resource>";
    const redirectUri = `https://${process.env.hostname}/api/auth/getAToken`;
    const resource = `https://graph.microsoft.com`;
    const createAuthorizationUrl = (state: string) => {
        let authorizationUrl = templateAuthzUrl.replace("<client_id>", process.env.MICROSOFT_APP_ID as string);
        authorizationUrl = authorizationUrl.replace("<redirect_uri>", redirectUri);
        authorizationUrl = authorizationUrl.replace("<state>", state);
        authorizationUrl = authorizationUrl.replace("<resource>", resource);
        return authorizationUrl;
    };

    const tokens = new JsonDB("tokens", true, false);

    router.get("/auth", (req, res) => {
        crypto.randomBytes(48, (ex, buf) => {
            const state = buf.toString("base64").replace(/\//g, "_").replace(/\+/g, "-");
            res.cookie("authstate", state);
            res.cookie("notifyUrl", req.query.notifyUrl);
            const authorizationUrl = createAuthorizationUrl(state);
            res.redirect(authorizationUrl);
        });
    });

    router.get("/getAToken", (req, res) => {
        if (req.cookies.authstate !== req.query.state) {
            res.status(500).send("error: state does not match");
            return;
        }
        if (!req.cookies.notifyUrl) {
            res.status(500).send("Missing return notification url");
            return;
        }
        const authenticationContext = new AuthenticationContext.AuthenticationContext(`https://login.windows.net/common`);
        authenticationContext.acquireTokenWithAuthorizationCode(
            req.query.code,
            redirectUri,
            resource,
            process.env.MICROSOFT_APP_ID as string,
            process.env.MICROSOFT_APP_PASSWORD as string,
            (err, response) => {
                if (err) {
                    res.redirect(`https://${process.env.HOSTNAME}/${req.cookies.notifyUrl}?Failed=${err.message}`);
                } else {
                    // persist the token
                    tokens.push(`/tokens/${(response as AuthenticationContext.TokenResponse).oid}`, {
                        userId: (response as AuthenticationContext.TokenResponse).userId,
                        accessToken: (response as AuthenticationContext.TokenResponse).accessToken,
                        refreshToken: (response as AuthenticationContext.TokenResponse).refreshToken,
                    });
                    // TODO: fix the path
                    res.redirect(`https://${process.env.HOSTNAME}/${req.cookies.notifyUrl}?Success`);
                }


            });
    });
    return router;
};
