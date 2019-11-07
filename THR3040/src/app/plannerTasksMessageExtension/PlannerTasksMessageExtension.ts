import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory } from "botbuilder";
import { MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder-teams";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import { JsonDB } from "node-json-db";
import * as AuthenticationContext from "adal-node";
import * as msRest from "@azure/ms-rest-js";

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/plannerTasksMessageExtension/config.html")
export default class PlannerTasksMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public static verifySignedIn(context: TurnContext, success: (accessToken: string) => Promise<MessagingExtensionResult>): Promise<MessagingExtensionResult> {
        const tokens = new JsonDB("tokens", true, false);

        let token: any;
        try {
            token = tokens.getData(`/tokens/${context.activity.from.aadObjectId}`);
        } catch (error) {
            token = undefined;
        }
        if (!token) {
            return Promise.resolve<MessagingExtensionResult>({
                type: "auth", // use "config" or "auth" here
                suggestedActions: {
                    actions: [
                        {
                            type: "openUrl",
                            value: `https://${process.env.HOSTNAME}/api/auth/auth?notifyUrl=plannerTasksMessageExtension/config.html`,
                            title: "Setup"
                        }
                    ]
                }
            });
        } else {
            return new Promise<MessagingExtensionResult>(async (resolve, reject) => {
                const authenticationContext = new AuthenticationContext.AuthenticationContext(`https://login.windows.net/common`);
                authenticationContext.acquireTokenWithRefreshToken(
                    token.refreshToken,
                    process.env.MICROSOFT_APP_ID as string,
                    process.env.MICROSOFT_APP_PASSWORD as string,
                    `https://graph.microsoft.com`,
                    async (refreshErr, refreshResponse) => {
                        if (refreshErr) {
                            reject(refreshErr.message);
                        }
                        resolve(await success((refreshResponse as any).accessToken));
                    });
            });
        }
    }


    public async onQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {
        log("onQuery");
        return PlannerTasksMessageExtension.verifySignedIn(context, async (accessToken) => {

            // if state is found in the query, then we need to persist the group so we can match channeldata.team and group
            let groupId;
            const groups = new JsonDB("groups", true, false);
            if (query.state) {
                groupId = query.state;
                groups.push(`/groups/${context.activity.channelData.team.id}`, groupId);
            } else {
                try {
                    groupId = groups.getData(`/groups/${context.activity.channelData.team.id}`).groupId;
                } catch (error) {
                    groupId = undefined;
                }
            }

            const creds = new msRest.TokenCredentials(accessToken);
            const client = new msRest.ServiceClient(creds, undefined);
            const request: msRest.RequestPrepareOptions = {
                url: `https://graph.microsoft.com/v1.0/groups/${groupId}/planner/plans`,
                method: "GET"
            };
            const response = await client.sendRequest(request);

            const request2: msRest.RequestPrepareOptions = {
                url: `https://graph.microsoft.com/v1.0/planner/plans/${response.parsedBody.value[0].id}/tasks`,
                method: "GET"
            };
            const response2 = await client.sendRequest(request2);

            const request3: msRest.RequestPrepareOptions = {
                url: `https://graph.microsoft.com/beta/planner/plans/${response.parsedBody.value[0].id}/details`,
                method: "GET"
            };
            const response3 = await client.sendRequest(request3);

            let tasks;

            if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
                // initial run
                // show them all
                tasks = response2.parsedBody.value;
            } else {
                // the rest
                tasks = response2.parsedBody.value.filter(task => (task.title as string).indexOf(query!.parameters![0].value) !== -1);
            }
            const attachments = tasks.map(task => {
                log(task);
                const encodedContext = encodeURI(JSON.stringify({ subEntityId: task.id, channelId: context.activity.channelData.channel.id }));
                const url = `https://teams.microsoft.com/l/entity/com.microsoft.teamspace.tab.planner/${response3.parsedBody.id}?context=${encodedContext}`;
                const card = CardFactory.adaptiveCard(
                    {
                        type: "AdaptiveCard",
                        body: [
                            {
                                type: "TextBlock",
                                size: "Large",
                                text: task.title
                            },
                            {
                                type: "ColumnSet",
                                columns: [
                                    {
                                        type: "Column",
                                        width: "auto",
                                        items: [
                                            {
                                                type: "Image",
                                                url: `https://${process.env.HOSTNAME}/assets/icon.png`
                                            }
                                        ]
                                    },
                                    {
                                        type: "Column",
                                        width: "stretch",
                                        items: [
                                            {
                                                type: "FactSet",
                                                facts: [
                                                    {
                                                        title: "Categories",
                                                        value: Object.getOwnPropertyNames(task.appliedCategories).map(cat => response3.parsedBody.categoryDescriptions[cat]).join(",")
                                                    },
                                                    {
                                                        title: "Created",
                                                        value: `{{DATE(${task.createdDateTime.substr(0, task.createdDateTime.indexOf("."))}Z, SHORT)}}`
                                                    }
                                                ]
                                            }
                                        ]
                                    }
                                ]
                            }
                        ],
                        actions: [
                            {
                                type: "Action.Submit",
                                title: "Mark as complete",
                                data: {
                                    action: "markAsComplete",
                                    taskId: task.id,
                                    etag: task["@odata.etag"]
                                }
                            },
                            {
                                type: "Action.OpenUrl",
                                title: "Go to plan and task",
                                url
                            }
                        ],
                        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                        version: "1.0"
                    });
                const preview = {
                    contentType: "application/vnd.microsoft.card.thumbnail",
                    content: {
                        title: task.title,
                        text: `Categories: ${Object.getOwnPropertyNames(task.appliedCategories).map(cat => response3.parsedBody.categoryDescriptions[cat]).join(",")} `,
                        images: [
                            {
                                url: `https://${process.env.HOSTNAME}/assets/icon.png`
                            }
                        ]
                    }
                };
                return { ...card, preview };
            });

            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments
            } as MessagingExtensionResult);
        });
    }


    public async onCardButtonClicked(context: TurnContext, value: any): Promise<void> {
        // Handle the Action.Submit action on the adaptive card
        if (value.action === "markAsComplete") {
            const tokens = new JsonDB("tokens", true, false);

            let token: any;
            try {
                token = tokens.getData(`/tokens/${context.activity.from.aadObjectId}`);
            } catch (error) {
                token = undefined;
            }

            const body = {
                percentComplete: 100
            };

            const request: msRest.RequestPrepareOptions = {
                url: `https://graph.microsoft.com/v1.0/planner/tasks/${value.taskId}`,
                method: "PATCH",
                body,
                headers: {
                    "If-Match": value.etag.substr(2) // remove the weak etag (W/)
                }
            };
            const creds = new msRest.TokenCredentials(token.accessToken);
            const client = new msRest.ServiceClient(creds, undefined);

            const response = await client.sendRequest(request);
            log(response);
            log(`Marked task as complete`);
        }
        return Promise.resolve();
    }






    // this is used when canUpdateConfiguration is set to true
    public async onQuerySettingsUrl(context: TurnContext): Promise<{ title: string, value: string }> {
        return Promise.resolve({
            title: "Planner Tasks Configuration",
            value: `https://${process.env.HOSTNAME}/plannerTasksMessageExtension/config.html`
        });
    }

    public async onSettings(context: TurnContext): Promise<void> {
        // take care of the setting returned from the dialog, with the value stored in state
        const setting = context.activity.value.state;
        log(`New setting: ${setting}`);
        return Promise.resolve();
    }

}
