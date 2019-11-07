import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory } from "botbuilder";
import { MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder-teams";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import { ITaskModuleResult, IMessagingExtensionActionRequest } from "botbuilder-teams-messagingextensions";
import { JsonDB } from "node-json-db";
import * as AuthenticationContext from "adal-node";
import * as msRest from "@azure/ms-rest-js";

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/createTaskMessageExtension/config.html")
@PreventIframe("/createTaskMessageExtension/action.html")
export default class CreateTaskMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public static verifySignedIn(context: TurnContext, success: (accessToken: string) => Promise<MessagingExtensionResult | ITaskModuleResult>): Promise<MessagingExtensionResult | ITaskModuleResult> {
        const tokens = new JsonDB("tokens", true, false);
        log("checking auth");
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
                            value: `https://${process.env.HOSTNAME}/api/auth/auth?notifyUrl=createTaskMessageExtension/config.html`,
                            title: "Configuration"
                        }
                    ]
                }
            });
        } else {
            return new Promise<MessagingExtensionResult | ITaskModuleResult>(async (resolve, reject) => {
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


    public async onFetchTask(context: TurnContext, value: IMessagingExtensionActionRequest): Promise<MessagingExtensionResult | ITaskModuleResult> {
        return CreateTaskMessageExtension.verifySignedIn(context, async (accessToken) => {


            // if state is found in the query, then we need to persist the group so we can match channeldata.team and group
            let groupId: string | undefined;
            let addLink: boolean;
            const groups = new JsonDB("groups", true, false);
            log(value);
            if (value.state) {
                // save the state
                const data = JSON.parse(value.state);
                groupId = data.groupId;
                addLink = data.addLink;
                groups.push(`/groups/${context.activity.channelData.team.id}`, { groupId, addLink });
            } else {
                // get the state
                try {
                    const data = groups.getData(`/groups/${context.activity.channelData.team.id}`);
                    groupId = data.groupId;
                    addLink = data.addLink;
                } catch (error) {
                    groupId = undefined;
                }
            }


            if (!groupId) {
                return Promise.resolve<MessagingExtensionResult>({
                    type: "config",
                    suggestedActions: {
                        actions: [
                            {
                                type: "openUrl",
                                value: `https://${process.env.HOSTNAME}/createTaskMessageExtension/config.html`,
                                title: "Configuration"
                            }
                        ]
                    }
                });
            }


            return Promise.resolve<ITaskModuleResult>({
                type: "continue",
                value: {
                    title: "Input form",
                    url: `https://${process.env.HOSTNAME}/createTaskMessageExtension/action.html`
                }
            });
        });
    }


    // handle action response in here
    // See documentation for `MessagingExtensionResult` for details
    public async onSubmitAction(context: TurnContext, value: IMessagingExtensionActionRequest): Promise<MessagingExtensionResult> {

        const tokens = new JsonDB("tokens", true, false);
        const groups = new JsonDB("groups", true, false);

        let token: any;
        try {
            token = tokens.getData(`/tokens/${context.activity.from.aadObjectId}`);
        } catch (error) {
            token = undefined;
        }

        const data = groups.getData(`/groups/${context.activity.channelData.team.id}`);
        const groupId = data.groupId;
        const addLink = data.addLink;

        const creds = new msRest.TokenCredentials(token.accessToken);
        const client = new msRest.ServiceClient(creds, undefined);
        const request: msRest.RequestPrepareOptions = {
            url: `https://graph.microsoft.com/v1.0/groups/${groupId}/planner/plans`,
            method: "GET"
        };
        const response = await client.sendRequest(request);
        const planId = response.parsedBody.value[0].id;

        const request2: msRest.RequestPrepareOptions = {
            url: `https://graph.microsoft.com/v1.0/planner/plans/${planId}/buckets`,
            method: "GET"
        };
        const response2 = await client.sendRequest(request2);
        const bucketId = response2.parsedBody.value[0].id;


        const request3: msRest.RequestPrepareOptions = {
            url: `https://graph.microsoft.com/v1.0/planner/tasks`,
            method: "POST",
            body: {
                planId,
                bucketId,
                title: value.data.title
            }
        };
        const response3 = await client.sendRequest(request3);

        if (addLink) {
            const request5: msRest.RequestPrepareOptions = {
                url: `https://graph.microsoft.com/v1.0/planner/tasks/${response3.parsedBody.id}/details`,
                method: "GET"
            };
            const response5 = await client.sendRequest(request5);
            const body = {
                previewType: "reference",
                description: context.activity.value.messagePayload.body.content,
                references: {
                }
            };
            body.references[`https%3A//teams%2Emicrosoft%2Ecom/l/message/${encodeURIComponent(context.activity.channelData.channel.id).replace(".", "%2E")}/${context.activity.value.messagePayload.id}?tenantId=${context.activity.channelData.tenant.id}`] = {
                "@odata.type": "#microsoft.graph.plannerExternalReference",
                "alias": "Conversation",
                "type": "Other"
            };
            const request4: msRest.RequestPrepareOptions = {
                url: `https://graph.microsoft.com/v1.0/planner/tasks/${response3.parsedBody.id}/details`,
                method: "PATCH",
                body,
                headers: {
                    "If-Match": response5.parsedBody["@odata.etag"].substr(2) // remove the weak etag (W/)
                }
            };

            const response4 = await client.sendRequest(request4);
        }
        const encodedContext = encodeURI(JSON.stringify({ subEntityId: response3.parsedBody.id, channelId: context.activity.channelData.channel.id }));
        const url = `https://teams.microsoft.com/l/entity/com.microsoft.teamspace.tab.planner/${response3.parsedBody.planId}?context=${encodedContext}`;
        const card = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Large",
                        text: response3.parsedBody.title
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
                            }
                        ]
                    }
                ],
                actions: [
                    {
                        type: "Action.OpenUrl",
                        title: "Go to plan and task",
                        url
                    }
                ],
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                version: "1.0"
            });
        return Promise.resolve({
            type: "result",
            attachmentLayout: "list",
            attachments: [card]
        } as MessagingExtensionResult);
    }



}
