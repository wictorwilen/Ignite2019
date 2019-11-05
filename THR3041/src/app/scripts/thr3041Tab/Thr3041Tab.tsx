import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@stardust-ui/react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import Customer from "../../defs/customer";
import * as msRest from "@azure/ms-rest-js";
import CustomerCard from "./CustomerCard";

/**
 * State for the thr3041TabTab React component
 */
export interface IThr3041TabState extends ITeamsBaseComponentState {
    entityId?: string;
    customers?: Customer[];
}

/**
 * Properties for the thr3041TabTab React component
 */
export interface IThr3041TabProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the THR3041 Tab content page
 */
export class Thr3041Tab extends TeamsBaseComponent<IThr3041TabProps, IThr3041TabState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                this.setState({
                    entityId: context.entityId
                });
                this.updateTheme(context.theme);

                const client = new msRest.ServiceClient(undefined, undefined);
                const request: msRest.RequestPrepareOptions = {
                    url: `https://${process.env.HOSTNAME}/api/customers?country=${this.state.entityId}`,
                    method: "GET"
                };
                client.sendRequest(request).then(result => {
                    this.setState({
                        customers: result.parsedBody
                    });
                });
            });
        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams"
            });
        }
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        return (
            <Provider theme={this.state.theme}>
                <Flex fill={true} column styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Flex.Item>
                        <Header content={`Customers from ${this.state.entityId}`} />
                    </Flex.Item>
                    <Flex.Item>
                        <div>
                            <div>
                                <Text content={this.state.entityId} />
                            </div>
                            <div>
                                {this.state.customers && this.state.customers.map(c => {
                                    return <CustomerCard {...c} avatar={{ image: c.avatar }} />;
                                })
                                }
                            </div>
                        </div>
                    </Flex.Item>
                    <Flex.Item styles={{
                        padding: ".8rem 0 .8rem .5rem"
                    }}>
                        <Text size="smaller" content="(C) Copyright Wictor Wilen" />
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }
}
