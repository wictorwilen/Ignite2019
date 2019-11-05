import * as React from "react";
import { Provider, Flex, Header, Input, Dropdown } from "@stardust-ui/react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import * as msRest from "@azure/ms-rest-js";

export interface IThr3041TabConfigState extends ITeamsBaseComponentState {
    country: string;
    countries: string[];
}

export interface IThr3041TabConfigProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of THR3041 Tab configuration page
 */
export class Thr3041TabConfig extends TeamsBaseComponent<IThr3041TabConfigProps, IThr3041TabConfigState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();

            microsoftTeams.getContext((context: microsoftTeams.Context) => {
                this.setState({
                    country: context.entityId
                }, () => {
                    this.setValidityState(this.state.country !== undefined && this.state.country.length > 0);
                });
                this.updateTheme(context.theme);

            });

            microsoftTeams.settings.registerOnSaveHandler((saveEvent: microsoftTeams.settings.SaveEvent) => {
                // Calculate host dynamically to enable local debugging
                const host = "https://" + window.location.host;
                microsoftTeams.settings.setSettings({
                    contentUrl: host + "/thr3041Tab/?data=",
                    suggestedDisplayName: "THR3041 Tab",
                    removeUrl: host + "/thr3041Tab/remove.html",
                    entityId: this.state.country
                });
                saveEvent.notifySuccess();
            });

            const client = new msRest.ServiceClient(undefined, undefined);
            const request: msRest.RequestPrepareOptions = {
                url: `https://${process.env.HOSTNAME}/api/countries`,
                method: "GET"
            };
            client.sendRequest(request).then(result => {
                this.setState({
                    countries: result.parsedBody
                });
            });
        } else {
        }
    }

    public render() {
        return (
            <Provider theme={this.state.theme}>
                <Flex fill={true}>
                    <Flex.Item>
                        <div>
                            <Header content="Choose a country" />
                            <Dropdown
                                items={this.state.countries}
                                value={this.state.country}
                                onSelectedChange={(event, data) => {
                                    if (data) {
                                        this.setState({ country: data.value as string }, () => {
                                            this.setValidityState(this.state.country !== undefined && this.state.country.length > 0);
                                        });
                                    }
                                }} />
                        </div>
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }
}
