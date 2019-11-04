import * as React from "react";
import { Provider, Flex, Header, Dropdown } from "@stardust-ui/react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

export interface IThr2005TabConfigState extends ITeamsBaseComponentState {
    value: string;
    items: string[];
    loading: boolean;
}

export interface IThr2005TabConfigProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of THR2005 Tab configuration page
 */
export class Thr2005TabConfig extends TeamsBaseComponent<IThr2005TabConfigProps, IThr2005TabConfigState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();

            microsoftTeams.getContext((context: microsoftTeams.Context) => {
                this.setState({
                    value: context.entityId,
                    loading: true
                });
                this.updateTheme(context.theme);
                this.setValidityState(false);
                fetch("https://ignitedemoapi.azurewebsites.net/api/products", {})
                    .then(result => {
                        result.json().then(json => {
                            this.setState({
                                items: json,
                                loading: false
                            });
                        });
                    });
            });

            microsoftTeams.settings.registerOnSaveHandler((saveEvent: microsoftTeams.settings.SaveEvent) => {
                // Calculate host dynamically to enable local debugging
                const host = "https://" + window.location.host;
                microsoftTeams.settings.setSettings({
                    contentUrl: host + "/thr2005Tab/?data=",
                    suggestedDisplayName: "THR2005 Tab",
                    removeUrl: host + "/thr2005Tab/remove.html",
                    entityId: this.state.value
                });
                saveEvent.notifySuccess();
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
                            <Header content="Select product" />
                            <Dropdown
                                items={this.state.items}
                                loading={this.state.loading}
                                loadingMessage="Retrieving data..."
                                value={this.state.value}
                                open={!this.state.loading}
                                onSelectedChange={(event, data) => {
                                    if (data) {
                                        this.setState({ value: data.value as string }, () => {
                                            this.setValidityState(this.state.value !== undefined && this.state.value.length > 0);
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
