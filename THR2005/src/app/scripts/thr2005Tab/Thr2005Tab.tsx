import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@stardust-ui/react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State for the thr2005TabTab React component
 */
export interface IThr2005TabState extends ITeamsBaseComponentState {
    entityId?: string;
    items: any[];
}

/**
 * Properties for the thr2005TabTab React component
 */
export interface IThr2005TabProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the THR2005 Tab content page
 */
export class Thr2005Tab extends TeamsBaseComponent<IThr2005TabProps, IThr2005TabState> {

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
                fetch(`https://ignitedemoapi.azurewebsites.net/api/sessions?product=${this.state.entityId}`, {})
                    .then(result => {
                        result.json().then(json => {
                            this.setState({
                                items: json.data
                            });
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
                        <Header content="This is your tab" />
                    </Flex.Item>
                    <Flex.Item>
                        <div>
                            <div>
                                <Text content={this.state.entityId} />
                            </div>
                            <div>
                                {this.state.items && this.state.items.map((session: any) => {
                                    return (
                                        <Flex gap="gap.medium" padding="padding.medium">
                                            <Flex.Item size="size.medium">
                                                <div
                                                    style={{
                                                        position: "relative",
                                                    }}
                                                >
                                                    <Text as="h1" content={session.sessionCode} />
                                                </div>
                                            </Flex.Item>
                                            <Flex.Item grow>
                                                <Flex column gap="gap.small" vAlign="stretch">
                                                    <Flex space="between">
                                                        <Header as="h3" content={session.title} />
                                                        <Text as="em" content={session.startDateTime} />
                                                    </Flex>
                                                    <Text content={session.description} />
                                                    <Flex.Item push>
                                                        <Text as="em" content={session.speakerNames.join(
                                                        )} />
                                                    </Flex.Item>
                                                </Flex>
                                            </Flex.Item>
                                        </Flex>
                                    );
                                })}                            </div>
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
