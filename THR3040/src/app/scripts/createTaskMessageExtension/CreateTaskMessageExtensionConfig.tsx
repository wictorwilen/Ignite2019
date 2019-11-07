import * as React from "react";
import { Provider, Flex, Header, Checkbox, Button, radioGroupBehavior } from "@stardust-ui/react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State for the CreateTaskMessageExtensionConfig React component
 */
export interface ICreateTaskMessageExtensionConfigState extends ITeamsBaseComponentState {
    addLink: boolean;
    groupId?: string;
}

/**
 * Properties for the CreateTaskMessageExtensionConfig React component
 */
export interface ICreateTaskMessageExtensionConfigProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the Create task configuration page
 */
export class CreateTaskMessageExtensionConfig extends TeamsBaseComponent<ICreateTaskMessageExtensionConfigProps, ICreateTaskMessageExtensionConfigState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        microsoftTeams.initialize();
        microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);

        // If we're here after the auth, then
        microsoftTeams.getContext((context: microsoftTeams.Context) => {
            this.setState({
                groupId: context.groupId
            });

            const err = this.getQueryVariable("Failed");
            if (err) {
                microsoftTeams.authentication.notifyFailure(err);
            }
        });
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        return (
            <Provider theme={this.state.theme}>
                <Flex fill={true}>
                    <Flex.Item>
                        <div>
                            <Header content="Create task configuration" />
                            <Checkbox
                                label="Do you want to automatiaclly add a link to the conversation in the task?"
                                toggle
                                checked={this.state.addLink}
                                onChange={() => {
                                    this.setState({
                                        addLink: !this.state.addLink
                                    });
                                }} />
                            <Button onClick={() =>
                                microsoftTeams.authentication.notifySuccess(JSON.stringify({
                                    addLink: this.state.addLink === true,
                                    groupId: this.state.groupId
                                }))} primary>Save</Button>
                        </div>
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }
}
