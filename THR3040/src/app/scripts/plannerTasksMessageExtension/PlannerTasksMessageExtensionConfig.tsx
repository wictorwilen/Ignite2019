import * as React from "react";
import { Provider, Flex, Header, Checkbox, Button } from "@stardust-ui/react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State for the PlannerTasksMessageExtensionConfig React component
 */
export interface IPlannerTasksMessageExtensionConfigState extends ITeamsBaseComponentState {
}

/**
 * Properties for the PlannerTasksMessageExtensionConfig React component
 */
export interface IPlannerTasksMessageExtensionConfigProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the Planner Tasks configuration page
 */
export class PlannerTasksMessageExtensionConfig extends TeamsBaseComponent<IPlannerTasksMessageExtensionConfigProps, IPlannerTasksMessageExtensionConfigState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        microsoftTeams.initialize();
        microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);

        // Just use this page for closing the window
        microsoftTeams.getContext((context: microsoftTeams.Context) => {
            alert(context.groupId);
            if (this.getQueryVariable("Success")) {
                microsoftTeams.authentication.notifySuccess(context.groupId);
            } else {
                const err = this.getQueryVariable("Failed");
                if (err) {
                    microsoftTeams.authentication.notifyFailure(err);
                } else {
                    microsoftTeams.authentication.notifyFailure("Unknown error");
                }
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
                        <div>Nothing here...</div>
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }
}
