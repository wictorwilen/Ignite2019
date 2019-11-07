import * as React from "react";
import { Provider, Flex, Header, Input, Button, Text } from "@stardust-ui/react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State for the CreateTaskMessageExtensionAction React component
 */
export interface ICreateTaskMessageExtensionActionState extends ITeamsBaseComponentState {
    title: string;
}

/**
 * Properties for the CreateTaskMessageExtensionAction React component
 */
export interface ICreateTaskMessageExtensionActionProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the Create task Task Module page
 */
export class CreateTaskMessageExtensionAction extends TeamsBaseComponent<ICreateTaskMessageExtensionActionProps, ICreateTaskMessageExtensionActionState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        microsoftTeams.initialize();
        microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
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
                        <div>
                            <Header content="Create a new task" />
                            <Text content="Task name" />
                            <Input
                                fluid
                                clearable
                                value={this.state.title}
                                onChange={(e, data) => {
                                    if (data) {
                                        this.setState({
                                            title: data.value
                                        });
                                    }
                                }}
                                required />
                            <Button onClick={() => microsoftTeams.tasks.submitTask({
                                    title: this.state.title
                                })} primary>Create</Button>
                        </div>
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }
}
