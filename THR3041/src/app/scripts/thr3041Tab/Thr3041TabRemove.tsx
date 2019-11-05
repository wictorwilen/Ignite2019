import * as React from "react";
import { Provider, Flex, Text, Header } from "@stardust-ui/react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";


export interface IThr3041TabRemoveState extends ITeamsBaseComponentState {
    value: string;
}
export interface IThr3041TabRemoveProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of THR3041 Tab remove page
 */
export class Thr3041TabRemove  extends TeamsBaseComponent<IThr3041TabRemoveProps, IThr3041TabRemoveState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
        } else {
        }
    }

    public render() {
        return (
            <Provider theme={this.state.theme}>
                <Flex fill={true}>
                    <Flex.Item>
                        <div>
                            <Header content="You're about to remove your tab..." />
                            <Text content="You can just add stuff here if you want to clean up when removing the tab. For instance, if you have stored data in an external repository, you can delete or archive it here. If you don't need this remove page you can remove it." />
                        </div>
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }
}
