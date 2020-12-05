import * as React from "react";
import { Provider,
    Flex,
    Text,
    Header,
    List,
    Alert,
    teamsTheme, teamsDarkTheme, teamsHighContrastTheme,
    ThemePrepared,
    WindowMaximizeIcon,
    ExclamationTriangleIcon,
    Label,
    Button,
    Input,
    ToDoListIcon } from "@fluentui/react-northstar";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
/**
 * State for the learnPersonalTabTab React component
 */
export interface ILearnPersonalTabState extends ITeamsBaseComponentState {
    entityId?: string;
}

/**
 * Properties for the learnPersonalTabTab React component
 */
export interface ILearnPersonalTabProps {

}

/**
 * Implementation of the LearnPersonalTab content page
 */
export class LearnPersonalTab extends TeamsBaseComponent<ILearnPersonalTabProps, ILearnPersonalTabState> {

    public async componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));


        if (await this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                microsoftTeams.appInitialization.notifySuccess();
                this.setState({
                    entityId: context.entityId
                });
                this.updateTheme(context.theme);
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
                            <Header as="h2" content="Tab changee" />
                            <div>
                                <Text content={this.state.entityId} />
                            </div>

                            <div>
                                <Button onClick={() => alert("It worked!")}>A sample button</Button>
                            </div>

                            <div>
                                <Button primary onClick={() => alert("It worked!")}>A success button</Button>
                                <Button content="Profile" styles={{ backgroundColor: 'Green', boxShadow: '0 0 0 2px #01852e', ":hover": {backgroundColor: '#03ab3c'} }} />
                            </div>
                        </div>
                    </Flex.Item>
                    <Flex.Item styles={{
                        padding: ".8rem 0 .8rem .5rem"
                    }}>
                        <Text size="smaller" content="(C) Copyright Test" />
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }

    private updateComponentTheme = (setTeamsTheme: string = "default"): void => {
        let theme: ThemePrepared;
      
        switch (setTeamsTheme) {
          case "default":
            theme = teamsTheme;
            break;
          case "dark":
            theme = teamsDarkTheme;
            break;
          case "contrast":
            theme = teamsHighContrastTheme;
            break;
          default:
            theme = teamsTheme;
            break;
        }
        // update the state
        this.setState(Object.assign({}, this.state, {
            setTeamsTheme: theme
        }));
      }
}
