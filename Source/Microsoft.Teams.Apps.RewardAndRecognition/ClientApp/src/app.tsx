// <copyright file="app.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
import * as React from "react";
import { AppRoute } from "./router/router";
import Resources from "./constants/resources";
import * as microsoftTeams from "@microsoft/teams-js";
import { Provider, themes } from "@fluentui/react-northstar";
import { TeamsThemeContext, getContext, ThemeStyle } from 'msteams-ui-components-react';
export interface IAppState {
    theme: string;
    themeStyle: number;
}
export default class App extends React.Component<{}, IAppState> {
    theme?: string | null;
    constructor(props: any) {
        super(props);
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.theme = params.get("theme");
        this.state = {
            theme: this.theme ? this.theme : Resources.default,
            themeStyle: ThemeStyle.Light,
        }
    }
    public componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            let theme = context.theme || "";
            this.updateTheme(theme);
            this.setState({
                theme: theme
            });
        });
        microsoftTeams.registerOnThemeChangeHandler((theme) => {
            this.updateTheme(theme);
            this.setState({
                theme: theme,
            }, () => {
                this.forceUpdate();
            });
        });
    }
    public setThemeComponent = () => {
        if (this.state.theme === "dark") {
            return (
                <Provider theme={themes.teamsDark}>
                    <div className="darkContainer">
                        {this.getAppDom()}
                    </div>
                </Provider>
            );
        }
        else if (this.state.theme === "contrast") {
            return (
                <Provider theme={themes.teamsHighContrast}>
                    <div className="highContrastContainer">
                        {this.getAppDom()}
                    </div>
                </Provider>
            );
        } else {
            return (
                <Provider theme={themes.teams}>
                    <div className="defaultContainer">
                        {this.getAppDom()}
                    </div>
                </Provider>
            );
        }
    }

    private updateTheme = (theme: string) => {
        if (theme === "dark") {
            this.setState({
                themeStyle: ThemeStyle.Dark
            });
        } else if (theme === "contrast") {
            this.setState({
                themeStyle: ThemeStyle.HighContrast
            });
        } else {
            this.setState({
                themeStyle: ThemeStyle.Light
            });
        }
    }

    public getAppDom = () => {
        const context = getContext({
            baseFontSize: 10,
            style: this.state.themeStyle
        });
        return (
            <TeamsThemeContext.Provider value={context}>
                <div className="appContainer">
                    <AppRoute />
                </div>
            </TeamsThemeContext.Provider>);
    }
    /**
    * Renders the component
    */
    public render(): JSX.Element {
        return (
            <div>
                {this.setThemeComponent()}
            </div>
        );
    }
}