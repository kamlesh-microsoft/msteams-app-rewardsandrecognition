/*
    <copyright file="router.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { BrowserRouter, Route, Switch } from "react-router-dom";
import ConfigurationAdminPage from '../components/config-admin-page';
import NominateAwards from '../components/nominate-awards';
import AwardsTab from "../components/manage-award-tab";
import { Suspense } from "react";
import "../i18n";
import PublishAward from "../components/publish-award";
import SignInPage from "../components/signin/signin";
import SignInSimpleStart from "../components/signin/signin-start";
import SignInSimpleEnd from "../components/signin/signin-end";
import DiscoverWrapperPage from "../components/publish-award";
import ErrorPage from '../components/error-page';
import { Loader } from "@fluentui/react-northstar";
import ConfigTab from "../components/config-tab";

export const AppRoute: React.FunctionComponent<{}> = () => {

    return (
        <Suspense fallback={<div> <Loader /></div>}>
            <BrowserRouter>
                <Switch>
                    <Route exact path='/config-admin-page' component={ConfigurationAdminPage} />
                    <Route exact path='/nominate-awards' component={NominateAwards} />
                    <Route exact path="/awards-tab" component={AwardsTab} />
                    <Route path="/publish-award" component={PublishAward} />
                    <Route exact path="/" component={DiscoverWrapperPage} />
                    <Route exact path="/discover" component={DiscoverWrapperPage} />
                    <Route exact path="/signin" component={SignInPage} />
                    <Route exact path="/signin-simple-start" component={SignInSimpleStart} />
                    <Route exact path="/signin-simple-end" component={SignInSimpleEnd} />
                    <Route exact path="/error" component={ErrorPage} />
                    <Route exact path="/configTab" component={ConfigTab} />
                </Switch>
            </BrowserRouter>
        </Suspense>
    );
};


