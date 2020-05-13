// <copyright file="publish-awards.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import * as microsoftTeams from "@microsoft/teams-js";
import { createBrowserHistory } from "history";
import { Button, Loader, Flex, Text, themes, Dialog, Alert, Icon } from "@fluentui/react-northstar";
import { getAllAwardNominations, publishAwardNominations } from "../api/nominate-awards-api";
import { getRewardCycle, setRewardCycle } from "../api/reward-cycle-api";
import { sendWinnerNotification } from "../api/notification-api";
import { getMembersInTeam, getUserRoleInTeam } from "../api/configure-admin-api";
import { getBotSetting } from "../api/setting-api";
import { getAllAwards } from "../api/awards-api";
import PublishAwardTable from "./publishaward-table";
import ApprovedAwardTable from "./result-table";
import "../styles/site.css";
import { RewardCycleState, RewardPublishState } from "../models/award-cycle-state";
import Resources from "../constants/resources";
import { getApplicationInsightsInstance } from "../helpers/app-insights";
import { ResultDetails } from "../models/result";
import { withTranslation, WithTranslation } from "react-i18next";
import { navigateToErrorPage, validateUserPartOfTeam } from "../helpers/utility";
let moment = require('moment');

interface IState {
    Loader: boolean,
    isUserPartOfTeam: boolean,
    theme: string | null,
    themeStyle: any;
    errorMessage: string | null;
    selectedNominees: string[];
    publishAwardDataSet: any;
    distinctAwards: any;
    pubishResults: any;
    awardWinner: Array<ResultDetails>;
    activeAwardCycle: any;
    isNominationPriviewAvailable: boolean;
    openDialog: boolean;
    isWinnerCardSent: boolean;
    isAdminUser: boolean;
    isPublishedAwards: boolean;
}

const browserHistory = createBrowserHistory({ basename: "" });

/** Component for displaying on publish award details. */
class PublishAward extends React.Component<WithTranslation, IState>
{
    locale?: string | null;
    telemetry?: any = null;
    appInsights: any;
    theme: string | null = null;
    userEmail?: any = null;
    userObjectId?: string = "";
    teamId?: string | null;
    activeCycleId: string | "";
    botId: string;
    appBaseUrl: string;
    currentAwardCycleDateRange?: string | "";

    constructor(props: any) {
        super(props);
        this.state = {
            Loader: true,
            isUserPartOfTeam: false,
            theme: this.theme ? this.theme : Resources.default,
            themeStyle: themes.teams,
            errorMessage: "",
            selectedNominees: [],
            publishAwardDataSet: [],
            distinctAwards: [],
            pubishResults: [],
            awardWinner: [],
            activeAwardCycle: {},
            isNominationPriviewAvailable: false,
            openDialog: false,
            isWinnerCardSent: false,
            isAdminUser: false,
            isPublishedAwards: false,
        };

        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.theme = params.get("theme");
        this.locale = params.get("locale");
        this.teamId = params.get("teamId");
        this.botId = '';
        this.appBaseUrl = window.location.origin;
        this.currentAwardCycleDateRange = "";
        this.activeCycleId = "";
    }

    /** Called once component is mounted. */
    async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId;
            this.userEmail = context.upn;
            this.teamId = context.teamId;
        });

        let flag = await validateUserPartOfTeam(this.teamId!, this.userObjectId!)
        if (flag) {
        await this.getBotSetting();
        this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);
        await this.validateUserProfileInTeam();
        await this.getRewardCycle();
        if (this.activeCycleId !== undefined || this.activeCycleId !== "") {
            await this.getPublishAwardDetails();
            }
        }
        else {
            navigateToErrorPage('');
        }
    }

    /**
   *Get bot id from API
   */
    async getBotSetting() {
        let response = await getBotSetting()
        if (response.status === 200 && response.data) {
            let settings = response.data;
            this.telemetry = settings.instrumentationKey;
            this.botId = settings.botId;
        }
        else {
            navigateToErrorPage(response.status);
        }
    }

    submitHandler = async (err, result) => {
        this.appInsights.trackTrace(`Submit handler - err: ${err} - result: ${result}`);
        await this.updatePublishState();
    };

    /**
    *Get award nomination details from API
    */
    async getRewardCycle() {
        this.appInsights.trackTrace({ message: `'getRewardCycle' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let response = await getRewardCycle(this.teamId!, this.state.isAdminUser)
        if (response.status === 200 && response.data) {
            this.appInsights.trackTrace({ message: `'getRewardCycle' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            let rewardcycle = response.data;
            this.currentAwardCycleDateRange = (moment(rewardcycle.rewardCycleStartDate).format("MMMM Do YYYY") + " to " + moment(rewardcycle.rewardCycleEndDate).format("MMMM Do YYYY")).toString();
            this.activeCycleId = rewardcycle.cycleId;

            this.setState({
                activeAwardCycle: rewardcycle
            });
        }
    }

    /**
    *Get award nomination details from API
    */
    async validateUserProfileInTeam() {
        this.appInsights.trackTrace({ message: `'getTeamMembersInTeam' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let teamMembers = await getMembersInTeam(this.teamId!);
        if (teamMembers.status === 200 && teamMembers.data) {
            this.appInsights.trackTrace({ message: `'getPublishAwardDetails' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });

            let member = teamMembers.data.find(element => element.aadobjectid === this.userObjectId);
            if (member !== null || member !== undefined) {
                this.setState({
                    isUserPartOfTeam: true
                });
                // check user role in team
                this.appInsights.trackTrace({ message: `'getUserRoleInTeam' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
                let adminDetails = await getUserRoleInTeam(this.teamId!);
                if (adminDetails.status === 200 && adminDetails.data) {
                    this.appInsights.trackTrace({ message: `'getUserRoleInTeam' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });

                    if (adminDetails.data.AdminObjectId === this.userObjectId) {
                        this.setState({
                            isAdminUser: true,
                            isPublishedAwards: false
                        });
                    }
                    else {
                        this.setState({
                            isAdminUser: false,
                            isPublishedAwards: true
                        });
                    }
                }
            }
        }
        else {
            this.appInsights.trackTrace({ message: `'getTeamMembersInTeam' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            navigateToErrorPage(teamMembers.status);
        }
    }

    /**
    *Navigate to manage award tab
    */
    onManageAwardButtonClick = (t: any) => {
        microsoftTeams.tasks.startTask({
            completionBotId: this.botId,
            title: t('manageAwardButtonText'),
            height: 700,
            width: 800,
            url: `${this.appBaseUrl}/awards-tab?telemetry=${this.telemetry}&theme=${this.theme}&teamId=${this.teamId}&locale=${this.locale}`,
        }, this.submitHandler);
    }

    /**
    *Navigate to configure admin
    */
    onConfigureAdminButtonClick = (t: any) => {
        microsoftTeams.tasks.startTask({
            completionBotId: this.botId,
            title: t('configureAdminTitle'),
            height: 460,
            width: 600,
            url: `${this.appBaseUrl}/config-admin-page?telemetry=${this.telemetry}&teamId=${this.teamId}&theme=${this.theme}&locale=${this.locale}`,
        }, this.submitHandler);
    }

    onChatButtonClick = (nominationDetails: any, t: any) => {
        let nominee = nominationDetails.NominatedByName;
        let msg = t('chatMessage') + nominee;
        let url = `https://teams.microsoft.com/l/chat/0/0?users=${nominee}&message=${msg}`;
        microsoftTeams.executeDeepLink(url);
    }

    /**
    *Get selected nominations
    */
    onNominationSelected = (nominationId: string, isSelected: boolean) => {
        if (nominationId !== "") {
            let selectNominees = this.state.selectedNominees;
            let selectedAwardWinner = this.state.awardWinner;
            let nomination = this.state.publishAwardDataSet.filter(row => row.NominationId === nominationId).shift();

            if (isSelected) {
                selectNominees.push(nominationId);

                let results: ResultDetails = {
                    AwardId: nomination.AwardId,
                    AwardName: nomination.AwardName,
                    NominationId: nominationId,
                    WinnerCount: 0,
                    TeamId: this.teamId!,
                    NominatedToName: nomination.NominatedToName,
                    NominatedToObjectId: nomination.NominatedToObjectId,
                    NominatedToPrincipalName: nomination.NominatedToPrincipalName,
                    AwardLink: this.state.distinctAwards.filter(row => row.AwardId === nomination.AwardId).shift().awardLink,
                    AwardCycle: this.currentAwardCycleDateRange,
                };
                selectedAwardWinner.push(results);
            }
            else {
                selectedAwardWinner.splice(selectNominees.indexOf(nominationId), 1);
                selectNominees.splice(selectNominees.indexOf(nominationId), 1);
            }

            this.setState({
                selectedNominees: selectNominees
            })

            this.setState({
                awardWinner: selectedAwardWinner
            })
        }
    }

    /**
    *Show publish award confirmation window
    */
    onPublishResultButtonClick = async (t: any) => {
        this.setState({
            selectedNominees: [],
        })
        let response = await this.publishAwards();
        if (response) {
            let notifyResponse = await sendWinnerNotification(this.state.awardWinner);
            if (notifyResponse.status === 200) {

                // Update active award cycle to published
                let awardCycle = this.state.activeAwardCycle;
                awardCycle.resultPublished = RewardPublishState.Published;
                awardCycle.rewardCycleState = RewardCycleState.InActive;
                awardCycle.resultPublishedOn = new Date();
                let awardPublish = await setRewardCycle(awardCycle);
                if (awardPublish.status == 200) {
                    this.setState({ openDialog: true })
                    this.setState({ isWinnerCardSent: true })
                }

                await this.getRewardCycle();
                await this.getPublishAwardDetails();
                await this.validateUserProfileInTeam();
            }
        }
        else {
            this.setState({ isWinnerCardSent: false })
        }
    }

    updatePublishState = async () => {
        this.setState({ openDialog: false, isNominationPriviewAvailable: false})
    }

    /**
    *Publish award nominations from API
    */
    async publishAwards() {
        this.appInsights.trackTrace({ message: `'publishAwards' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let awards = await publishAwardNominations(this.teamId!, this.state.selectedNominees.toString());
        if (awards.status === 200 && awards.data) {
            this.appInsights.trackTrace({ message: `'publishAwards' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            return true;
        }
        else {
            this.appInsights.trackTrace({ message: `'getPublishAwardDetails' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            return false;
        }
    }

    /**
    *Get award nomination details from API
    */
    async getPublishAwardDetails() {
        this.appInsights.trackTrace({ message: `'getPublishAwardDetails' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let nominations = await getAllAwardNominations(this.teamId!, this.state.isPublishedAwards, this.activeCycleId!);
        if (nominations.status === 200 && nominations.data) {
            this.appInsights.trackTrace({ message: `'getPublishAwardDetails' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            if (nominations.data !== null && nominations.data.length > 0) {
                this.setState({
                    publishAwardDataSet: nominations.data,
                    isNominationPriviewAvailable: true
                });
            }
            else {
                this.setState({
                    isNominationPriviewAvailable: false
                });
            }

            this.appInsights.trackTrace({ message: `'getAwards' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            let awards = await getAllAwards(this.teamId!);
            if (awards.status === 200 && awards.data) {
                this.appInsights.trackTrace({ message: `'getAllAwards' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });

                this.setState({
                    distinctAwards: awards.data
                });
            }
        }
        else {
            this.appInsights.trackTrace({ message: `'getPublishAwardDetails' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        }
        this.setState({
            Loader: false
        });
    }

    openPublishDialog = () => this.setState({ openDialog: true })
    closePublishDialog = () => this.setState({ openDialog: false })

    /**
   *Get wrapper page for selected awards for publish
   */
    private getPublishConfirmationPage = () => {
        if (!this.state.Loader) {
            return (
                <div>
                    <ApprovedAwardTable awardWinner={this.state.awardWinner}
                        distinctAwards={this.state.distinctAwards}
                    />
                </div>
            );
        }
    }

    private pageHeader = (t: any) => {
        return (<Flex gap="gap.small" >
            {this.currentAwardCycleDateRange != "" && <Text weight="bold" align="center" content={t('rewardCycleText') + this.currentAwardCycleDateRange} />}
            {this.state.isAdminUser &&
                <>
                    <Flex.Item push>
                    <Button secondary className="publish-award-button" onClick={() => this.onConfigureAdminButtonClick(t)} content={t('configureAdminTitle')}></Button>
                    </Flex.Item>
                <Button className="publish-award-button" content={t('manageAwardButtonText')} onClick={() => this.onManageAwardButtonClick(t)} />
                <Dialog
                    cancelButton={t('cancelButtonText')}
                    confirmButton={<Button primary content={t('confirmButtonText')}></Button>}
                    content={this.getPublishConfirmationPage()}
                    header={t('publishResultHeaderText')}
                    trigger={<Button primary disabled={this.state.selectedNominees.length === 0} content={t('grantAwardButtonText')}></Button>}
                    onConfirm={() => this.onPublishResultButtonClick(t)} />
                {this.state.openDialog &&
                    <Dialog
                        open={this.state.openDialog}
                        header={t('publishResultHeaderText')}
                        content={this.state.isWinnerCardSent ? t('resultPublishSuccessMessage') : t('resultPublishFailedMessage')}
                        confirmButton={<Button primary onClick={this.updatePublishState} content={t('buttonTextOk')}></Button>}
                        headerAction={<Icon name="close" onClick={this.closePublishDialog} />}
                    />}
                </>}
        </Flex>);
    }

    /**
   *Get wrapper for page which acts as container for all child components
   */
    private getWrapperPage = (t: any) => {
        if (this.state.Loader) {
            return (
                <div className="loader">
                    <Loader />
                </div>
            );
        } else if (!this.state.Loader && this.state.isUserPartOfTeam && this.state.isNominationPriviewAvailable) {
            return (
                <div>
                    <PublishAwardTable showCheckbox={this.state.isAdminUser}
                        publishData={this.state.publishAwardDataSet}
                        distinctAwards={this.state.distinctAwards}
                        onCheckBoxChecked={this.onNominationSelected}
                        onChatButtonClick={this.onChatButtonClick}
                    />
                </div>
            );
        }
        else if (!this.state.Loader && !this.state.isNominationPriviewAvailable) {
            return (<div className="publish-award-footer-page"><Alert content={t('nominationPreviewMessage')} /></div>)
        }
    }

    /**
  * Renders the component
  */
    public render() {
        const { t } = this.props;
        return (
            <div className="page-container">
                {this.pageHeader(t)}
                <div>
                    {this.getWrapperPage(t)}
                </div>
            </div>
        );
    }
}

export default withTranslation()(PublishAward);