// <copyright file="reward-cycle.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import { Button, Checkbox, Divider, Flex, Input, RadioGroup, Text, Loader } from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { Fabric, Customizer } from 'office-ui-fabric-react/lib';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import * as React from "react";
import { useEffect, useState } from "react";
import { useTranslation } from "react-i18next";
import { getRewardCycle, setRewardCycle } from "../api/reward-cycle-api";
import { createBrowserHistory } from "history";
import { getApplicationInsightsInstance } from "../helpers/app-insights";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import { Occurrence } from "../models/occurence";
import { RewardCycleDetail } from "../models/reward-cycle";
import Constants from "../constants/constants";
import { DarkCustomizations } from "../helpers/theme/DarkCustomizations";
import { DefaultCustomizations } from "../helpers/theme/DefaultCustomizations";
let moment = require('moment');
initializeIcons();

const browserHistory = createBrowserHistory({ basename: "" });

interface IRewardCycleState {
    selectedValue: number,
    noOfOccurence: string | undefined,
    isReccurringChecked: boolean,
    error: string
}

interface ICycle {
    cycleId: string | undefined;
    rewardCycleStartDate: Date | null | undefined;
    rewardCycleEndDate: Date | null | undefined;
    numberOfOccurrences: number | undefined;
    teamId: string | undefined;
    isRecurring: number | undefined;
    rangeOfOccurrence: number | undefined;
    rangeOfOccurrenceEndDate: Date | null | undefined;
    cycleStatus: number | undefined;
    createdByPrincipalName: string | undefined;
    createdByObjectId: string | undefined;
    createdOn: Date | null | undefined;
    resultPublished: number | undefined;
}

interface IProps {
    teamId: string,
}

const controlClass = mergeStyleSets({
    control: {
        margin: '0 0 15px 0',
        maxWidth: '300px'
    }
});

const RewardCycle: React.FC<IProps> = props => {
    let search = window.location.search;
    let params = new URLSearchParams(search);
    let theme = params.get("theme");
    let telemetry = params.get("telemetry");
    let appInsights = getApplicationInsightsInstance(telemetry, browserHistory);
    let datePickerTheme;
    let userObjectId: string | undefined;
    let userEmail: string | undefined;
    if (theme === Constants.dark) { datePickerTheme = DarkCustomizations }
    else if (theme === Constants.contrast) { datePickerTheme = DarkCustomizations }
    else { datePickerTheme = DefaultCustomizations }

    const { t } = useTranslation();
    const [startDate, setStartDate] = useState<Date | null | undefined>(null);
    const [endDate, setEndDate] = useState<Date | null | undefined>(null);
    const [minEndDate, setMinEndDate] = useState<Date>(new Date(moment().add(7, 'd').format()));
    const [calendarDate, setCalendarDate] = useState<Date | null | undefined>(null);
    const [rewardCycleState, setRewardCycleState] =
        useState<IRewardCycleState>({ selectedValue: Occurrence.None, noOfOccurence: '', isReccurringChecked: false, error: '' });
    const [cycleState, setCycleState] =
        useState<ICycle>({
            cycleId: undefined,
            createdByObjectId: '',
            createdByPrincipalName: '',
            createdOn: undefined,
            isRecurring: undefined,
            numberOfOccurrences: undefined,
            rangeOfOccurrence: undefined,
            rangeOfOccurrenceEndDate: undefined,
            resultPublished: 0,
            rewardCycleEndDate: undefined,
            rewardCycleStartDate: undefined,
            cycleStatus: undefined,
            teamId: '',
        });

    const [loader, setLoader] = useState(true);
    const [submitLoading, setSubmitLoading] = useState(false);

    useEffect(() => {
        microsoftTeams.initialize();

        microsoftTeams.getContext((context) => {
            userObjectId = context.userObjectId;
            userEmail = context.upn;
        });

        const fetchData = async () => {
            appInsights.trackTrace({ message: `'getRewardCycle' - Initiated request`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
            let response = await getRewardCycle(props.teamId, true)

            if (response.status === 200 && response.data) {
                let rewardcycle = response.data;
                appInsights.trackTrace({ message: `'getRewardCycle' - Request success`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
                setCycleState({
                    cycleId: rewardcycle.cycleId,
                    createdByObjectId: rewardcycle.createdByObjectId,
                    createdByPrincipalName: rewardcycle.createdByPrincipalName,
                    createdOn: rewardcycle.createdOn,
                    isRecurring: rewardcycle.isRecurring,
                    numberOfOccurrences: rewardcycle.numberOfOccurrences,
                    rangeOfOccurrence: rewardcycle.rangeOfOccurrence,
                    rangeOfOccurrenceEndDate: rewardcycle.rangeOfOccurrenceEndDate,
                    resultPublished: rewardcycle.resultPublished,
                    rewardCycleEndDate: rewardcycle.rewardCycleEndDate,
                    rewardCycleStartDate: rewardcycle.rewardCycleStartDate,
                    cycleStatus: rewardcycle.rewardCycleState,
                    teamId: props.teamId
                });

                if (rewardcycle.rewardCycleStartDate) {
                    setStartDate(new Date(rewardcycle.rewardCycleStartDate));
                    setMinEndDate(new Date(moment(rewardcycle.rewardCycleStartDate).add(7, 'd').format()));
                }
                if (rewardcycle.rewardCycleEndDate) { setEndDate(new Date(rewardcycle.rewardCycleEndDate)); }
                if (rewardcycle.rangeOfOccurrenceEndDate) {
                    setCalendarDate(new Date(rewardcycle.rangeOfOccurrenceEndDate));
                }
                setRewardCycleState({ isReccurringChecked: rewardcycle.isRecurring === 1 ? true : false, noOfOccurence: rewardcycle.numberOfOccurrences !== 0 ? rewardcycle.numberOfOccurrences : undefined, selectedValue: rewardcycle.rangeOfOccurrence!, error: '' });
            }
            else {
                appInsights.trackTrace({ message: `'getRewardCycle' - Request failed`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
            }
            setLoader(false);
        };
        fetchData();
    }, []);

    /**
     * Handle change event for cycle start date picker.
     * @param date | cycle start date.
     */
    const onSelectStartDate = (date: Date | null | undefined): void => {
        setStartDate(date);
        setMinEndDate(new Date(moment(date).add(7, 'd').format()));
    };

    /**
     * Handle change event for cycle end date picker.
     * @param date | cycle end date.
     */
    const onSelectEndDate = (date: Date | null | undefined): void => {
        setEndDate(date);
    };

    /**
     * Handle change event for end by date picker.
     * @param date | end by date.
     */
    const onSelectCalendarDate = (date: Date | null | undefined): void => {
        setCalendarDate(date);
    };

    /**
     * Handling input change event.
     * @param event 
     */
    const handleInputChange = (event: any): void => {
        let p = event.target;
        setRewardCycleState({ ...rewardCycleState, [p.name]: p.value })

    }

    /**
     * Handling check box change event.
     * @param isChecked | boolean value.
     */
    const handleCheckBoxChange = (isChecked: boolean): void => {
        setRewardCycleState({ isReccurringChecked: !isChecked, selectedValue: Occurrence.None, noOfOccurence: '', error: rewardCycleState.error })
        setCalendarDate(null);
    }

    /**
     * This method is used to handle done button click by setting the reward cycle and sending notification card to the team.
     * @param start | cycle start date.
     * @param end | cycle end date.
     */
    const onSetCycle = async (start: Date | null | undefined, end: Date | null | undefined): Promise<void> => {

        if (!(start && end)) {
            setRewardCycleState({ ...rewardCycleState, error: t('requiredDatesError') });
            return;
        }

        if (rewardCycleState.selectedValue === Occurrence.EndAfter) {
            if (parseInt(rewardCycleState.noOfOccurence!) <= 0) {
                setRewardCycleState({ ...rewardCycleState, error: t('noOfOccurrenceError') });
                return;
            }

            if (!rewardCycleState.noOfOccurence || rewardCycleState.noOfOccurence === '') {
                setRewardCycleState({ ...rewardCycleState, error: t('noOfOccurrenceError') });
                return;
            }
        }

        let startCycle = moment(start)
            .set('hour', moment().hour())
            .set('minute', moment().minute())
            .set('second', moment().second());

        let endCycle = moment(end)
            .set('hour', moment().hour())
            .set('minute', moment().minute())
            .set('second', moment().second());

        let endByDate = undefined;
        if (rewardCycleState.selectedValue === Occurrence.EndBy) {
            if (!calendarDate || calendarDate === null) {
                setRewardCycleState({ ...rewardCycleState, error: t('requiredEndByDate') });
                return;
            }

            endByDate = moment.utc(moment(calendarDate)
                .set('hour', moment().hour())
                .set('minute', moment().minute())
                .set('second', moment().second()));

        }
        setSubmitLoading(true);

        let rewardCycleDetail: RewardCycleDetail = {
            RewardCycleStartDate: moment.utc(startCycle),
            RewardCycleEndDate: moment.utc(endCycle),
            IsRecurring: rewardCycleState.isReccurringChecked ? 1 : 0,
            NumberOfOccurrences: rewardCycleState.selectedValue === Occurrence.EndAfter ? parseInt(rewardCycleState.noOfOccurence!) : 0,
            ResultPublished: cycleState.resultPublished,
            RewardCycleState: start.getDate() === new Date().getDate() ? 1 : 0,
            CycleId: cycleState.cycleId,
            CreatedByPrincipalName: userEmail,
            RangeOfOccurrence: rewardCycleState.selectedValue,
            RangeOfOccurrenceEndDate: endByDate,
            TeamId: props.teamId,
            CreatedByObjectId: userObjectId,
            CreatedOn: cycleState.createdOn,
            ResultPublishedOn: null,
        };

        appInsights.trackTrace({ message: `'setRewardCycle' - Initiated request`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
        let response = await setRewardCycle(rewardCycleDetail);
        if (response.status === 200 && response.data) {
            appInsights.trackTrace({ message: `'setRewardCycle' - Request success`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
            let toBot = {
                Command: Constants.NominateAwardsCommand,
                RewardCycleStartDate: rewardCycleDetail.RewardCycleStartDate,
                RewardCycleEndDate: rewardCycleDetail.RewardCycleEndDate,
                RewardCycleId: response.data.cycleId,
                TeamId: props.teamId
            };
            let obj = JSON.parse(JSON.stringify(toBot));
            microsoftTeams.tasks.submitTask(obj);
        }
        else {
            appInsights.trackTrace({ message: `'setRewardCycle' - Request failed`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
            setRewardCycleState({ ...rewardCycleState, error: t('errorText') });
        }
        setSubmitLoading(false);
    };

    /**
     * Handle radio group change event.
     * @param e | event
     * @param props | props
     */
    const handleChange = (e: any, props: any) => {
        setRewardCycleState({ noOfOccurence: '', isReccurringChecked: rewardCycleState.isReccurringChecked, selectedValue: props.value, error: '' });
        setCalendarDate(null);
    }

    /**
     * Radio group items.
     */
    const getItems = () => {
        return [
            {
                key: 'none',
                label: t('noEndDate'),
                value: Occurrence.None,
            },
            {
                key: 'endby',
                label: (
                    <div style={{ marginTop: "0.5rem" }}>
                        <Text content={t('endBy')} />
                        <Fabric>
                            <Customizer {...datePickerTheme}>
                                <DatePicker
                                    className={controlClass.control}
                                    allowTextInput={true}
                                    showMonthPickerAsOverlay={true}
                                    minDate={endDate!}
                                    isMonthPickerVisible={true}
                                    value={calendarDate!}
                                    onSelectDate={onSelectCalendarDate}
                                />
                            </Customizer>
                        </Fabric>
                    </div>
                ),
                value: Occurrence.EndBy,
            },
            {
                key: 'endafter',
                label: (
                    <Flex column>
                        <Text content={t('endAfter')} />
                        <Input type="number"
                            min={1}
                            name="noOfOccurence"
                            value={rewardCycleState.noOfOccurence!}
                            onChange={handleInputChange}
                            defaultValue={undefined}
                        />
                    </Flex>
                ),
                value: Occurrence.EndAfter,
            }
        ];
    }


    return (

        <div>
            {loader ?
                <div className="tab-container">
                    <Loader />
                </div>
                :
                <div>
                    <div className="tab-container">
                        {rewardCycleState.error && <Flex hAlign="center"><Text content={rewardCycleState.error} error /></Flex>}
                        <Flex gap="gap.small">
                            <Flex.Item size="size.half">
                                <div>
                                    <Fabric>
                                        <Customizer {...datePickerTheme}>
                                            <DatePicker
                                                className={controlClass.control}
                                                label={t('startDate')}
                                                isRequired={true}
                                                allowTextInput={true}
                                                showMonthPickerAsOverlay={true}
                                                minDate={new Date()}
                                                isMonthPickerVisible={true}
                                                value={startDate!}
                                                onSelectDate={onSelectStartDate}
                                            />
                                        </Customizer>
                                    </Fabric>
                                </div>
                            </Flex.Item>
                            <Flex.Item size="size.half">
                                <div>
                                    <Fabric>
                                        <Customizer {...datePickerTheme}>
                                            <DatePicker
                                                className={controlClass.control}
                                                label={t('endDate')}
                                                isRequired={true}
                                                allowTextInput={true}
                                                minDate={minEndDate}
                                                isMonthPickerVisible={true}
                                                showMonthPickerAsOverlay={true}
                                                value={endDate!}
                                                onSelectDate={onSelectEndDate}
                                            />
                                        </Customizer>
                                    </Fabric>
                                </div>
                            </Flex.Item>
                        </Flex>
                        <Divider />
                        <Flex>
                            <Text weight="semibold" content="Recurring" />
                            <Flex.Item push>
                                <Checkbox toggle
                                    checked={rewardCycleState.isReccurringChecked}
                                    onChange={() => handleCheckBoxChange(rewardCycleState.isReccurringChecked)}
                                />
                            </Flex.Item>
                        </Flex>
                        <Divider styles={{ marginBottom: "1rem" }} />
                        {rewardCycleState.isReccurringChecked && <Flex column gap="gap.small">
                            <div>
                                <Text content={t('rangeOfOccurences')} />
                                <RadioGroup vertical
                                    defaultCheckedValue={rewardCycleState.selectedValue}
                                    items={getItems()}
                                    onCheckedValueChange={handleChange}
                                />
                            </div>
                        </Flex>}
                    </div>
                    <div className="tab-footer">
                        <Flex hAlign="end">
                            <Button primary content={t('doneButtonText')} onClick={() => onSetCycle(startDate, endDate)} loading={submitLoading} />
                        </Flex>
                    </div>
                </div>
            }
        </div>
    );
}

export default RewardCycle;