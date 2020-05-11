// <copyright file="publishaward-table.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Table, Text, Button, Accordion } from "@fluentui/react-northstar";
import CheckboxBase from "./checkbox-base";
import "../styles/site.css";
import { useTranslation } from 'react-i18next';

interface IPublishAwardTableProps {
    showCheckbox: boolean,
    publishData: [],
    distinctAwards: [],
    onCheckBoxChecked: (nominationId: string, isChecked: boolean) => void,
    onChatButtonClick: (nominationDetails: any, t: any) => void
}

const PublishAwardTable: React.FunctionComponent<IPublishAwardTableProps> = props => {
    const { t } = useTranslation();
    const awardsTableHeader = {
        key: "header",
        items: props.showCheckbox === true ?
            [
                { content: < Text content={""} />, key: "check-box", className: "table-checkbox-cell" },
                { content: <Text weight="regular" content={t('nomineesTableHeaderText')} /> },
                { content: <Text weight="regular" content={t('nominationReasonTableHeaderText')} /> },
                { content: <Text weight="regular" content={t('endorsedByTableHeaderText')} /> },
                { content: <Text weight="regular" content={t('chatWithNominatorTableHeaderText')} /> }
            ]
            :
            [
                { content: <Text weight="regular" content={t('nomineesTableHeaderText')} /> },
                { content: <Text weight="regular" content={t('nominationReasonTableHeaderText')} /> },
            ],
    };

    let awardsTableRows = props.publishData.map((value: any, index) => (
        {
            key: value.AwardId,
            style: {},
            items: props.showCheckbox === true ?
                [
                    { content: <CheckboxBase onCheckboxChecked={props.onCheckBoxChecked} value={value.NominationId} />, key: index + "1", className: "table-checkbox-cell" },
                    { content: <Text content={value.NominatedToName} title={value.NominatedToName} />, key: index + "2", truncateContent: true },
                    { content: <Text content={value.ReasonForNomination} title={value.ReasonForNomination} />, key: index + "3", truncateContent: true },
                    { content: <Text content={value.EndorseCount} title={value.EndorseCount} />, key: index + "4", truncateContent: true },
                    {
                        content: <Button secondary onClick={() => props.onChatButtonClick(value, t)} title={value.NominatedByName} content={t('chatButtonText')} ></Button >
                    }
                ]
                :
                [
                    { content: <Text content={value.NominatedToName} title={value.NominatedToName} />, key: index + "2", truncateContent: true },
                    { content: <Text content={value.ReasonForNomination} title={value.ReasonForNomination} />, key: index + "3", truncateContent: true },
                ],
        }
    ));

    let panels = props.distinctAwards.map((value: any) => (
        {
            title: value.AwardName,
            content: <Table rows={awardsTableRows.filter(row => row.key === value.AwardId)} header={awardsTableHeader} className="table-cell-content" />
        }
    ));

    return (
        <div>
            <Accordion defaultActiveIndex={[0]} panels={panels} />
        </div>
    );
}

export default PublishAwardTable;