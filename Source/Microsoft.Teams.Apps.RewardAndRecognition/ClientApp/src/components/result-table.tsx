// <copyright file="result-table.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Table, Text } from "@fluentui/react-northstar";
import "../styles/site.css";
import { useTranslation } from 'react-i18next';

interface IApprovedAwardTableProps {
    awardWinner: any,
    distinctAwards: any
}

const ApprovedAwardTable: React.FunctionComponent<IApprovedAwardTableProps> = props => {
    const { t } = useTranslation();
    const awardsTableHeader = {
        key: "header",
        items:
            [
                { content: <Text weight="regular" content={t('awardName')} /> },
                { content: <Text weight="regular" content={t('winnersCountText')} /> },
            ],
    };

    let awardsTableRows = props.distinctAwards.map((value: any, index) => (
        {
            key: value.AwardId,
            style: {},
            items:
                [
                    { content: <Text content={value.AwardName} title={value.AwardName} />, key: index + "1", truncateContent: true },
                    { content: <Text content={props.awardWinner.filter(a => a.AwardId === value.AwardId).length} title={props.awardWinner.filter(a => a.AwardId === value.AwardId).length} />, key: index + "2", truncateContent: true },
                ]
        }
    ));

    return (
        <div>
            <Table rows={awardsTableRows} header={awardsTableHeader} className="table-cell-content" />
        </div>
    );
}

export default ApprovedAwardTable;