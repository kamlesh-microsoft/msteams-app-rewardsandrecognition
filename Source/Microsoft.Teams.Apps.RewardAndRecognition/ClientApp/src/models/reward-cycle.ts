/*
    <copyright file="reward-cycle.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

export class RewardCycleDetail {
    CycleId: string | undefined;
    RewardCycleStartDate: Date | null | undefined;
    RewardCycleEndDate: Date | null | undefined;
    NumberOfOccurrences: number | undefined;
    TeamId: string | undefined;
    IsRecurring: number | undefined;
    RangeOfOccurrence: number | undefined;
    RangeOfOccurrenceEndDate: Date | null | undefined;
    RewardCycleState: number | undefined;
    CreatedByPrincipalName: string | undefined;
    CreatedByObjectId: string | undefined;
    CreatedOn: Date | null | undefined;
    ResultPublished: number | undefined;
    ResultPublishedOn: Date | null | undefined;
}