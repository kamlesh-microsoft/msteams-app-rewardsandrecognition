/*
    <copyright file="nominate-entity.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

export class NominateEntity {
    AwardId: string | undefined;
    RewardCycleId: string | undefined;
    AwardName?: string | undefined;
    ReasonForNomination?: string | undefined;
    TeamId: string | undefined;
    NominatedOn: Date | undefined;
    NominatedToName?: string | undefined;
    NominatedToPrincipalName: string | undefined;
    NominatedToObjectId: string | undefined;
    NominatedByName?: string | undefined;
    NominatedByPrincipalName: string | undefined;
    NominatedByObjectId?: string | null;
    IsGroupNomination?: string | undefined;
    GroupName?: string | undefined;
    AwardImageLink?: string | undefined;
}