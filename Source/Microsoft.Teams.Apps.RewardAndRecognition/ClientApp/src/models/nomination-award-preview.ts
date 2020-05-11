/*
    <copyright file="nomination-award-preview.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

export class NominationAwardPreview {
    NominatedByName: string | "" = "";
    ImageUrl: string | "" = "";
    ReasonForNomination: string | undefined;
    AwardRecipients: any[] = [];
    AwardId: string | undefined;
    AwardName?: string | undefined;
    TeamId: string | undefined;
    NominatedToPrincipalName: any[] = [];
    NominatedToObjectId: any[] = [];
    NominatedByPrincipalName: string | undefined;
    NominatedByObjectId?: string | null;
    telemetry?: any = null;
    locale?: string | null;
    theme: string | null = null;
}