/*
    <copyright file="award.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

export class AwardDetails {
    AwardId: string | undefined;
    AwardName: string | undefined;
    AwardDescription: string | undefined;
    AwardLink?: string | undefined;
    TeamId: string | undefined;
    CreatedBy: string | undefined;
    CreatedOn: Date | undefined;
}