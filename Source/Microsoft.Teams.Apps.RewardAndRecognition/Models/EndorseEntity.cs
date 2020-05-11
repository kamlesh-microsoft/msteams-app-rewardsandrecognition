// <copyright file="EndorseEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    using System;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Endorse table storage entity.
    /// </summary>
    public class EndorseEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets team id.
        /// </summary>
        public string TeamId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets endorsed award name.
        /// </summary>
        public string EndorseForAward { get; set; }

        /// <summary>
        /// Gets or sets endorsed award id.
        /// </summary>
        public string EndorseForAwardId { get; set; }

        /// <summary>
        /// Gets or sets award cycle.
        /// </summary>
        public string AwardCycle { get; set; }

        /// <summary>
        /// Gets or sets endorsed to user principal name.
        /// </summary>
        public string EndorsedToPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets the Azure Active Directory Id of the endorsed user.
        /// </summary>
        public string EndorsedToObjectId { get; set; }

        /// <summary>
        /// Gets or sets the endorsed by principal name.
        /// </summary>
        public string EndorsedByPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets AAD object id of endorsed user.
        /// </summary>
        public string EndorsedByObjectId { get; set; }

        /// <summary>
        /// Gets or sets the date time when the award was endorsed.
        /// </summary>
        public DateTime EndorsedOn { get; set; }
    }
}
