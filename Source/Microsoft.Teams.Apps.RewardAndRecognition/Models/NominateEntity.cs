// <copyright file="NominateEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using Microsoft.Azure.Search;
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// Class contains details of nomination.
    /// </summary>
    public class NominateEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets team id.
        /// </summary>
        [IsFilterable]
        [JsonProperty("TeamId")]
        public string TeamId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets unique identifier of Nomination.
        /// </summary>
        [Key]
        [JsonProperty("NominationId")]
        public string NominationId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets name of award.
        /// </summary>
        [IsSearchable]
        [IsFilterable]
        [JsonProperty("AwardName")]
        public string AwardName { get; set; }

        /// <summary>
        /// Gets or sets unique identifier of award id.
        /// </summary>
        [JsonProperty("AwardId")]
        public string AwardId { get; set; }

        /// <summary>
        /// Gets or sets award image URL.
        /// </summary>
        [JsonProperty("AwardImageLink")]
        public string AwardImageLink { get; set; }

        /// <summary>
        /// Gets or sets date on when the nomination is set.
        /// </summary>
        [JsonProperty("NominatedOn")]
        public DateTime? NominatedOn { get; set; }

        /// <summary>
        /// Gets or sets nominee name.
        /// </summary>
        [IsSearchable]
        [IsFilterable]
        [JsonProperty("NominatedToName")]
        public string NominatedToName { get; set; }

        /// <summary>
        /// Gets or sets User principal name of nominee.
        /// </summary>
        [IsSearchable]
        [IsFilterable]
        [JsonProperty("NominatedToPrincipalName")]
        public string NominatedToPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets User principal name of nominator.
        /// </summary>
        [IsSearchable]
        [JsonProperty("NominatedByName")]
        public string NominatedByName { get; set; }

        /// <summary>
        /// Gets or sets User principal name of nominator.
        /// </summary>
        [IsSearchable]
        [IsFilterable]
        [JsonProperty("NominatedByPrincipalName")]
        public string NominatedByPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets AAD object id of nominator.
        /// </summary>
        [JsonProperty("NominatedByObjectId")]
        public string NominatedByObjectId { get; set; }

        /// <summary>
        /// Gets or sets AAD object id of nominee.
        /// </summary>
        [JsonProperty("NominatedToObjectId")]
        public string NominatedToObjectId { get; set; }

        /// <summary>
        /// Gets or sets note that was given to the nominee.
        /// </summary>
        [JsonProperty("ReasonForNomination")]
        public string ReasonForNomination { get; set; }

        /// <summary>
        /// Gets or sets reward cycle identifier.
        /// </summary>
        [IsFilterable]
        [JsonProperty("RewardCycleId")]
        public string RewardCycleId { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether group nomination or not.
        /// </summary>
        [JsonProperty("IsGroupNomination")]
        public string IsGroupNomination { get; set; }

        /// <summary>
        /// Gets or sets group name.
        /// </summary>
        [JsonProperty("GroupName")]
        public string GroupName { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether award Granted or not.
        /// </summary>
        [JsonProperty("AwardGranted")]
        public bool AwardGranted { get; set; }

        /// <summary>
        /// Gets or sets a date time of award publish.
        /// </summary>
        [JsonProperty("AwardPublishedOn")]
        public DateTime? AwardPublishedOn { get; set; }

        /// <summary>
        /// Gets time stamp from storage table.
        /// </summary>
        [IsSortable]
        [JsonProperty("Timestamp")]
        public new DateTimeOffset Timestamp => base.Timestamp;
    }
}
