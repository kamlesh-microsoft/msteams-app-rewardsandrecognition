// <copyright file="TaskModuleResponseDetails.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    using System;
    using Newtonsoft.Json;

    /// <summary>
    /// This class process to get task module response details.
    /// </summary>
    public class TaskModuleResponseDetails
    {
        /// <summary>
        /// Gets or sets name of admin name for the team.
        /// </summary>
        [JsonProperty("AdminName")]
        public string AdminName { get; set; }

        /// <summary>
        /// Gets or sets admin user principal name.
        /// </summary>
        [JsonProperty("AdminPrincipalName")]
        public string AdminPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets Note that was given to the team.
        /// </summary>
        [JsonProperty("NoteForTeam")]
        public string NoteForTeam { get; set; }

        /// <summary>
        /// Gets or sets team id.
        /// </summary>
        [JsonProperty("TeamId")]
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or sets unique identifier of Nomination.
        /// </summary>
        [JsonProperty("NominationId")]
        public string NominationId { get; set; }

        /// <summary>
        /// Gets or sets name of award.
        /// </summary>
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
        [JsonProperty("AwardLink")]
        public string AwardLink { get; set; }

        /// <summary>
        /// Gets or sets nominee name.
        /// </summary>
        [JsonProperty("NominatedToName")]
        public string NominatedToName { get; set; }

        /// <summary>
        /// Gets or sets AAD object id of nominee.
        /// </summary>
        [JsonProperty("NominatedToObjectId")]
        public string NominatedToObjectId { get; set; }

        /// <summary>
        /// Gets or sets User principal name of nominee.
        /// </summary>
        [JsonProperty("NominatedToPrincipalName")]
        public string NominatedToPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets User principal name of nominator.
        /// </summary>
        [JsonProperty("NominatedByName")]
        public string NominatedByName { get; set; }

        /// <summary>
        /// Gets or sets reward cycle identifier.
        /// </summary>
        [JsonProperty("RewardCycleId")]
        public string RewardCycleId { get; set; }

        /// <summary>
        /// Gets or sets note that was given to the nominee.
        /// </summary>
        [JsonProperty("ReasonForNomination")]
        public string ReasonForNomination { get; set; }

        /// <summary>
        /// Gets or sets start date of reward cycle.
        /// </summary>
        public DateTime RewardCycleStartDate { get; set; }

        /// <summary>
        /// Gets or sets end date of reward cycle.
        /// </summary>
        public DateTime RewardCycleEndDate { get; set; }

        /// <summary>
        /// Gets or sets commands from which task module is invoked.
        /// </summary>
        [JsonProperty("command")]
        public string Command { get; set; }
    }
}