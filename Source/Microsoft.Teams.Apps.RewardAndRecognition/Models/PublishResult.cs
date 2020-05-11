// <copyright file="PublishResult.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Class contains details of publish awards.
    /// </summary>
    public class PublishResult : NominateEntity
    {
        /// <summary>
        /// Gets or sets award cycle.
        /// </summary>
        [JsonProperty("AwardCycle")]
        public string AwardCycle { get; set; }

        /// <summary>
        /// Gets or sets endorse count.
        /// </summary>
        [JsonProperty("EndorseCount")]
        public int EndorseCount { get; set; }
    }
}
