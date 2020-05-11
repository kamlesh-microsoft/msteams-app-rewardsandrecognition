// <copyright file="PublishAwardEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    using System.Collections.Generic;

    /// <summary>
    /// Class contains details of publish award details.
    /// </summary>
    public class PublishAwardEntity
    {
        /// <summary>
        /// Gets or sets award cycle.
        /// </summary>
        public string AwardCycle { get; set; }

        /// <summary>
        /// Gets or sets name of award.
        /// </summary>
        public string AwardName { get; set; }

        /// <summary>
        /// Gets or sets awards.
        /// </summary>
        public IEnumerable<PublishResult> Nominations { get; set; }
    }
}
