// <copyright file="OccurrenceType.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    /// <summary>
    /// Enum to specify the type of occurrence state.
    /// </summary>
    public enum OccurrenceType
    {
        /// <summary>
        /// NoEndDate award cycle.
        /// </summary>
        NoEndDate = 0,

        /// <summary>
        /// EndDate award cycle.
        /// </summary>
        EndDate = 1,

        /// <summary>
        /// EndDate award cycle.
        /// </summary>
        Occurrence = 2,
    }
}
