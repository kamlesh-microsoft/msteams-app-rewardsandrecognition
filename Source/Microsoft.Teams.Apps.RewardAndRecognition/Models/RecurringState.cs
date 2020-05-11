// <copyright file="RecurringState.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    /// <summary>
    /// Enum to specify the reward recurring state.
    /// </summary>
    public enum RecurringState
    {
        /// <summary>
        /// Non recursive award cycle.
        /// </summary>
        NonRecursive = 0,

        /// <summary>
        /// Recursive award cycle.
        /// </summary>
        Recursive = 1,
    }
}
