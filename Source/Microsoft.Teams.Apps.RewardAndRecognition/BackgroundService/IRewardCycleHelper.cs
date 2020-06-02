// <copyright file="IRewardCycleHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.BackgroundService
{
    using System.Threading.Tasks;

    /// <summary>
    /// Interface for reward cycle helper.
    /// </summary>
    public interface IRewardCycleHelper
    {
        /// <summary>
        /// This method is used to start reward cycle is the start date matches the current date and stops the reward cycle if the end date matches the current date
        /// </summary>
        /// <returns>A <see cref="Task"/>Representing the result of the asynchronous operation.</returns>
        Task<bool> CheckOrUpdateCycleStatusAsync();
    }
}
