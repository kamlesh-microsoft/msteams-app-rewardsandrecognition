// <copyright file="RewardCycleHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.BackgroundService
{
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;

    /// <summary>
    /// Helper class to start or stop reward cycle.
    /// </summary>
    public class RewardCycleHelper : IRewardCycleHelper
    {
        /// <summary>
        /// Helper for storing reward cycle details to azure table storage.
        /// </summary>
        private readonly IRewardCycleStorageProvider rewardCycleStorageProvider;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<RewardCycleHelper> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="RewardCycleHelper"/> class.
        /// </summary>
        /// <param name="rewardCycleStorageProvider">Reward cycle storage provider.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public RewardCycleHelper(IRewardCycleStorageProvider rewardCycleStorageProvider, ILogger<RewardCycleHelper> logger)
        {
            this.rewardCycleStorageProvider = rewardCycleStorageProvider;
            this.logger = logger;
        }

        /// <summary>
        /// This method is used to start reward cycle is the start date matches the current date and stops the reward cycle if the end date matches the current date.
        /// </summary>
        /// <returns>Task.</returns>
        public async Task<bool> CheckOrUpdateCycleStatusAsync()
        {
            this.logger.LogInformation("Check and update reward cycle");
            return await this.rewardCycleStorageProvider.UpdateCycleStatusAsync();
        }
    }
}
