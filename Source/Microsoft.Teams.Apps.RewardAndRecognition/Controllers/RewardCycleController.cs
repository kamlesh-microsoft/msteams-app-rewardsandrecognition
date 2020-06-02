// <copyright file="RewardCycleController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Controllers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;

    /// <summary>
    /// This endpoint is used to manage reward cycle.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize]
    public class RewardCycleController : BaseRewardAndRecognitionController
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<RewardCycleController> logger;

        /// <summary>
        /// Provider for fetching information about active award cycle details from storage table.
        /// </summary>
        private readonly IRewardCycleStorageProvider storageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="RewardCycleController"/> class.
        /// </summary>
        /// <param name="logger">Sends logs to the application insights service.</param>
        /// <param name="storageProvider">Reward cycle storage provider.</param>
        public RewardCycleController(ILogger<RewardCycleController> logger, IRewardCycleStorageProvider storageProvider)
        {
            this.logger = logger;
            this.storageProvider = storageProvider;
        }

        /// <summary>
        /// This method returns reward cycle for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="isActiveCycle">Reward cycle state.</param>
        /// <returns>Reward cycle details.</returns>
        [HttpGet("rewardcycledetails")]
        public async Task<IActionResult> GetRewardCycleAsync(string teamId, bool isActiveCycle = true)
        {
            RewardCycleEntity rewardCycle;
            try
            {
                if (isActiveCycle)
                {
                    rewardCycle = await this.storageProvider.GetActiveRewardCycleAsync(teamId);
                }
                else
                {
                    rewardCycle = await this.storageProvider.GetPublishedRewardCycleAsync(teamId);
                }

                return this.Ok(rewardCycle);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "failed to get awards" + teamId);
                throw;
            }
        }

        /// <summary>
        /// Post call to store reward cycle details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="rewardCycleEntity">Holds reward cycle detail entity data.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost("rewardcycle")]
        public async Task<IActionResult> PostAsync([FromBody] RewardCycleEntity rewardCycleEntity)
        {
            try
            {
                this.logger.LogInformation("set reward cycle");
                if (rewardCycleEntity?.CycleId == null)
                {
                    rewardCycleEntity.CycleId = Guid.NewGuid().ToString();
                    rewardCycleEntity.CreatedOn = DateTime.UtcNow;
                }

                return this.Ok(await this.storageProvider.StoreOrUpdateRewardCycleAsync(rewardCycleEntity));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to award service.");
                throw;
            }
        }
    }
}