// <copyright file="AwardsController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;

    /// <summary>
    /// This endpoint is used to manage awards.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize]
    public class AwardsController : BaseRewardAndRecognitionController
    {
        private readonly ILogger<AwardsController> logger;
        private readonly IAwardsStorageProvider storageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="AwardsController"/> class.
        /// </summary>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="storageProvider">Awards storage provider.</param>
        public AwardsController(ILogger<AwardsController> logger, IAwardsStorageProvider storageProvider)
        {
            this.logger = logger;
            this.storageProvider = storageProvider;
        }

        /// <summary>
        /// This method returns all awards for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns>Awards</returns>
        [HttpGet("allawards")]
        public async Task<IActionResult> GetAwardsAsync(string teamId)
        {
            try
            {
                var awards = await this.storageProvider.GetAwardsAsync(teamId);
                return this.Ok(awards);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "failed to get awards" + teamId);
                throw;
            }
        }

        /// <summary>
        /// This method returns award details for a given team Id and awardId.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="awardId">Award Id.</param>
        /// <returns>Award Details</returns>
        [HttpGet("awarddetails")]
        public async Task<IActionResult> GetAwardDetailsAsync(string teamId, string awardId)
        {
            try
            {
                var award = await this.storageProvider.GetAwardDetailsAsync(teamId, awardId);
                return this.Ok(award);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "failed to get award details" + awardId);
                throw;
            }
        }

        /// <summary>
        /// Post call to store award details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="awardEntity">Holds award detail entity data.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost("award")]
        public async Task<IActionResult> PostAsync([FromBody] AwardEntity awardEntity)
        {
            try
            {
                if (string.IsNullOrEmpty(awardEntity?.AwardName))
                {
                    this.logger.LogError("Error while creating award details data in Microsoft Azure Table storage.");
                    return this.BadRequest();
                }

                var claims = this.GetUserClaims();
                this.logger.LogInformation("Adding award");
                if (awardEntity.AwardId == null)
                {
                    awardEntity.AwardId = Guid.NewGuid().ToString();
                    awardEntity.CreatedOn = DateTime.UtcNow;
                    awardEntity.CreatedBy = claims.FromId;
                }
                else
                {
                    awardEntity.ModifiedBy = claims.FromId;
                }

                return this.Ok(await this.storageProvider.StoreOrUpdateAwardAsync(awardEntity));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to award service.");
                throw;
            }
        }

        /// <summary>
        /// Delete call to delete award details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="awardIds">User selected response Ids.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpDelete("awards")]
        public async Task<IActionResult> DeleteAsync(string awardIds)
        {
            try
            {
                if (awardIds == null)
                {
                    this.logger.LogError("Error while deleting award details data in Microsoft Azure Table storage.");
                    return this.BadRequest();
                }

                IList<string> awards = awardIds.Split(",");
                return this.Ok(await this.storageProvider.DeleteAwardsAsync(awards));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while deleting awards");
                throw;
            }
        }
    }
}