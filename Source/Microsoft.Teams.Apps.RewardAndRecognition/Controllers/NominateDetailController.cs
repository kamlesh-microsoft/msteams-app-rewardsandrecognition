// <copyright file="NominateDetailController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;

    /// <summary>
    /// This endpoint is used to manage award nominations.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize]
    public class NominateDetailController : BaseRewardAndRecognitionController
    {
        private readonly ILogger<AwardsController> logger;
        private readonly INominateAwardStorageProvider storageProvider;
        private readonly IEndorseDetailStorageProvider endorseStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="NominateDetailController"/> class.
        /// </summary>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="storageProvider">Nominate award detail storage provider.</param>
        /// <param name="endorseStorageProvider">Endorse detail storage provider.</param>
        public NominateDetailController(ILogger<AwardsController> logger, INominateAwardStorageProvider storageProvider, IEndorseDetailStorageProvider endorseStorageProvider)
        {
            this.logger = logger;
            this.storageProvider = storageProvider;
            this.endorseStorageProvider = endorseStorageProvider;
        }

        /// <summary>
        /// Post call to save nominated award details in Azure Table storage.
        /// </summary>
        /// <param name="nominateDetails">Class contains details of on award nomination.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost("nomination")]
        public async Task<IActionResult> SaveNominateDetailsAsync([FromBody]NominateEntity nominateDetails)
        {
            try
            {
                if (nominateDetails == null)
                {
                    return this.BadRequest();
                }

                this.logger.LogInformation("Initiated call to on storage provider service.");
                var result = await this.storageProvider.StoreOrUpdateNominatedDetailsAsync(nominateDetails);
                this.logger.LogInformation("POST call for nominated award details in storage is successful.");
                return this.Ok(result);
            }
            catch (Exception ex)
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                this.logger.LogError(ex, "Error while saving nominated award details.");
                throw;
            }
        }

        /// <summary>
        /// This method is used to fetch nomination details for a given team Id and aadObjectId.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="aadObjectId">Azure active directory object Id.</param>
        /// <param name="cycleId">Active reward cycle id.</param>
        /// <returns>Nomination details.</returns>
        [HttpGet("nominationdetail")]
        public async Task<IActionResult> GetNominationDetailsAsync(string teamId, string aadObjectId, string cycleId)
        {
            try
            {
                var nominationDetails = await this.storageProvider.GetNominateDetailsAsync(teamId, aadObjectId, cycleId);
                return this.Ok(nominationDetails);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "failed to get nomination details" + aadObjectId);
                throw;
            }
        }

        /// <summary>
        /// This method returns all nominations for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="isAwardGranted">True for published awards, else false.</param>
        /// <param name="awardCycleId">Active award cycle.</param>
        /// <returns>Returns all nominations</returns>
        [HttpGet("allnominations")]
        public async Task<IActionResult> GetNominationDetailsAsync(string teamId, bool isAwardGranted, string awardCycleId)
        {
            try
            {
                var publishAwardDetails = new List<PublishResult>();
                var nominations = await this.storageProvider.GetNominationDetailsAsync(teamId, isAwardGranted, awardCycleId);
                var endorseDetails = await this.endorseStorageProvider.GetEndorseDetailAsync(teamId, awardCycleId, nominatedToPrincipalName: string.Empty);
                publishAwardDetails = nominations.Select(nomination => new PublishResult()
                {
                    AwardCycle = string.Empty,
                    AwardName = nomination.AwardName,
                    NominatedByName = nomination.NominatedByName,
                    AwardId = nomination.AwardId,
                    NominatedByObjectId = nomination.NominatedByObjectId,
                    NominationId = nomination.NominationId,
                    NominatedByPrincipalName = nomination.NominatedByPrincipalName,
                    NominatedToName = nomination.NominatedToName,
                    NominatedToObjectId = nomination.NominatedToObjectId,
                    NominatedToPrincipalName = nomination.NominatedToPrincipalName,
                    RewardCycleId = nomination.RewardCycleId,
                    ReasonForNomination = nomination.ReasonForNomination,
                    EndorseCount = endorseDetails.Where(agg => agg.EndorsedToPrincipalName.Equals(nomination.NominatedToPrincipalName, StringComparison.OrdinalIgnoreCase)
                    && nomination.AwardId.Equals(agg.EndorseForAwardId, StringComparison.OrdinalIgnoreCase)).Count(),
                }).ToList();

                return this.Ok(publishAwardDetails);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "failed to get nominations" + teamId);
                throw;
            }
        }

        /// <summary>
        /// This method updates all published nominations.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="nominationIds">Published nomination ids.</param>
        /// <returns>Returns all nominations</returns>
        [HttpGet("publishnominations")]
        public async Task<IActionResult> UpdateNominationDetailsAsync(string teamId, string nominationIds)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(nominationIds))
                {
                    return this.BadRequest();
                }

                var result = await this.storageProvider.PublishNominationDetailsAsync(teamId, nominationIds.Split(','));
                return this.Ok(result);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "failed to publish nominations" + teamId);
                throw;
            }
        }
    }
}