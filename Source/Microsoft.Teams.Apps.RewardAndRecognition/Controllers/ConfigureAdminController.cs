// <copyright file="ConfigureAdminController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;

    /// <summary>
    /// Controller to handle configure admin API operations.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize]
    public class ConfigureAdminController : BaseRewardAndRecognitionController
    {
        /// <summary>
        /// Microsoft Application ID.
        /// </summary>
        private readonly string appId;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Rewards and recognition Bot adapter to get context.
        /// </summary>
        private readonly BotFrameworkAdapter botAdapter;

        /// <summary>
        /// Provider to fetch admin details from Azure Table Storage.
        /// </summary>
        private readonly IConfigureAdminStorageProvider storageProvider;

        /// <summary>
        /// Provider to fetch team details from Azure Table Storage.
        /// </summary>
        private readonly ITeamStorageProvider teamStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigureAdminController"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the application insights service.</param>
        /// <param name="botAdapter">Reward and Recognition bot adapter.</param>
        /// <param name="microsoftAppCredentials">Microsoft Application credentials for Bot/ME.</param>
        /// <param name="storageProvider">Provider to store admin details in Azure Table Storage.</param>
        /// <param name="teamStorageProvider">Store or update teams details in Azure table storage.</param>
        public ConfigureAdminController(
            ILogger<ConfigureAdminController> logger,
            BotFrameworkAdapter botAdapter,
            MicrosoftAppCredentials microsoftAppCredentials,
            IConfigureAdminStorageProvider storageProvider,
            ITeamStorageProvider teamStorageProvider)
            : base()
        {
            this.logger = logger;
            this.botAdapter = botAdapter;
            this.teamStorageProvider = teamStorageProvider;
            this.appId = microsoftAppCredentials != null ? microsoftAppCredentials.MicrosoftAppId : throw new ArgumentNullException(nameof(microsoftAppCredentials));
            this.storageProvider = storageProvider;
        }

        /// <summary>
        /// Get list of members present in a team.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns>List of members in team.</returns>
        [HttpGet("teammembers")]
        public async Task<IActionResult> GetTeamMembersAsync(string teamId)
        {
            try
            {
                if (teamId == null)
                {
                    this.logger.LogInformation("Mobile test : BadRequest");
                    return this.BadRequest(new { message = "Team ID cannot be empty." });
                }

                this.logger.LogInformation("Mobile test : GetTeamMembersAsync: " + teamId);

                IEnumerable<TeamsChannelAccount> teamsChannelAccounts = new List<TeamsChannelAccount>();

                var teamDetails = await this.teamStorageProvider.GetTeamDetailAsync(teamId);
                this.logger.LogInformation("Mobile test : GetTeamDetailAsync:  success" + teamId);
                string serviceUrl = teamDetails.ServiceUrl;

                this.logger.LogInformation("Mobile test : GetTeamDetailAsync:  serviceUrl" + serviceUrl);

                var conversationReference = new ConversationReference
                {
                    ChannelId = teamId,
                    ServiceUrl = serviceUrl,
                };
                await this.botAdapter.ContinueConversationAsync(
                    this.appId,
                    conversationReference,
                    async (context, token) =>
                    {
                        teamsChannelAccounts = await TeamsInfo.GetTeamMembersAsync(context, teamId, CancellationToken.None);
                    }, default);

                this.logger.LogInformation("GET call for fetching team members from team roster is successful.");
                return this.Ok(teamsChannelAccounts.Select(member => new { content = member.Email, header = member.Name, aadobjectid = member.AadObjectId }));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error occurred while getting team member list.");
                throw;
            }
        }

        /// <summary>
        /// Post call to save admin details in Azure Table storage.
        /// </summary>
        /// <param name="adminDetails">Class contains details of admin.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost("admindetail")]
        public async Task<IActionResult> SaveAdminDetailsAsync([FromBody]AdminEntity adminDetails)
        {
            try
            {
                if (adminDetails == null)
                {
                    return this.BadRequest();
                }

                this.logger.LogInformation("Initiated call to on storage provider service.");
                var result = await this.storageProvider.UpsertAdminDetailAsync(adminDetails);
                this.logger.LogInformation("POST call for saving admin details in storage is successful.");
                return this.Ok(result);
            }
            #pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                this.logger.LogError(ex, "Error while saving on admin details.");
                throw;
            }
        }

        /// <summary>
        /// This method returns all admin details for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns>Admin entity</returns>
        [HttpGet("alladmindetails")]
        public async Task<IActionResult> GetAdminDetailsAsync(string teamId)
        {
            try
            {
                var adminDetails = await this.storageProvider.GetAdminDetailAsync(teamId);
                return this.Ok(adminDetails);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "failed to get admin details" + teamId);
                throw;
            }
        }
    }
}
