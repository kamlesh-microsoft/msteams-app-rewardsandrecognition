// <copyright file="NotificationController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RewardAndRecognition.Cards;
    using Microsoft.Teams.Apps.RewardAndRecognition.Helpers;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;

    /// <summary>
    /// This endpoint is used to send messages to team.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize]
    public class NotificationController : BaseRewardAndRecognitionController
    {
        /// <summary>
        /// default value for channel activity to send notifications.
        /// </summary>
        private const string Channel = "msteams";

        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly IBotFrameworkHttpAdapter adapter;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly IConfiguration configuration;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<NotificationController> logger;

        /// <summary>
        /// Provider for fetching information about team details from storage table.
        /// </summary>
        private readonly ITeamStorageProvider teamStorageProvider;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Initializes a new instance of the <see cref="NotificationController"/> class.
        /// </summary>
        /// <param name="adapter">bot adapter.</param>
        /// <param name="configuration">configuration.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="teamStorageProvider">Store or update teams details in Azure table storage.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        public NotificationController(IBotFrameworkHttpAdapter adapter, IConfiguration configuration, ILogger<NotificationController> logger, ITeamStorageProvider teamStorageProvider, IStringLocalizer<Strings> localizer)
        {
            this.adapter = adapter;
            this.configuration = configuration;
            this.logger = logger;
            this.teamStorageProvider = teamStorageProvider;
            this.localizer = localizer;
        }

        /// <summary>
        /// Get award winners details.
        /// </summary>
        /// <param name="details">Notification details.</param>
        /// <returns>Sends winner card to teams channel.</returns>
        [HttpPost("winnernotification")]
        public async Task<IActionResult> WinnerNominationAsync([FromBody]IList<AwardWinnerNotification> details)
        {
            try
            {
                if (details == null)
                {
                    return this.BadRequest();
                }

                var emails = string.Join(",", details.Select(row => row.NominatedToPrincipalName)).Split(",").Select(row => row.Trim()).Distinct();
                string teamId = details.First().TeamId;
                var claims = this.GetUserClaims();
                var teamDetails = await this.teamStorageProvider.GetTeamDetailAsync(teamId);
                string serviceUrl = teamDetails.ServiceUrl;
                string appId = this.configuration["MicrosoftAppId"];
                string appBaseUrl = this.configuration.GetValue<string>("Bot:AppBaseUri");
                string manifestId = this.configuration.GetValue<string>("Bot:ManifestId");
                MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);
                var conversationParameters = new ConversationParameters()
                {
                    ChannelData = new TeamsChannelData() { Team = new TeamInfo() { Id = teamId }, Channel = new ChannelInfo() { Id = teamId } },
                    Activity = (Activity)MessageFactory.Carousel(WinnerCarouselCard.GetAwardWinnerCard(appBaseUrl, details, this.localizer, manifestId)),
                    Bot = new ChannelAccount() { Id = appId },
                    IsGroup = true,
                    TenantId = this.configuration.GetValue<string>("Bot:TenantId"),
                };

                await ((BotFrameworkAdapter)this.adapter).CreateConversationAsync(
                    Channel,
                    serviceUrl,
                    new MicrosoftAppCredentials(this.configuration.GetValue<string>("MicrosoftAppId"), this.configuration.GetValue<string>("MicrosoftAppPassword")),
                    conversationParameters,
                    async (turnContext, cancellationToken) =>
                    {
                        Activity mentionActivity = await CardHelper.GetMentionActivityAsync(emails, claims.FromId, teamId, turnContext, this.localizer, this.logger, MentionActivityType.Winner, default);
                        await ((BotFrameworkAdapter)this.adapter).ContinueConversationAsync(
                            this.configuration.GetValue<string>("MicrosoftAppId"),
                            turnContext.Activity.GetConversationReference(),
                            async (continueConversationTurnContext, continueConversationCancellationToken) =>
                            {
                                mentionActivity.ApplyConversationReference(turnContext.Activity.GetConversationReference());
                                await continueConversationTurnContext.SendActivityAsync(mentionActivity, continueConversationCancellationToken);
                            }, cancellationToken);
                    }, default);

                // Let the caller know proactive messages have been sent
                return new ContentResult()
                {
                    StatusCode = (int)HttpStatusCode.OK,
                };
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "problem while sending the winner card");
                throw;
            }
        }
    }
}
