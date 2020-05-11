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
        /// Channel conversation type to send notification.
        /// </summary>
        private const string ChannelConversationType = "channel";
        private readonly IBotFrameworkHttpAdapter adapter;
        private readonly IConfiguration configuration;
        private readonly ILogger<NotificationController> logger;
        private readonly ITeamStorageProvider teamStorageProvider;
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
                MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);
                var conversationReference = new ConversationReference()
                {
                    ChannelId = Channel,
                    Bot = new ChannelAccount() { Id = appId },
                    ServiceUrl = serviceUrl,
                    Conversation = new ConversationAccount()
                    {
                        ConversationType = ChannelConversationType,
                        IsGroup = true,
                        Id = teamId,
                        TenantId = this.configuration.GetValue<string>("Bot:TenantId"),
                    },
                };

                await ((BotFrameworkAdapter)this.adapter).ContinueConversationAsync(
                               appId,
                               conversationReference,
                               async (conversationTurnContext, conversationCancellationToken) =>
                               {
                                   System.Diagnostics.Debug.WriteLine(conversationTurnContext.Activity.ServiceUrl);
                                   var result = await conversationTurnContext.SendActivityAsync(MessageFactory.Carousel(WinnerCarouselCard.GetAwardWinnerCard(appBaseUrl, details, this.localizer)));
                                   Activity mentionActivity = await CardHelper.GetMentionActivityAsync(emails, claims.FromId, teamId, conversationTurnContext, this.localizer, this.logger, default);
                                   conversationReference.Conversation.Id = $"{teamId};messageid={result.Id}";
                                   await conversationTurnContext.SendActivityAsync(mentionActivity);
                               },
                               default);

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
