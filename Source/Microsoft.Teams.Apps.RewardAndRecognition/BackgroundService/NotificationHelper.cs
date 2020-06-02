// <copyright file="NotificationHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.BackgroundService
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition.Cards;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Helper class to send nomination reminder notification.
    /// </summary>
    public class NotificationHelper : INotificationHelper
    {
        /// <summary>
        /// Default value for channel activity to send notifications.
        /// </summary>
        private const string Channel = "msteams";

        /// <summary>
        /// Channel conversation type to send notification.
        /// </summary>
        private const string ChannelConversationType = "channel";

        /// <summary>
        /// Nominate reminder notification days back.
        /// </summary>
        private const int LookBackDays = 3;

        /// <summary>
        /// Retry policy with jitter.
        /// </summary>
        /// <remarks>
        /// Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </remarks>
        private static AsyncRetryPolicy retryPolicy = Policy.Handle<Exception>()
            .WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(1000), 2));

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly IBotFrameworkHttpAdapter adapter;

        /// <summary>
        /// Microsoft application credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly IOptions<RewardAndRecognitionActivityHandlerOptions> options;

        /// <summary>
        /// Helper for storing reward cycle details to azure table storage.
        /// </summary>
        private readonly IRewardCycleStorageProvider rewardCycleStorageProvider;

        /// <summary>
        /// Helper for fetching reward details from azure table storage.
        /// </summary>
        private readonly IAwardsStorageProvider awardsStorageProvider;

        /// <summary>
        /// Helper for fetching teams details from azure table storage.
        /// </summary>
        private readonly ITeamStorageProvider teamStorageProvider;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<RewardCycleHelper> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="NotificationHelper"/> class.
        /// </summary>
        /// <param name="rewardCycleStorageProvider">Reward cycle storage provider.</param>
        /// <param name="awardsStorageProvider">Award storage provider.</param>
        /// <param name="teamStorageProvider">Teams storage provider.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        /// <param name="adapter">Bot adapter.</param>
        /// <param name="microsoftAppCredentials">MicrosoftAppCredentials of bot.</param>
        public NotificationHelper(
            IRewardCycleStorageProvider rewardCycleStorageProvider,
            ITeamStorageProvider teamStorageProvider,
            IAwardsStorageProvider awardsStorageProvider,
            ILogger<RewardCycleHelper> logger,
            IStringLocalizer<Strings> localizer,
            IOptions<RewardAndRecognitionActivityHandlerOptions> options,
            IBotFrameworkHttpAdapter adapter,
            MicrosoftAppCredentials microsoftAppCredentials)
        {
            this.rewardCycleStorageProvider = rewardCycleStorageProvider;
            this.logger = logger;
            this.localizer = localizer;
            this.options = options ?? throw new ArgumentNullException(nameof(options));
            this.adapter = adapter;
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.awardsStorageProvider = awardsStorageProvider;
            this.teamStorageProvider = teamStorageProvider;
        }

        /// <summary>
        /// This method is used to send nomination reminder notification.
        /// </summary>
        /// <returns>Returns true if nomination reminder sent successfully else false.</returns>
        public async Task<bool> SendNominationReminderNotificationAsync()
        {
            var activeRewardCycle = await this.rewardCycleStorageProvider.GetActiveAwardCycleForAllTeamsAsync();
            foreach (var currentCyle in activeRewardCycle)
            {
                if (currentCyle.RewardCycleEndDate.ToUniversalTime().Day == DateTime.UtcNow.AddDays(-LookBackDays).Day)
                {
                    // Send nomination reminder notification
                    await this.SendCardToTeamAsync(currentCyle);
                }
            }

            return true;
        }

        /// <summary>
        /// Send the given attachment to the specified team.
        /// </summary>
        /// <param name="rewardCycleEntity">Reward cycle model object.</param>
        /// <returns>A task that sends notification card in channel.</returns>
        private async Task SendCardToTeamAsync(RewardCycleEntity rewardCycleEntity)
        {
            try
            {
                var awardsList = await this.awardsStorageProvider.GetAwardsAsync(rewardCycleEntity.TeamId);
                var valuesfromTaskModule = new TaskModuleResponseDetails()
                {
                    RewardCycleStartDate = rewardCycleEntity.RewardCycleStartDate,
                    RewardCycleEndDate = rewardCycleEntity.RewardCycleEndDate,
                    RewardCycleId = rewardCycleEntity.CycleId,
                };

                var teamDetails = await this.teamStorageProvider.GetTeamDetailAsync(rewardCycleEntity.TeamId);
                string serviceUrl = teamDetails.ServiceUrl;

                MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);
                string teamsChannelId = rewardCycleEntity.TeamId;

                var conversationReference = new ConversationReference()
                {
                    ChannelId = Channel,
                    Bot = new ChannelAccount() { Id = this.microsoftAppCredentials.MicrosoftAppId },
                    ServiceUrl = serviceUrl,
                    Conversation = new ConversationAccount() { ConversationType = ChannelConversationType, IsGroup = true, Id = teamsChannelId, TenantId = teamsChannelId },
                };

                this.logger.LogInformation($"sending notification to channelId- {teamsChannelId}");

                await retryPolicy.ExecuteAsync(async () =>
                {
                    try
                    {
                        var conversationParameters = new ConversationParameters()
                        {
                            ChannelData = new TeamsChannelData() { Team = new TeamInfo() { Id = rewardCycleEntity.TeamId }, Channel = new ChannelInfo() { Id = rewardCycleEntity.TeamId } },
                            Activity = (Activity)MessageFactory.Carousel(NominateCarouselCard.GetAwardsCard(this.options.Value.AppBaseUri, awardsList, this.localizer, valuesfromTaskModule)),
                            Bot = new ChannelAccount() { Id = this.microsoftAppCredentials.MicrosoftAppId },
                            IsGroup = true,
                            TenantId = this.options.Value.TenantId,
                        };

                        await ((BotFrameworkAdapter)this.adapter).CreateConversationAsync(
                            Channel,
                            serviceUrl,
                            this.microsoftAppCredentials,
                            conversationParameters,
                            async (conversationTurnContext, conversationCancellationToken) =>
                            {
                                Activity mentionActivity = MessageFactory.Text(this.localizer.GetString("NominationReminderNotificationText"));
                                await ((BotFrameworkAdapter)this.adapter).ContinueConversationAsync(
                                    this.microsoftAppCredentials.MicrosoftAppId,
                                    conversationTurnContext.Activity.GetConversationReference(),
                                    async (continueConversationTurnContext, continueConversationCancellationToken) =>
                                    {
                                        mentionActivity.ApplyConversationReference(conversationTurnContext.Activity.GetConversationReference());
                                        await continueConversationTurnContext.SendActivityAsync(mentionActivity, continueConversationCancellationToken);
                                    }, conversationCancellationToken);
                            },
                            default);
                    }
                    catch (Exception ex)
                    {
                        this.logger.LogError(ex, "Error while performing retry logic to send notification to channel.");
                        throw;
                    }
                });
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while sending notification to channel from background service.");
            }
        }
    }
}
