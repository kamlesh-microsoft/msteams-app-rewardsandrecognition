// <copyright file="RewardAndRecognitionActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition.Cards;
    using Microsoft.Teams.Apps.RewardAndRecognition.Helpers;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// The RewardAndRecognitionActivityHandler is responsible for reacting to incoming events from Teams sent from BotFramework.
    /// </summary>
    public sealed class RewardAndRecognitionActivityHandler : TeamsActivityHandler
    {
        private readonly IOptions<RewardAndRecognitionActivityHandlerOptions> options;

        private readonly string instrumentationKey;

        private readonly ILogger<RewardAndRecognitionActivityHandler> logger;

        private readonly IStringLocalizer<Strings> localizer;

        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Represents the Application base Uri.
        /// </summary>
        private readonly string appBaseUrl;

        /// <summary>
        /// Provider for fetching information about admin details from storage table.
        /// </summary>
        private readonly IConfigureAdminStorageProvider configureAdminStorageProvider;

        /// <summary>
        /// Provider for fetching information about team details from storage table.
        /// </summary>
        private readonly ITeamStorageProvider teamStorageProvider;

        /// <summary>
        /// Provider for fetching information about awards from storage table.
        /// </summary>
        private readonly IAwardsStorageProvider awardsStorageProvider;

        /// <summary>
        /// Provider for fetching information about endorsement details from storage table.
        /// </summary>
        private readonly IEndorseDetailStorageProvider endorseDetailStorageProvider;

        /// <summary>
        /// Provider for fetching information about active award cycle details from storage table.
        /// </summary>
        private readonly IRewardCycleStorageProvider rewardCycleStorageProvider;

        /// <summary>
        /// Provider to search nomination details in Azure search service.
        /// </summary>
        private readonly INominateDetailSearchService nominateDetailSearchService;

        /// <summary>
        /// Initializes a new instance of the <see cref="RewardAndRecognitionActivityHandler"/> class.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="telemetryClient">The application insights telemetry client. </param>
        /// <param name="options">The options.</param>
        /// <param name="telemetryOptions">Telemetry instrumentation key</param>
        /// <param name="configureAdminStorageProvider">Provider for fetching information about admin details from storage table.</param>
        /// <param name="teamStorageProvider">Provider for fetching information about team details from storage table.</param>
        /// <param name="awardsStorageProvider">Provider for fetching information about awards from storage table.</param>
        /// <param name="endorseDetailStorageProvider">Provider for fetching information about endorsement details from storage table.</param>
        /// <param name="rewardCycleStorageProvider">Provider for fetching information about active award cycle details from storage table.</param>
        /// <param name="nominateDetailSearchService">Provider to search nomination details in Azure search service.</param>
        public RewardAndRecognitionActivityHandler(
            ILogger<RewardAndRecognitionActivityHandler> logger,
            IStringLocalizer<Strings> localizer,
            TelemetryClient telemetryClient,
            IOptions<RewardAndRecognitionActivityHandlerOptions> options,
            IOptions<TelemetryOptions> telemetryOptions,
            IConfigureAdminStorageProvider configureAdminStorageProvider,
            ITeamStorageProvider teamStorageProvider,
            IAwardsStorageProvider awardsStorageProvider,
            IEndorseDetailStorageProvider endorseDetailStorageProvider,
            IRewardCycleStorageProvider rewardCycleStorageProvider,
            INominateDetailSearchService nominateDetailSearchService)
        {
            this.logger = logger;
            this.localizer = localizer;
            this.telemetryClient = telemetryClient;
            this.options = options ?? throw new ArgumentNullException(nameof(options));
            this.instrumentationKey = telemetryOptions?.Value.InstrumentationKey;
            this.appBaseUrl = this.options.Value.AppBaseUri;
            this.configureAdminStorageProvider = configureAdminStorageProvider;
            this.teamStorageProvider = teamStorageProvider;
            this.awardsStorageProvider = awardsStorageProvider;
            this.endorseDetailStorageProvider = endorseDetailStorageProvider;
            this.rewardCycleStorageProvider = rewardCycleStorageProvider;
            this.nominateDetailSearchService = nominateDetailSearchService;
        }

        /// <summary>
        /// Handle when a message is addressed to the bot.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// For more information on bot messaging in Teams, see the documentation
        /// https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/conversation-basics?tabs=dotnet#receive-a-message .
        /// </remarks>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnMessageActivityAsync), turnContext);
                await this.SendTypingIndicatorAsync(turnContext);
                await turnContext.SendActivityAsync(MessageFactory.Text(this.localizer.GetString("UnsupportedBotCommand")), cancellationToken);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error processing message: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Overriding to send welcome card once Bot/ME is installed in team.
        /// </summary>
        /// <param name="membersAdded">A list of all the members added to the conversation, as described by the conversation update activity.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Welcome card  when bot is added first time by user.</returns>
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

            var activity = turnContext.Activity;
            this.logger.LogInformation($"conversationType: {activity.Conversation?.ConversationType}, membersAdded: {membersAdded?.Count}");

            if (membersAdded.Where(member => member.Id == activity.Recipient.Id).FirstOrDefault() != null)
            {
                this.logger.LogInformation($"Bot added {activity.Conversation.Id}");
                await turnContext.SendActivityAsync(MessageFactory.Attachment(WelcomeCard.GetCard(this.appBaseUrl, this.localizer)), cancellationToken);
            }

            var teamsDetails = turnContext.Activity.TeamsGetTeamInfo();
            TeamEntity teamEntity = new TeamEntity
            {
                TeamId = teamsDetails.Id,
                BotInstalledOn = DateTime.UtcNow,
                ServiceUrl = turnContext.Activity.ServiceUrl,
            };
            bool operationStatus = await this.teamStorageProvider.StoreOrUpdateTeamDetailAsync(teamEntity);
            if (!operationStatus)
            {
                this.logger.LogInformation($"Unable to store bot installed detail in table storage.");
            }
        }

        /// <summary>
        /// Overriding to send card when R&R admin member is removed from team.
        /// </summary>
        /// <param name="membersRemoved">A member removed from team, as described by the conversation update activity.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Notification card  when bot or member is removed from team.</returns>
        protected override async Task OnMembersRemovedAsync(IList<ChannelAccount> membersRemoved, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

            var activity = turnContext.Activity;
            this.logger.LogInformation($"conversationType: {activity.Conversation?.ConversationType}, membersRemoved: {membersRemoved?.Count}");
            var teamsDetails = turnContext.Activity.TeamsGetTeamInfo();
            var admin = await this.configureAdminStorageProvider.GetAdminDetailAsync(teamsDetails.Id);

            if (membersRemoved.Where(member => member.AadObjectId == admin.AdminObjectId).FirstOrDefault() != null)
            {
                this.logger.LogInformation($"Member removed {activity.Conversation.Id}");
                await turnContext.SendActivityAsync(MessageFactory.Attachment(WelcomeCard.GetCard(this.appBaseUrl, this.localizer)), cancellationToken);
            }
            else if (membersRemoved.Where(member => member.Id == activity.Recipient.Id).FirstOrDefault() != null)
            {
                this.logger.LogInformation($"Bot removed {activity.Conversation.Id}");
                var teamEntity = await this.teamStorageProvider.GetTeamDetailAsync(teamsDetails.Id);
                bool operationStatus = await this.teamStorageProvider.DeleteTeamDetailAsync(teamEntity);
                if (!operationStatus)
                {
                    this.logger.LogInformation($"Unable to remove team details from table storage.");
                }
            }
        }

        /// <summary>
        /// Handle message extension action fetch task received by the bot.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="action">Messaging extension action value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Response of messaging extension action.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionfetchtaskasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionFetchTaskAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

                if (!await this.CheckTeamsValidaionAsync(turnContext))
                {
                    return CardHelper.GetTaskModuleInvalidTeamCard(this.localizer);
                }

                this.RecordEvent(nameof(this.OnTeamsMessagingExtensionFetchTaskAsync), turnContext);

                var activity = turnContext.Activity;
                var teamsDetails = turnContext.Activity.TeamsGetTeamInfo();
                var cycleStatus = await this.rewardCycleStorageProvider.GetActiveRewardCycleAsync(teamsDetails.Id);
                bool isCycleRunning = !(cycleStatus == null || cycleStatus.RewardCycleState == (int)RewardCycleState.InActive);

                return CardHelper.GetTaskModuleBasedOnCommand(this.appBaseUrl, this.instrumentationKey, this.localizer, teamsDetails.Id, isCycleRunning);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while fetching task module received by the bot which is invoked through ME.");
                throw;
            }
        }

        /// <summary>
        /// Invoked when the user submits a response/suggests a response/updates a response.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="action">Messaging extension action commands.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionsubmitactionasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionAction action,
            CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnTeamsMessagingExtensionSubmitActionAsync), turnContext);
                action = action ?? throw new ArgumentNullException(nameof(action));
                var valuesfromTaskModule = JsonConvert.DeserializeObject<TaskModuleResponseDetails>(action.Data.ToString());
                if (valuesfromTaskModule.Command.ToUpperInvariant() == Constants.SaveNominatedDetailsAction)
                {
                    var mentionActivity = await CardHelper.GetMentionActivityAsync(valuesfromTaskModule.NominatedToPrincipalName.Split(",").Select(row => row.Trim()).ToList(), turnContext.Activity.From.AadObjectId, valuesfromTaskModule.TeamId, turnContext, this.localizer, this.logger, cancellationToken);
                    var notificationCard = await turnContext.SendActivityAsync(MessageFactory.Attachment(EndorseCard.GetEndorseCard(this.appBaseUrl, this.localizer, valuesfromTaskModule)));
                    turnContext.Activity.Conversation.Id = $"{valuesfromTaskModule.TeamId};messageid={notificationCard.Id}";
                    await turnContext.SendActivityAsync(mentionActivity);
                    this.logger.LogInformation("Nominated an award");
                    return null;
                }

                return null;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error at OnTeamsMessagingExtensionSubmitActionAsync(): {ex.Message}", SeverityLevel.Error);
                await turnContext.SendActivityAsync(this.localizer.GetString("ErrorMessage"));
                throw;
            }
        }

        /// <summary>
        /// When OnTurn method receives a fetch invoke activity on bot turn, it calls this method..
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents a task module response.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            var activity = (Activity)turnContext.Activity;
            this.RecordEvent(nameof(this.OnTeamsTaskModuleFetchAsync), turnContext);
            try
            {
                var teamsDetails = turnContext.Activity.TeamsGetTeamInfo();
                var valuesforTaskModule = JsonConvert.DeserializeObject<AdaptiveCardAction>(((JObject)activity.Value).GetValue("data", StringComparison.OrdinalIgnoreCase)?.ToString());
                var cycleStatus = await this.rewardCycleStorageProvider.GetActiveRewardCycleAsync(teamsDetails.Id);
                if (cycleStatus != null && cycleStatus.CycleId != valuesforTaskModule.RewardCycleId && valuesforTaskModule.Command != Constants.ConfigureAdminAction)
                {
                    return CardHelper.GetTaskModuleResponse(applicationBasePath: this.appBaseUrl, instrumentationKey: this.instrumentationKey, localizer: this.localizer, command: valuesforTaskModule.Command, teamId: teamsDetails.Id, isCycleClosed: true);
                }
                else if ((cycleStatus == null || cycleStatus.RewardCycleState == (int)RewardCycleState.InActive) && valuesforTaskModule.Command != Constants.ConfigureAdminAction)
                {
                    return CardHelper.GetTaskModuleResponse(applicationBasePath: this.appBaseUrl, instrumentationKey: this.instrumentationKey, localizer: this.localizer, command: valuesforTaskModule.Command, teamId: teamsDetails.Id, isCycleRunning: false);
                }

                switch (valuesforTaskModule.Command)
                {
                    case Constants.ConfigureAdminAction:
                        bool isActivityIdPresent = !string.IsNullOrEmpty(turnContext.Activity.Conversation.Id.Split(';')[1].Split("=")[1]);
                        this.logger.LogInformation("Fetch task module to show configure R&R admin card.");
                        return CardHelper.GetTaskModuleResponse(applicationBasePath: this.appBaseUrl, instrumentationKey: this.instrumentationKey, localizer: this.localizer, command: valuesforTaskModule.Command, teamId: teamsDetails.Id, isActivityIdPresent: isActivityIdPresent);

                    case Constants.EndorseAction:
                        bool isEndorsementSuccess = await this.CheckEndorseStatusAsync(turnContext, valuesforTaskModule);
                        this.logger.LogInformation("Fetch and show task module to endorse an award nomination.");
                        return CardHelper.GetEndorseTaskModuleResponse(applicationBasePath: this.appBaseUrl, this.localizer, valuesforTaskModule.NominatedToName, valuesforTaskModule.AwardName, cycleStatus.RewardCycleEndDate, isEndorsementSuccess);

                    case Constants.NominateAction:
                        this.logger.LogInformation("Fetch and show task module to configure new nominate award card.");
                        return CardHelper.GetTaskModuleResponse(applicationBasePath: this.appBaseUrl, instrumentationKey: this.instrumentationKey, localizer: this.localizer, command: valuesforTaskModule.Command, teamId: teamsDetails.Id, awardId: valuesforTaskModule.AwardId);

                    default:
                        this.logger.LogInformation($"Invalid command for task module fetch activity.Command is : {valuesforTaskModule.Command} ");
                        await turnContext.SendActivityAsync(this.localizer.GetString("ErrorMessage"));
                        return null;
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error in fetching task module.");
                throw;
            }
        }

        /// <summary>
        /// When OnTurn method receives a submit invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents a task module response.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

                var activity = (Activity)turnContext.Activity;
                this.RecordEvent(nameof(this.OnTeamsTaskModuleFetchAsync), turnContext);
                IMessageActivity notificationCard;
                Activity mentionActivity;
                var valuesfromTaskModule = JsonConvert.DeserializeObject<TaskModuleResponseDetails>(((JObject)activity.Value).GetValue("data", StringComparison.OrdinalIgnoreCase)?.ToString());
                switch (valuesfromTaskModule.Command.ToUpperInvariant())
                {
                    case Constants.SaveAdminDetailsAction:
                        mentionActivity = await CardHelper.GetMentionActivityAsync(valuesfromTaskModule.AdminPrincipalName.Split(",").ToList(), turnContext.Activity.From.AadObjectId, valuesfromTaskModule.TeamId, turnContext, this.localizer, this.logger, cancellationToken);
                        var cardDetail = await turnContext.SendActivityAsync(MessageFactory.Attachment(AdminCard.GetAdminCard(this.localizer, valuesfromTaskModule, this.options.Value.ManifestId)));
                        turnContext.Activity.Conversation.Id = $"{turnContext.Activity.Conversation.Id};messageid={cardDetail.Id}";
                        await turnContext.SendActivityAsync(mentionActivity);
                        this.logger.LogInformation("R&R admin has been configured");
                        break;

                    case Constants.CancelCommand:
                        break;

                    case Constants.UpdateAdminDetailCommand:
                        mentionActivity = await CardHelper.GetMentionActivityAsync(valuesfromTaskModule.AdminPrincipalName.Split(",").ToList(), turnContext.Activity.From.AadObjectId, valuesfromTaskModule.TeamId, turnContext, this.localizer, this.logger, cancellationToken);
                        notificationCard = MessageFactory.Attachment(AdminCard.GetAdminCard(this.localizer, valuesfromTaskModule, this.options.Value.ManifestId));
                        notificationCard.Id = turnContext.Activity.Conversation.Id.Split(';')[1].Split("=")[1];
                        notificationCard.Conversation = turnContext.Activity.Conversation;
                        await turnContext.UpdateActivityAsync(notificationCard);
                        await turnContext.SendActivityAsync(mentionActivity);
                        this.logger.LogInformation("Card is updated.");
                        break;

                    case Constants.NominateAction:
                        var awardsList = await this.awardsStorageProvider.GetAwardsAsync(valuesfromTaskModule.TeamId);
                        await turnContext.SendActivityAsync(MessageFactory.Carousel(NominateCarouselCard.GetAwardsCard(this.appBaseUrl, awardsList, this.localizer, valuesfromTaskModule)));
                        break;

                    case Constants.SaveNominatedDetailsAction:
                        turnContext.Activity.Conversation.Id = valuesfromTaskModule.TeamId;
                        var result = await turnContext.SendActivityAsync(MessageFactory.Attachment(EndorseCard.GetEndorseCard(this.appBaseUrl, this.localizer, valuesfromTaskModule)));
                        turnContext.Activity.Conversation.Id = $"{valuesfromTaskModule.TeamId};messageid={result.Id}";
                        mentionActivity = await CardHelper.GetMentionActivityAsync(valuesfromTaskModule.NominatedToPrincipalName.Split(",").Select(row => row.Trim()).ToList(), turnContext.Activity.From.AadObjectId, valuesfromTaskModule.TeamId, turnContext, this.localizer, this.logger, cancellationToken);
                        await turnContext.SendActivityAsync(mentionActivity);
                        this.logger.LogInformation("Nominated an award");
                        break;

                    case Constants.OkCommand:
                        return null;

                    default:
                        this.logger.LogInformation($"Invalid command for task module submit activity.Command is : {valuesfromTaskModule.Command} ");
                        await turnContext.SendActivityAsync(this.localizer.GetString("ErrorMessage"));
                        break;
                }

                return null;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error at OnTeamsTaskModuleSubmitAsync(): {ex.Message}", SeverityLevel.Error);
                await turnContext.SendActivityAsync(this.localizer.GetString("ErrorMessage"));
                throw;
            }
        }

        /// <summary>
        /// Invoked when the user opens the messaging extension or searching any content in it.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="query">Contains messaging extension query keywords.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Messaging extension response object to fill compose extension section.</returns>
        /// <remarks>
        /// https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionqueryasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionQuery query,
            CancellationToken cancellationToken)
        {
            IInvokeActivity turnContextActivity = turnContext?.Activity;
            try
            {
                if (!await this.CheckTeamsValidaionAsync(turnContext))
                {
                    return new MessagingExtensionResponse
                    {
                        ComposeExtension = new MessagingExtensionResult
                        {
                            Text = this.localizer.GetString("InvalidTeamText"),
                            Type = "message",
                        },
                    };
                }

                MessagingExtensionQuery messageExtensionQuery = JsonConvert.DeserializeObject<MessagingExtensionQuery>(turnContextActivity.Value.ToString());
                string searchQuery = SearchHelper.GetSearchQueryString(messageExtensionQuery);
                turnContextActivity.TryGetChannelData<TeamsChannelData>(out var teamsChannelData);
                var cycleStatus = await this.rewardCycleStorageProvider.GetActiveRewardCycleAsync(teamsChannelData.Channel.Id);

                if (cycleStatus != null)
                {
                    return new MessagingExtensionResponse
                    {
                        ComposeExtension = await SearchHelper.GetSearchResultAsync(this.appBaseUrl, searchQuery, cycleStatus.CycleId, teamsChannelData.Channel.Id, messageExtensionQuery.QueryOptions.Count, messageExtensionQuery.QueryOptions.Skip, this.nominateDetailSearchService, this.localizer),
                    };
                }

                return new MessagingExtensionResponse
                {
                    ComposeExtension = new MessagingExtensionResult
                    {
                        Text = this.localizer.GetString("CycleValidationMessage"),
                        Type = "message",
                    },
                };
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to handle the messaging extension command {turnContextActivity.Name}: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Validates endorsement status.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="valuesforTaskModule">Get the binded values from the card.</param>
        /// <returns>Returns the true, if endorsement is successful, else false.</returns>
        private async Task<bool> CheckEndorseStatusAsync(ITurnContext<IInvokeActivity> turnContext, AdaptiveCardAction valuesforTaskModule)
        {
            var teamsDetails = turnContext.Activity.TeamsGetTeamInfo();
            var teamsChannelAccounts = await TeamsInfo.GetTeamMembersAsync(turnContext, teamsDetails.Id, CancellationToken.None);
            var userDetails = teamsChannelAccounts.Where(member => member.AadObjectId == turnContext.Activity.From.AadObjectId).FirstOrDefault();
            var endorseEntity = await this.endorseDetailStorageProvider.GetEndorseDetailAsync(teamsDetails.Id, valuesforTaskModule.RewardCycleId, valuesforTaskModule.NominatedToPrincipalName);
            var result = endorseEntity.Where(row => row.EndorseForAwardId == valuesforTaskModule.AwardId && row.EndorsedByObjectId == userDetails.AadObjectId).FirstOrDefault();
            if (result == null)
            {
                var endorsedetails = new EndorseEntity
                {
                    TeamId = teamsDetails.Id,
                    EndorsedByObjectId = userDetails.AadObjectId,
                    EndorsedByPrincipalName = userDetails.Email,
                    EndorseForAward = valuesforTaskModule.AwardName,
                    EndorsedToPrincipalName = valuesforTaskModule.NominatedToPrincipalName,
                    EndorsedToObjectId = valuesforTaskModule.NominatedToObjectId,
                    EndorsedOn = DateTime.UtcNow,
                    EndorseForAwardId = valuesforTaskModule.AwardId,
                    AwardCycle = valuesforTaskModule.RewardCycleId,
                };

                return await this.endorseDetailStorageProvider.StoreOrUpdateEndorseDetailAsync(endorsedetails);
            }

            return false;
        }

        /// <summary>
        /// Validates bot is part of teams channel.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>Returns the true, if bot is installed in that channel , else false.</returns>
        private async Task<bool> CheckTeamsValidaionAsync(ITurnContext<IInvokeActivity> turnContext)
        {
            var teamsDetails = turnContext.Activity.TeamsGetTeamInfo();
            if (teamsDetails == null)
            {
                return false;
            }

            var teamData = await this.teamStorageProvider.GetTeamDetailAsync(teamsDetails.Id);
            return teamData != null && teamData.TeamId == teamsDetails.Id;
        }

        /// <summary>
        /// Records event occurred in the application in Application Insights telemetry client.
        /// </summary>
        /// <param name="eventName"> Name of the event.</param>
        /// <param name="turnContext"> Context object containing information cached for a single turn of conversation with a user.</param>
        private void RecordEvent(string eventName, ITurnContext turnContext)
        {
            this.telemetryClient.TrackEvent(eventName, new Dictionary<string, string>
            {
                { "userId", turnContext.Activity.From.AadObjectId },
                { "tenantId", turnContext.Activity.Conversation.TenantId },
            });
        }

        /// <summary>
        /// Send typing indicator to the user.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <returns>A task that represents typing indicator activity.</returns>
        private async Task SendTypingIndicatorAsync(ITurnContext turnContext)
        {
            try
            {
                var typingActivity = turnContext.Activity.CreateReply();
                typingActivity.Type = ActivityTypes.Typing;
                await turnContext.SendActivityAsync(typingActivity);
            }
#pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
#pragma warning restore CA1031 // Do not catch general exception types
            {
                // Do not fail on errors sending the typing indicator
                this.logger.LogWarning(ex, "Failed to send a typing indicator.");
            }
        }
    }
}