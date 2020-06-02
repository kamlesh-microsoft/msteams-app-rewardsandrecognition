// <copyright file="CardHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Xml;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RewardAndRecognition.Cards;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    /// Class that handles the card configuration.
    /// </summary>
    public static class CardHelper
    {
        /// <summary>
        ///  Represents the task module height.
        /// </summary>
        private const int TaskModuleHeight = 460;

        /// <summary>
        /// Represents the task module width.
        /// </summary>
        private const int TaskModuleWidth = 600;

        /// <summary>
        ///  Represents the nomination task module height.
        /// </summary>
        private const int NominationTaskModuleHeight = 600;

        /// <summary>
        /// Represents the nomination task module width.
        /// </summary>
        private const int NominationTaskModuleWidth = 700;

        /// <summary>
        /// Represents the error message task module height.
        /// </summary>
        private const int ErrorMessageTaskModuleHeight = 200;

        /// <summary>
        /// Represents the error message task module width.
        /// </summary>
        private const int ErrorMessageTaskModuleWidth = 400;

        /// <summary>
        /// Represents the endorse message task module height.
        /// </summary>
        private const int EndorseMessageTaskModuleHeight = 220;

        /// <summary>
        /// Represents the endorse message task module width.
        /// </summary>
        private const int EndorseMessageTaskModuleWidth = 480;

        /// <summary>
        /// Get messaging extension action response.
        /// </summary>
        /// <param name="applicationBasePath">Represents the Application base Uri.</param>
        /// <param name="instrumentationKey">Instrumentation key of the telemetry client.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="teamId">Team id from where the ME action is called.</param>
        /// <param name="isCycleRunning">Gets the value false if cycle is not running currently.</param>
        /// <returns>Returns task module response.</returns>
        public static MessagingExtensionActionResponse GetTaskModuleBasedOnCommand(string applicationBasePath, string instrumentationKey, IStringLocalizer<Strings> localizer, string teamId = null, bool isCycleRunning = true)
        {
            if (!isCycleRunning)
            {
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo()
                        {
                            Card = ValidationMessageCard.GetErrorAdaptiveCard(localizer.GetString("CycleValidationMessage")),
                            Height = ErrorMessageTaskModuleHeight,
                            Width = ErrorMessageTaskModuleWidth,
                            Title = localizer.GetString("NominatePeopleTitle"),
                        },
                    },
                };
            }
            else
            {
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Url = $"{applicationBasePath}/nominate-awards?telemetry={instrumentationKey}&teamId={teamId}&theme={{theme}}&locale={{locale}}",
                            Height = NominationTaskModuleHeight,
                            Width = NominationTaskModuleWidth,
                            Title = localizer.GetString("NominatePeopleTitle"),
                        },
                    },
                };
            }
        }

        /// <summary>
        /// Get messaging extension action response.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns task module response.</returns>
        public static MessagingExtensionActionResponse GetTaskModuleInvalidTeamCard(IStringLocalizer<Strings> localizer)
        {
            return new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        Card = ValidationMessageCard.GetErrorAdaptiveCard(localizer.GetString("InvalidTeamText")),
                        Height = ErrorMessageTaskModuleHeight,
                        Width = ErrorMessageTaskModuleWidth,
                        Title = localizer.GetString("NominatePeopleTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Get task module response.
        /// </summary>
        /// <param name="applicationBasePath">Represents the Application base Uri.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="nominatedToName">Nominated to name.</param>
        /// <param name="awardName">Award name.</param>
        /// <param name="rewardCycleEndDate">Cycle end date.</param>
        /// <param name="isEndorsementSuccess">Get the endorsement status.</param>
        /// <returns>Returns task module response.</returns>
        public static TaskModuleResponse GetEndorseTaskModuleResponse(string applicationBasePath, IStringLocalizer<Strings> localizer, string nominatedToName, string awardName, DateTime rewardCycleEndDate, bool isEndorsementSuccess)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = EndorseCard.GetEndorseStatusCard(applicationBasePath, localizer, awardName, nominatedToName, rewardCycleEndDate, isEndorsementSuccess),
                        Height = EndorseMessageTaskModuleHeight,
                        Width = EndorseMessageTaskModuleWidth,
                        Title = localizer.GetString("EndorseTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Get task module response.
        /// </summary>
        /// <param name="applicationBasePath">Represents the Application base Uri.</param>
        /// <param name="instrumentationKey">Telemetry instrumentation key.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="command">Get the command from the user.</param>
        /// <param name="teamId">Team id from where the ME action is called.</param>
        /// <param name="awardId">Award id to fetch the award details.</param>
        /// <param name="isCycleRunning">Gets the value false if cycle is not running currently.</param>
        /// <param name="isActivityIdPresent">Gets the boolean value based on activity id.</param>
        /// <param name="isCycleClosed">Gets the value true if cycle is closed.</param>
        /// <returns>Returns task module response.</returns>
        public static TaskModuleResponse GetTaskModuleResponse(string applicationBasePath, string instrumentationKey, IStringLocalizer<Strings> localizer, string command, string teamId = null, string awardId = null, bool isCycleRunning = true, bool isActivityIdPresent = true, bool isCycleClosed = false)
        {
            if ((!isCycleRunning || isCycleClosed) && command != Constants.ConfigureAdminAction)
            {
                return new TaskModuleResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo()
                        {
                            Card = ValidationMessageCard.GetErrorAdaptiveCard(isCycleClosed == true ? localizer.GetString("CycleClosedMessage") : localizer.GetString("CycleValidationMessage")),
                            Height = ErrorMessageTaskModuleHeight,
                            Width = ErrorMessageTaskModuleWidth,
                            Title = command == Constants.NominateAction ? localizer.GetString("NominatePeopleTitle") : localizer.GetString("EndorseTitle"),
                        },
                    },
                };
            }
            else if (command == Constants.ConfigureAdminAction)
            {
                return new TaskModuleResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo()
                        {
                            Url = $"{applicationBasePath}/config-admin-page?telemetry={instrumentationKey}&teamId={teamId}&isActivityIdPresent={isActivityIdPresent}&theme={{theme}}&locale={{locale}}",
                            Height = TaskModuleHeight,
                            Width = TaskModuleWidth,
                            Title = localizer.GetString("ConfigureAdminTitle"),
                            FallbackUrl = $"{applicationBasePath}/config-admin-page?telemetry={instrumentationKey}&teamId={teamId}&isActivityIdPresent={isActivityIdPresent}&theme={{theme}}&locale={{locale}}",
                        },
                    },
                };
            }
            else
            {
                return new TaskModuleResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Url = $"{applicationBasePath}/nominate-awards?telemetry={instrumentationKey}&teamId={teamId}&awardId={awardId}&theme={{theme}}&locale={{locale}}",
                            Height = NominationTaskModuleHeight,
                            Width = NominationTaskModuleWidth,
                            Title = localizer.GetString("NominatePeopleTitle"),
                            FallbackUrl = $"{applicationBasePath}/nominate-awards?telemetry={instrumentationKey}&teamId={teamId}&awardId={awardId}&theme={{theme}}&locale={{locale}}",
                        },
                    },
                };
            }
        }

        /// <summary>
        /// Methods mentions user in respective channel of which they are part after grouping.
        /// </summary>
        /// <param name="mentionToEmails">List of email ID whom to be mentioned.</param>
        /// <param name="userObjectId">Azure active directory object id of the user.</param>
        /// <param name="teamId">Team id where bot is installed.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="logger">Instance to send logs to the application insights service.</param>
        /// <param name="mentionType">Mention activity type.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that sends notification in newly created channel and mention its members.</returns>
        internal static async Task<Activity> GetMentionActivityAsync(IEnumerable<string> mentionToEmails, string userObjectId, string teamId, ITurnContext turnContext, IStringLocalizer<Strings> localizer, ILogger logger, MentionActivityType mentionType, CancellationToken cancellationToken)
        {
            try
            {
                StringBuilder mentionText = new StringBuilder();
                List<Entity> entities = new List<Entity>();
                List<Mention> mentions = new List<Mention>();
                IEnumerable<TeamsChannelAccount> channelMembers = await TeamsInfo.GetTeamMembersAsync(turnContext, teamId, cancellationToken);

                IEnumerable<ChannelAccount> mentionToMemberDetails = channelMembers.Where(member => mentionToEmails.Contains(member.Email)).Select(member => new ChannelAccount { Id = member.Id, Name = member.Name });
                ChannelAccount mentionByMemberDetails = channelMembers.Where(member => member.AadObjectId == userObjectId).Select(member => new ChannelAccount { Id = member.Id, Name = member.Name }).FirstOrDefault();

                foreach (ChannelAccount member in mentionToMemberDetails)
                {
                    Mention mention = new Mention
                    {
                        Mentioned = new ChannelAccount()
                        {
                            Id = member.Id,
                            Name = member.Name,
                        },
                        Text = $"<at>{XmlConvert.EncodeName(member.Name)}</at>",
                    };
                    mentions.Add(mention);
                    entities.Add(mention);
                    mentionText.Append(mention.Text).Append(", ");
                }

                Mention mentionBy = new Mention
                {
                    Mentioned = new ChannelAccount()
                    {
                        Id = mentionByMemberDetails.Id,
                        Name = mentionByMemberDetails.Name,
                    },
                    Text = $"<at>{XmlConvert.EncodeName(mentionByMemberDetails.Name)}</at>",
                };

                string text = string.Empty;

                switch (mentionType)
                {
                    case MentionActivityType.SetAdmin:
                        entities.Add(mentionBy);
                        text = localizer.GetString("SetAdminMentionText", mentionText.ToString().Trim().TrimEnd(','), mentionBy.Text);
                        break;
                    case MentionActivityType.Nomination:
                        entities.Add(mentionBy);
                        text = localizer.GetString("NominationMentionText", mentionText.ToString().Trim().TrimEnd(','), mentionBy.Text);
                        break;
                    case MentionActivityType.Winner:
                        text = $"{localizer.GetString("WinnerMentionText")} {mentionText.ToString().Trim().TrimEnd(',')}";
                        break;
                    default:
                        break;
                }

                Activity notificationActivity = MessageFactory.Text(text);
                notificationActivity.Entities = entities;
                return notificationActivity;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, $"Error while mentioning channel member in respective channels.");
                return null;
            }
        }
    }
}