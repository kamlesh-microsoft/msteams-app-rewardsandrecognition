// <copyright file="SearchHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;

    /// <summary>
    /// Class that handles the search activities for messaging extension.
    /// </summary>
    public static class SearchHelper
    {
        /// <summary>
        /// Represents the image pixel height.
        /// </summary>
        private const int PixelHeight = 80;

        /// <summary>
        /// Represents the image pixel width.
        /// </summary>
        private const int PixelWidth = 110;

        /// <summary>
        /// Search text parameter name defined in the application manifest file.
        /// </summary>
        private const string SearchTextParameterName = "searchText";

        /// <summary>
        /// Get the value of the searchText parameter in the messaging extension query.
        /// </summary>
        /// <param name="query">Contains messaging extension query keywords.</param>
        /// <returns>A value of the searchText parameter.</returns>
        public static string GetSearchQueryString(MessagingExtensionQuery query)
        {
            var messageExtensionInputText = query?.Parameters.FirstOrDefault(parameter => parameter.Name.Equals(SearchTextParameterName, StringComparison.OrdinalIgnoreCase));
            return messageExtensionInputText?.Value?.ToString();
        }

        /// <summary>
        /// Get the results from Azure search service and populate the result (card + preview).
        /// </summary>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="query">Query which the user had typed in message extension search.</param>
        /// <param name="cycleId">Current reward cycle id.</param>
        /// <param name="teamId">Get the results based on the TeamId.</param>
        /// <param name="count">Count for pagination.</param>
        /// <param name="skip">Skip for pagination.</param>
        /// <param name="searchService">Search service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns><see cref="Task"/> Returns MessagingExtensionResult which will be used for providing the card.</returns>
        public static async Task<MessagingExtensionResult> GetSearchResultAsync(
            string applicationBasePath,
            string query,
            string cycleId,
            string teamId,
            int? count,
            int? skip,
            INominateDetailSearchService searchService,
            IStringLocalizer<Strings> localizer)
        {
            MessagingExtensionResult composeExtensionResult = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = AttachmentLayoutTypes.List,
                Attachments = new List<MessagingExtensionAttachment>(),
            };

            IList<NominateEntity> searchServiceResults;
            searchServiceResults = await searchService?.SearchNominationDetailsAsync(query, cycleId, teamId, count, skip);
            if (searchServiceResults != null)
            {
                foreach (var nominatedDetail in searchServiceResults)
                {
                    AdaptiveCard endorseAdaptiveCard = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
                    {
                        Body = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Width = "50",
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Text = nominatedDetail.AwardName,
                                                HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                                Weight = AdaptiveTextWeight.Bolder,
                                                Size = AdaptiveTextSize.Large,
                                            },
                                        },
                                    },
                                    new AdaptiveColumn
                                    {
                                        Width = "100",
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveImage
                                            {
                                                Url = string.IsNullOrEmpty(nominatedDetail.AwardImageLink) ? new Uri(string.Format(CultureInfo.InvariantCulture, "{0}/Content/DefaultAwardImage.png", applicationBasePath?.Trim('/'))) : new Uri(nominatedDetail.AwardImageLink),
                                                HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                                PixelHeight = PixelHeight,
                                                PixelWidth = PixelWidth,
                                            },
                                        },
                                    },
                                },
                            },
                            new AdaptiveTextBlock
                            {
                                Text = nominatedDetail.NominatedToName,
                                Wrap = true,
                                HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                Weight = AdaptiveTextWeight.Bolder,
                                Spacing = AdaptiveSpacing.Large,
                            },
                            new AdaptiveTextBlock
                            {
                                Text = localizer.GetString("NominatedByText", nominatedDetail.NominatedByName),
                                Wrap = true,
                                HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                Spacing = AdaptiveSpacing.Default,
                            },
                            new AdaptiveTextBlock
                            {
                                Text = nominatedDetail.ReasonForNomination,
                                Wrap = true,
                                HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                Spacing = AdaptiveSpacing.Default,
                            },
                        },
                        Actions = new List<AdaptiveAction>
                        {
                            new AdaptiveSubmitAction
                            {
                                Title = localizer.GetString("EndorseButtonText"),
                                Data = new AdaptiveCardAction
                                {
                                    MsteamsCardAction = new CardAction
                                    {
                                        Type = Constants.FetchActionType,
                                    },
                                    Command = Constants.EndorseAction,
                                    NominatedToPrincipalName = nominatedDetail.NominatedToPrincipalName,
                                    AwardName = nominatedDetail.AwardName,
                                    NominatedToName = nominatedDetail.NominatedToName,
                                    NominatedToObjectId = nominatedDetail.NominatedToObjectId,
                                    RewardCycleId = nominatedDetail.RewardCycleId,
                                },
                            },
                        },
                    };

                    ThumbnailCard previewCard = new ThumbnailCard
                    {
                        Title = nominatedDetail.NominatedToName,
                        Subtitle = nominatedDetail.AwardName,
                    };

                    composeExtensionResult.Attachments.Add(new Attachment
                    {
                        ContentType = AdaptiveCard.ContentType,
                        Content = endorseAdaptiveCard,
                    }.ToMessagingExtensionAttachment(previewCard.ToAttachment()));
                }
            }

            return composeExtensionResult;
        }
    }
}
