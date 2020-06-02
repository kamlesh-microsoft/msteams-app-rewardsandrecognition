// <copyright file="WinnerCarouselCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
namespace Microsoft.Teams.Apps.RewardAndRecognition.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Web;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.CodeAnalysis;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Newtonsoft.Json;

    /// <summary>
    ///  This class process tour carousel feature to show winners.
    /// </summary>
    public static class WinnerCarouselCard
    {
        /// <summary>
        /// Represents the image pixel height.
        /// </summary>
        private const int PixelHeight = 220;

        /// <summary>
        /// Represents the image pixel width.
        /// </summary>
        private const int PixelWidth = 416;

        /// <summary>
        /// Render the set of attachments that comprise carousel.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="winners">Award winner details.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="manifestId">Unique id of manifest.</param>
        /// <returns>The card that comprise the winner details.</returns>
        public static IEnumerable<Attachment> GetAwardWinnerCard(string applicationBasePath, IEnumerable<AwardWinnerNotification> winners, IStringLocalizer<Strings> localizer, string manifestId)
        {
            var attachments = new List<Attachment>();
            string context = HttpUtility.UrlEncode(JsonConvert.SerializeObject(new { channelId = winners.First().TeamId }));
            foreach (var winner in winners.GroupBy(rows => rows.AwardId))
            {
                AdaptiveCard carouselCard = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
                {
                    Body = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Text = localizer.GetString("AwardWinnerCardTitle"),
                            Weight = AdaptiveTextWeight.Bolder,
                            Size = AdaptiveTextSize.Large,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = $"{localizer.GetString("WinnerCardRewardCycleTitle")}: {winner.First().AwardCycle}",
                            Size = AdaptiveTextSize.Small,
                            Spacing = AdaptiveSpacing.Small,
                            Wrap = true,
                        },
                        new AdaptiveImage
                        {
                            Url = string.IsNullOrEmpty(winner.First().AwardLink) ? new Uri(string.Format(CultureInfo.InvariantCulture, "{0}/Content/DefaultAwardImage.png", applicationBasePath?.Trim('/'))) : new Uri(winner.First().AwardLink),
                            PixelWidth = PixelWidth,
                            PixelHeight = PixelHeight,
                            Size = AdaptiveImageSize.Auto,
                            Style = AdaptiveImageStyle.Default,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = winner.OrderByDescending(row => row.NominatedOn).FirstOrDefault().AwardName,
                            Size = AdaptiveTextSize.Large,
                            Weight = AdaptiveTextWeight.Bolder,
                            Spacing = AdaptiveSpacing.Small,
                            Wrap = true,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = string.Join(", ", string.Join(",", winner.Select(rows => rows.NominatedToName)).Split(",").Distinct()),
                            Size = AdaptiveTextSize.Small,
                            Spacing = AdaptiveSpacing.Medium,
                            Wrap = true,
                        },
                    },
                    Actions = new List<AdaptiveAction>
                    {
                        new AdaptiveOpenUrlAction
                        {
                            Title = localizer.GetString("ViewWinnerTabText"),
                            Url = new System.Uri(string.Format(CultureInfo.InvariantCulture, Constants.TabDeepLink, manifestId, context)),
                        },
                    },
                };

                attachments.Add(new Attachment
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = carouselCard,
                });
            }

            return attachments;
        }
    }
}
