// <copyright file="EndorseCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    ///  This class process endorse card when award is nominated.
    /// </summary>
    public static class EndorseCard
    {
        /// <summary>
        /// Represents the image pixel height.
        /// </summary>
        private const int PixelHeight = 85;

        /// <summary>
        /// Represents the image pixel width.
        /// </summary>
        private const int PixelWidth = 115;

        /// <summary>
        /// Represents the image pixel height/width of endorse icon.
        /// </summary>
        private const int PixelWidthIcon = 15;

        /// <summary>
        /// Represents the image pixel height of endorse container.
        /// </summary>
        private const int PixelMinHeightContainer = 144;

        /// <summary>
        /// This method will construct endorse card with corresponding details.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="nominatedDetails">Nominated details to show in card.</param>
        /// <returns>Endorse card with nominated details.</returns>
        public static Attachment GetEndorseCard(string applicationBasePath, IStringLocalizer<Strings> localizer, TaskModuleResponseDetails nominatedDetails)
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = nominatedDetails?.AwardName,
                                        Wrap = true,
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Size = AdaptiveTextSize.Large,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = string.IsNullOrEmpty(nominatedDetails.AwardLink) ? new Uri(string.Format(CultureInfo.InvariantCulture, "{0}/Content/DefaultAwardImage.png", applicationBasePath?.Trim('/'))) : new Uri(nominatedDetails.AwardLink),
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
                        Text = nominatedDetails.NominatedToName,
                        Wrap = true,
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Weight = AdaptiveTextWeight.Bolder,
                        Spacing = AdaptiveSpacing.Large,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("NominatedByText", nominatedDetails.NominatedByName),
                        Wrap = true,
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Spacing = AdaptiveSpacing.Default,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = nominatedDetails.ReasonForNomination,
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
                            NominatedToPrincipalName = nominatedDetails.NominatedToPrincipalName,
                            AwardName = nominatedDetails.AwardName,
                            NominatedToName = nominatedDetails.NominatedToName,
                            NominatedToObjectId = nominatedDetails.NominatedToObjectId,
                            AwardId = nominatedDetails.AwardId,
                            RewardCycleId = nominatedDetails.RewardCycleId,
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
        }

        /// <summary>
        /// Construct the card to render endorse message to task module.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="awardName">Award name.</param>
        /// <param name="nominatedToName">Nominated users.</param>
        /// <param name="rewardCycleEndDate">Cycle end date.</param>
        /// <param name="isEndorsementSuccess">Gets the endorsement status.</param>
        /// <returns>Card attachment.</returns>
        public static Attachment GetEndorseStatusCard(string applicationBasePath, IStringLocalizer<Strings> localizer, string awardName, string nominatedToName, DateTime rewardCycleEndDate, bool isEndorsementSuccess)
        {
            var endCycleDate = "{{DATE(" + rewardCycleEndDate.ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.InvariantCulture) + ", SHORT)}}";
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveContainer()
                    {
                        PixelMinHeight = PixelMinHeightContainer,
                        Items = new List<AdaptiveElement>()
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Width = AdaptiveColumnWidth.Auto,
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveImage
                                            {
                                                Spacing = AdaptiveSpacing.Large,
                                                PixelHeight = PixelWidthIcon,
                                                PixelWidth = PixelWidthIcon,
                                                Url = isEndorsementSuccess ? new Uri(string.Format(CultureInfo.InvariantCulture, "{0}/Content/SuccessIcon.png", applicationBasePath?.Trim('/'))) : new Uri(string.Format(CultureInfo.InvariantCulture, "{0}/Content/ErrorIcon.png", applicationBasePath?.Trim('/'))),
                                            },
                                        },
                                    },
                                    new AdaptiveColumn
                                    {
                                        Width = AdaptiveColumnWidth.Stretch,
                                        Height = AdaptiveHeight.Auto,
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Text = isEndorsementSuccess == true ? localizer.GetString("SuccessfulEndorseMessage", nominatedToName, awardName, endCycleDate) : localizer.GetString("AlreadyendorsedMessage", nominatedToName),
                                                Wrap = true,
                                                Size = AdaptiveTextSize.Default,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("OkButtonText"),
                        Data = new AdaptiveCardAction
                        {
                            MsteamsCardAction = new CardAction
                            {
                                Type = Constants.MessageBackActionType,
                            },
                            Command = Constants.OkCommand,
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
        }
    }
}
