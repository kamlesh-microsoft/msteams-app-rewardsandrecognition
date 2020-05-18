// <copyright file="RewardAndRecognitionActivityHandlerOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition
{
    /// <summary>
    /// The RewardAndRecognitionActivityHandlerOptions are the options for the <see cref="RewardAndRecognitionActivityHandler" /> bot.
    /// </summary>
    public sealed class RewardAndRecognitionActivityHandlerOptions
    {
        /// <summary>
        /// Gets or sets a value indicating whether the response to a message should be all uppercase.
        /// </summary>
        public bool UpperCaseResponse { get; set; }

        /// <summary>
        /// Gets or sets unique id of Tenant.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets application base Uri.
        /// </summary>
        public string AppBaseUri { get; set; }

        /// <summary>
        /// Gets or sets unique id of manifest.
        /// </summary>
        public string ManifestId { get; set; }
    }
}