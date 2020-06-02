// <copyright file="INotificationHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.BackgroundService
{
    using System.Threading.Tasks;

    /// <summary>
    /// Interface for notification helper.
    /// </summary>
    public interface INotificationHelper
    {
        /// <summary>
        /// This method is used to send nomination reminder notification to teams channel.
        /// </summary>
        /// <returns>A <see cref="Task"/>Returns true if reward cycle set successfully, else false.</returns>
        Task<bool> SendNominationReminderNotificationAsync();
    }
}
