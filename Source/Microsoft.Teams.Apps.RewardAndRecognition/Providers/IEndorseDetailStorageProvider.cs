// <copyright file="IEndorseDetailStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    /// Interface for endorse detail storage provider.
    /// </summary>
    public interface IEndorseDetailStorageProvider
    {
        /// <summary>
        /// Store or update endorse details in table storage.
        /// </summary>
        /// <param name="endorseEntity">Represents endorse entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents endorse entity is saved or updated.</returns>
        Task<bool> StoreOrUpdateEndorseDetailAsync(EndorseEntity endorseEntity);

        /// <summary>
        /// Get already saved teams entity from storage table.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="awardCycleId">Active award cycle.</param>
        /// <param name="nominatedToPrincipalName">Nominated to name</param>
        /// <returns><see cref="Task"/>Returns endorse entity which is already saved.</returns>
        Task<IEnumerable<EndorseEntity>> GetEndorseDetailAsync(string teamId, string awardCycleId, string nominatedToPrincipalName);
    }
}
