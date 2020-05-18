// <copyright file="INominateAwardStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    /// Interface for Nominate award storage provider.
    /// </summary>
    public interface INominateAwardStorageProvider
    {
        /// <summary>
        /// This method is used to fetch nomination details for a given team Id and aadObjectId.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="nominatedToObjectId">Azure active directory object Id.</param>
        /// <param name="awardCycleId">Active award cycle.</param>
        /// <param name="awardId">Award unique id.</param>
        /// <param name="nominatedByObjectId">Azure active directory object Id of nominator.</param>
        /// <returns>Returns true if same group of user already nominated, else return false.</returns>
        Task<bool> GetNominateDetailsAsync(string teamId, string nominatedToObjectId, string awardCycleId, string awardId, string nominatedByObjectId);

        /// <summary>
        /// Get already nominate entity detail from storage table.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns><see cref="Task"/> Already saved nominate entity detail.</returns>
        Task<NominateEntity> GetNominationAwardDetailsAsync(string teamId);

        /// <summary>
        /// Store or update Nominated award details in table storage.
        /// </summary>
        /// <param name="nominateEntity">Represents nominate entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/>Returns nominate entity which is saved.</returns>
        Task<NominateEntity> StoreOrUpdateNominatedDetailsAsync(NominateEntity nominateEntity);

        /// <summary>
        /// This method is used to fetch nomination details for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="isAwardGranted">Flag is award granted.</param>
        /// <param name="awardCycleId">Active award cycle.</param>
        /// <returns>Nomination details.</returns>
        Task<IEnumerable<NominateEntity>> GetNominationDetailsAsync(string teamId, bool isAwardGranted, string awardCycleId);

        /// <summary>
        /// This method is used to publish nomination details for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="nominationIds">Published nomination ids.</param>
        /// <returns>Nomination details.</returns>
        Task<bool> PublishNominationDetailsAsync(string teamId, IEnumerable<string> nominationIds);
    }
}
