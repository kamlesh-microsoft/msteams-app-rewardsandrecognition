// <copyright file="TeamStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System;
    using System.Linq;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Awards storage provider.
    /// </summary>
    public class TeamStorageProvider : StorageBaseProvider, ITeamStorageProvider
    {
        private const string TeamConfigurationTable = "TeamConfiguration";

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamStorageProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        public TeamStorageProvider(IOptionsMonitor<StorageOptions> storageOptions)
            : base(storageOptions, TeamConfigurationTable)
        {
            if (storageOptions == null)
            {
                throw new ArgumentNullException(nameof(storageOptions));
            }
        }

        /// <summary>
        /// Store or update team detail in Azure table storage.
        /// </summary>
        /// <param name="teamEntity">Represents team entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents team entity is saved or updated.</returns>
        public async Task<bool> StoreOrUpdateTeamDetailAsync(TeamEntity teamEntity)
        {
            await this.EnsureInitializedAsync();
            teamEntity = teamEntity ?? throw new ArgumentNullException(nameof(teamEntity));
            teamEntity.RowKey = Guid.NewGuid().ToString();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(teamEntity);
            var result = await this.CloudTable.ExecuteAsync(addOrUpdateOperation);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Get already team detail from Azure table storage.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns><see cref="Task"/> Already saved team detail.</returns>
        public async Task<TeamEntity> GetTeamDetailAsync(string teamId)
        {
            await this.EnsureInitializedAsync();
            var teamEntity = new TeamEntity();
            var query = new TableQuery<TeamEntity>().Where(TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, teamId));
            TableContinuationToken tableContinuationToken = null;

            do
            {
                var queryResponse = await this.CloudTable.ExecuteQuerySegmentedAsync(query, tableContinuationToken);
                tableContinuationToken = queryResponse.ContinuationToken;
                teamEntity = queryResponse.Results.FirstOrDefault();
            }
            while (tableContinuationToken != null);

            return teamEntity;
        }

        /// <summary>
        /// This method delete the team detail record from table.
        /// </summary>
        /// <param name="teamEntity">Team configuration table entity.</param>
        /// <returns>A <see cref="Task"/> of type bool where true represents entity record is successfully deleted from table while false indicates failure in deleting data.</returns>
        public async Task<bool> DeleteTeamDetailAsync(TeamEntity teamEntity)
        {
            await this.EnsureInitializedAsync().ConfigureAwait(false);
            TableOperation insertOrMergeOperation = TableOperation.Delete(teamEntity);
            TableResult result = await this.CloudTable.ExecuteAsync(insertOrMergeOperation).ConfigureAwait(false);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }
    }
}
