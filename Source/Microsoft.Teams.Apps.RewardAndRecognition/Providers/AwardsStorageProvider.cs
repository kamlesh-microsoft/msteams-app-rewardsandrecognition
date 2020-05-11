// <copyright file="AwardsStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Awards storage provider.
    /// </summary>
    public class AwardsStorageProvider : StorageBaseProvider, IAwardsStorageProvider
    {
        /// <summary>
        /// Award detail table.
        /// </summary>
        private const string AwardTable = "AwardDetail";

        /// <summary>
        /// Initializes a new instance of the <see cref="AwardsStorageProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        public AwardsStorageProvider(IOptionsMonitor<StorageOptions> storageOptions)
            : base(storageOptions, AwardTable)
        {
            if (storageOptions == null)
            {
                throw new ArgumentNullException(nameof(storageOptions));
            }
        }

        /// <summary>
        /// This method is used to fetch all the awards for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns>All the awards for a given team Id.</returns>
        public async Task<IEnumerable<AwardEntity>> GetAwardsAsync(string teamId)
        {
            await this.EnsureInitializedAsync();
            string filter = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, teamId);
            var query = new TableQuery<AwardEntity>().Where(filter);
            TableContinuationToken continuationToken = null;
            var awards = new List<AwardEntity>();

            do
            {
                var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                awards.AddRange(queryResult?.Results);
                continuationToken = queryResult?.ContinuationToken;
            }
            while (continuationToken != null);

            return awards.OrderByDescending(record => record.Timestamp);
        }

        /// <summary>
        /// This method is used to fetch award details for a given team Id and awardId.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="awardId">Award Id.</param>
        /// <returns>Award details.</returns>
        public async Task<AwardEntity> GetAwardDetailsAsync(string teamId, string awardId)
        {
            await this.EnsureInitializedAsync();
            var operation = TableOperation.Retrieve<AwardEntity>(teamId, awardId);
            var award = await this.CloudTable.ExecuteAsync(operation);
            return award.Result as AwardEntity;
        }

        /// <summary>
        /// Store or update awards in table storage.
        /// </summary>
        /// <param name="awardEntity">Represents award entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents award entity is saved or updated.</returns>
        public async Task<AwardEntity> StoreOrUpdateAwardAsync(AwardEntity awardEntity)
        {
            await this.EnsureInitializedAsync();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(awardEntity);
            var result = await this.CloudTable.ExecuteAsync(addOrUpdateOperation);
            return result.Result as AwardEntity;
        }

        /// <summary>
        /// Delete award details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="awardIds">Holds award Id data.</param>
        /// <returns>A task that represents award entity data is saved or updated.</returns>
        public async Task<bool> DeleteAwardsAsync(IEnumerable<string> awardIds)
        {
            if (awardIds == null)
            {
                throw new ArgumentNullException(nameof(awardIds));
            }

            await this.EnsureInitializedAsync();
            AwardEntity entity;

            foreach (var awardId in awardIds)
            {
                string awardIdCondition = TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, awardId);
                TableQuery<AwardEntity> query = new TableQuery<AwardEntity>().Where(awardIdCondition);
                var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
                entity = queryResult?.Results[0];
                TableOperation deleteOperation = TableOperation.Delete(entity);
                var result = await this.CloudTable.ExecuteAsync(deleteOperation);
            }

            return true;
        }
    }
}
