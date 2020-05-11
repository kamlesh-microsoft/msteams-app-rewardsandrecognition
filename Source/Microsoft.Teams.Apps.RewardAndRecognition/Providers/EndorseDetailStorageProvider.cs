// <copyright file="EndorseDetailStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Endorse storage provider.
    /// </summary>
    public class EndorseDetailStorageProvider : StorageBaseProvider, IEndorseDetailStorageProvider
    {
        /// <summary>
        /// Endorse detail table.
        /// </summary>
        private const string EndorseTable = "EndorseDetail";

        /// <summary>
        /// Initializes a new instance of the <see cref="EndorseDetailStorageProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        public EndorseDetailStorageProvider(IOptionsMonitor<StorageOptions> storageOptions)
            : base(storageOptions, EndorseTable)
        {
            if (storageOptions == null)
            {
                throw new ArgumentNullException(nameof(storageOptions));
            }
        }

        /// <summary>
        /// Store or update endorse details in Azure table storage.
        /// </summary>
        /// <param name="endorseEntity">Represents endorse entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents endorse entity is saved or updated.</returns>
        public async Task<bool> StoreOrUpdateEndorseDetailAsync(EndorseEntity endorseEntity)
        {
            await this.EnsureInitializedAsync();
            endorseEntity = endorseEntity ?? throw new ArgumentNullException(nameof(endorseEntity));
            endorseEntity.RowKey = Guid.NewGuid().ToString();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(endorseEntity);
            var result = await this.CloudTable.ExecuteAsync(addOrUpdateOperation);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Get already saved endorse details from Azure Table Storage.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="awardCycleId">Active award cycle.</param>
        /// <param name="nominatedToPrincipalName">Nominated to name</param>
        /// <returns><see cref="Task"/> Already endorse details.</returns>
        public async Task<IEnumerable<EndorseEntity>> GetEndorseDetailAsync(string teamId, string awardCycleId, string nominatedToPrincipalName)
        {
            await this.EnsureInitializedAsync();

            var endorseEntity = new List<EndorseEntity>();
            string partitionKeyCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, teamId);
            string awardCycleIdCondition = TableQuery.GenerateFilterCondition("AwardCycle", QueryComparisons.Equal, awardCycleId);
            string nominatedToPrincipalNameCondition = TableQuery.GenerateFilterCondition("EndorsedToPrincipalName", QueryComparisons.Equal, nominatedToPrincipalName);
            string condition = TableQuery.CombineFilters(partitionKeyCondition, TableOperators.And, awardCycleIdCondition);

            if (!string.IsNullOrWhiteSpace(nominatedToPrincipalName))
            {
                condition = TableQuery.CombineFilters(condition, TableOperators.And, nominatedToPrincipalNameCondition);
            }

            TableQuery<EndorseEntity> query = new TableQuery<EndorseEntity>().Where(condition);
            TableContinuationToken tableContinuationToken = null;

            do
            {
                var queryResponse = await this.CloudTable.ExecuteQuerySegmentedAsync(query, tableContinuationToken);
                tableContinuationToken = queryResponse.ContinuationToken;
                endorseEntity.AddRange(queryResponse?.Results);
            }
            while (tableContinuationToken != null);

            return endorseEntity;
        }
    }
}