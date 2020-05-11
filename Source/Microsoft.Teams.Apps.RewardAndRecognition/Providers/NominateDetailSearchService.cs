// <copyright file="NominateDetailSearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Search;
    using Microsoft.Azure.Search.Models;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    /// Search service which will help in creating index, indexer and data source if it doesn't exist
    /// for indexing table which will be used for search by message extension.
    /// </summary>
    public class NominateDetailSearchService : INominateDetailSearchService, IDisposable
    {
        private const string NominateDetailIndexName = "nominatedetaildata-index";
        private const string NominateDetailIndexerName = "nominatedetaildata-indexer";
        private const string NominateDetailDataSourceName = "nominatedetaildata-storage";

        // Default to 10 results, same as page size of a messaging extension query
        private const int DefaultSearchResultCount = 10;
        private readonly Lazy<Task> initializeTask;
        private readonly SearchServiceClient searchServiceClient;
        private readonly SearchIndexClient searchIndexClient;
        private readonly int searchIndexingIntervalInMinutes;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly SearchServiceOptions searchServiceOptions;

        /// <summary>
        /// Provider to store award nomination details to Azure table storage.
        /// </summary>
        private readonly INominateAwardStorageProvider nominateAwardStorageProvider;

        // Flag: Has Dispose already been called?
        private bool disposed = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="NominateDetailSearchService"/> class.
        /// </summary>
        /// <param name="searchServiceOptions">A set of key/value application configuration properties.</param>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        /// <param name="nominateAwardStorageProvider">Provider to store on nomination details in Azure table storage.</param>
        public NominateDetailSearchService(
            IOptionsMonitor<SearchServiceOptions> searchServiceOptions,
            IOptionsMonitor<StorageOptions> storageOptions,
            INominateAwardStorageProvider nominateAwardStorageProvider)
        {
            searchServiceOptions = searchServiceOptions ?? throw new ArgumentNullException(nameof(searchServiceOptions));
            storageOptions = storageOptions ?? throw new ArgumentNullException(nameof(storageOptions));

            this.searchServiceOptions = searchServiceOptions.CurrentValue;
            string searchServiceValue = this.searchServiceOptions.SearchServiceName;
            this.searchServiceClient = new SearchServiceClient(
                searchServiceValue,
                new SearchCredentials(this.searchServiceOptions.SearchServiceAdminApiKey));
            this.searchIndexClient = new SearchIndexClient(
                searchServiceValue,
                NominateDetailIndexName,
                new SearchCredentials(this.searchServiceOptions.SearchServiceQueryApiKey));
            this.searchIndexingIntervalInMinutes = Convert.ToInt32(this.searchServiceOptions.SearchIndexingIntervalInMinutes, CultureInfo.InvariantCulture);
            this.nominateAwardStorageProvider = nominateAwardStorageProvider;
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(storageOptions.CurrentValue.ConnectionString));
        }

        /// <summary>
        /// Provide search result for table to be used by SME based on Azure search service.
        /// </summary>
        /// <param name="searchQuery">searchQuery to be provided by message extension.</param>
        /// <param name="cycleId">Current reward cycle id.</param>
        /// <param name="teamId">Get the results based on the TeamId.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <returns>List of search results.</returns>
        public async Task<IList<NominateEntity>> SearchNominationDetailsAsync(string searchQuery, string cycleId, string teamId, int? count = null, int? skip = null)
        {
            await this.EnsureInitializedAsync();
            IList<NominateEntity> nominateEntity = new List<NominateEntity>();

            SearchParameters searchParameters = new SearchParameters
            {
                OrderBy = new[] { "Timestamp desc" },
                Filter = $"TeamId eq '{teamId}' and RewardCycleId eq '{cycleId}'",
                Top = count ?? DefaultSearchResultCount,
                Skip = skip ?? 0,
                IncludeTotalResultCount = false,
                Select = new[] { "AwardName", "AwardId", "AwardImageLink", "NominatedOn", "NominatedToName", "NominatedByPrincipalName", "NominatedByObjectId", "NominatedToObjectId", "ReasonForNomination", "NominatedToObjectId", "NominatedByName", "NominatedToPrincipalName", "IsGroupNomination", "GroupName", "RewardCycleId" },
            };

            var docs = await this.searchIndexClient.Documents.SearchAsync<NominateEntity>(searchQuery, searchParameters);
            if (docs != null)
            {
                nominateEntity = docs.Results.Select(result => result.Document).ToList();
            }

            return nominateEntity;
        }

        /// <summary>
        /// Dispose of unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Protected implementation of Dispose pattern.
        /// </summary>
        /// <param name="disposing">True if already disposed else false.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (this.disposed)
            {
                return;
            }

            if (disposing)
            {
                this.searchIndexClient.Dispose();
                this.searchServiceClient.Dispose();
            }

            this.disposed = true;
        }

        /// <summary>
        /// Create index, indexer and data source it doesn't exist.
        /// </summary>
        /// <param name="storageConnectionString">Connection string to the data store.</param>
        /// <returns>Tracking task.</returns>
        private async Task InitializeAsync(string storageConnectionString)
        {
            await this.nominateAwardStorageProvider.GetNominationAwardDetailsAsync(string.Empty); // When there is no nominations and messaging extension is open by team members, table initialization is required here before creating search index or data source or indexer.
            await this.CreateIndexAsync();
            await this.CreateDataSourceAsync(storageConnectionString);
            await this.CreateIndexerAsync();
        }

        /// <summary>
        /// Create index in Azure search service if it doesn't exist.
        /// </summary>
        /// <returns><see cref="Task"/> That represents index is created if it is not created.</returns>
        private async Task CreateIndexAsync()
        {
            if (!this.searchServiceClient.Indexes.Exists(NominateDetailIndexName))
            {
                var tableIndex = new Index()
                {
                    Name = NominateDetailIndexName,
                    Fields = FieldBuilder.BuildForType<NominateEntity>(),
                };
                await this.searchServiceClient.Indexes.CreateAsync(tableIndex);
            }
        }

        /// <summary>
        /// Add data source if it doesn't exist in Azure search service.
        /// </summary>
        /// <param name="connectionString">Connection string to the data store.</param>
        /// <returns><see cref="Task"/> That represents data source is added to Azure search service.</returns>
        private async Task CreateDataSourceAsync(string connectionString)
        {
            if (!this.searchServiceClient.DataSources.Exists(NominateDetailDataSourceName))
            {
                var dataSource = DataSource.AzureTableStorage(
                    name: NominateDetailDataSourceName,
                    storageConnectionString: connectionString,
                    tableName: Constants.NominateAwardTable);

                await this.searchServiceClient.DataSources.CreateAsync(dataSource);
            }
        }

        /// <summary>
        /// Create indexer if it doesn't exist in Azure search service.
        /// </summary>
        /// <returns><see cref="Task"/> That represents indexer is created if not available in Azure search service.</returns>
        private async Task CreateIndexerAsync()
        {
            if (!this.searchServiceClient.Indexers.Exists(NominateDetailIndexerName))
            {
                var indexer = new Indexer()
                {
                    Name = NominateDetailIndexerName,
                    DataSourceName = NominateDetailDataSourceName,
                    TargetIndexName = NominateDetailIndexName,
                    Schedule = new IndexingSchedule(TimeSpan.FromMinutes(this.searchIndexingIntervalInMinutes)),
                };

                await this.searchServiceClient.Indexers.CreateAsync(indexer);
            }
        }

        /// <summary>
        /// Initialization of InitializeAsync method which will help in indexing.
        /// </summary>
        /// <returns>Task with initialized data.</returns>
        private Task EnsureInitializedAsync()
        {
            return this.initializeTask.Value;
        }
    }
}
