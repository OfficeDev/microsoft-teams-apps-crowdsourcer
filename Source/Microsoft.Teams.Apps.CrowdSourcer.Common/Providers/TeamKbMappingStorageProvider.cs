// <copyright file="TeamKbMappingStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.CrowdSourcer.Common.Models;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// StorageProvider class.
    /// </summary>
    public class TeamKbMappingStorageProvider : ITeamKbMappingStorageProvider
    {
        private const string CrowdSourcerTableName = "crowdsourcer";
        private readonly Lazy<Task> initializeTask;
        private CloudTableClient cloudTableClient;
        private CloudTable configurationCloudTable;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamKbMappingStorageProvider"/> class.
        /// </summary>
        /// <param name="configuration">Application configuration settings.</param>
        public TeamKbMappingStorageProvider(IConfiguration configuration)
        {
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(configuration["StorageConnectionString"]));
        }

        /// <summary>
        /// This method is used to add or update team kb mapping.
        /// </summary>
        /// <param name="entity">table entity.</param>
        /// <returns>table entity inserted or updated.</returns>
        public async Task<TeamKbMapping> UpdateTeamKbMappingAsync(TeamKbMapping entity)
        {
            if (entity == null)
            {
                throw new ArgumentNullException("entity");
            }

            await this.EnsureInitializedAsync();
            TableOperation insertOrMergeOperation = TableOperation.InsertOrReplace(entity);
            TableResult result = await this.configurationCloudTable.ExecuteAsync(insertOrMergeOperation);
            return result.Result as TeamKbMapping;
        }

        /// <summary>
        /// This method gives the kb details based on team id.
        /// </summary>
        /// <param name="teamId">row key of the table.</param>
        /// <returns>table entity.</returns>
        public async Task<TeamKbMapping> GetKbMappingAsync(string teamId)
        {
            await this.EnsureInitializedAsync();
            TableOperation retrieveOperation = TableOperation.Retrieve<TeamKbMapping>(TeamKbMapping.TeamKbMappingPartitionKey, teamId);
            TableResult result = await this.configurationCloudTable.ExecuteAsync(retrieveOperation);
            return result?.Result as TeamKbMapping;
        }

        /// <summary>
        /// Get kb mapping details.
        /// </summary>
        /// <returns>TeamKbMappingResponse.</returns>
        public async Task<List<TeamKbMapping>> GetAllKbMappingsAsync()
        {
            await this.EnsureInitializedAsync();
            TableQuery<TeamKbMapping> query = new TableQuery<TeamKbMapping>();
            TableContinuationToken token = null;
            var entities = new List<TeamKbMapping>();
            do
            {
                var queryResult = await this.configurationCloudTable.ExecuteQuerySegmentedAsync(query, token);
                entities.AddRange(queryResult.Results);
                token = queryResult.ContinuationToken;
            }
            while (token != null);

            return entities;
        }

        /// <summary>
        /// Create crowdsourcer table if it doesnt exists.
        /// </summary>
        /// <param name="connectionString">storage account connection string.</param>
        /// <returns><see cref="Task"/> representing the asynchronous operation task which represents table is created if its not existing.</returns>
        private async Task<CloudTable> InitializeAsync(string connectionString)
        {
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
            this.cloudTableClient = storageAccount.CreateCloudTableClient();
            this.configurationCloudTable = this.cloudTableClient.GetTableReference(CrowdSourcerTableName);
            if (!await this.configurationCloudTable.ExistsAsync())
            {
                await this.configurationCloudTable.CreateIfNotExistsAsync();
            }

            return this.configurationCloudTable;
        }

        /// <summary>
        /// this method is called to ensure InitializeAsync method is called before any storage operation.
        /// </summary>
        /// <returns>Task.</returns>
        private async Task EnsureInitializedAsync()
        {
            await this.initializeTask.Value;
        }
    }
}
