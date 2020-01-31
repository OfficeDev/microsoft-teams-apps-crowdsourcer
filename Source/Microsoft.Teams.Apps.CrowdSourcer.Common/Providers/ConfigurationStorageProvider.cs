// <copyright file="ConfigurationStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Providers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.CrowdSourcer.Common.Models;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Configuration StorageProvider.
    /// </summary>
    public class ConfigurationStorageProvider : IConfigurationStorageProvider
    {
        private const string CrowdSourcerTableName = "crowdsourcerconfig";
        private readonly Lazy<Task> initializeTask;
        private CloudTableClient cloudTableClient;
        private CloudTable configurationCloudTable;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationStorageProvider"/> class.
        /// </summary>
        /// <param name="configuration">config settings.</param>
        public ConfigurationStorageProvider(IConfiguration configuration)
        {
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(configuration["StorageConnectionString"]));
        }

        /// <summary>
        /// get knowledge base Id from storage.
        /// </summary>
        /// <returns>Kb Configuration.</returns>
        public async Task<KbConfiguration> GetKbConfigAsync()
        {
            await this.EnsureInitializedAsync();
            TableOperation retrieveOperation = TableOperation.Retrieve<KbConfiguration>(KbConfiguration.KbConfigurationPartitionKey, KbConfiguration.KbConfigurationRowKey);
            TableResult result = await this.configurationCloudTable.ExecuteAsync(retrieveOperation);
            return result?.Result as KbConfiguration;
        }

        /// <summary>
        /// create knowledge base Id configuration in storage.
        /// </summary>
        /// <param name="entity">KbConfiguration entity.</param>
        /// <returns>Knowledge base configuration details.</returns>
        public async Task<KbConfiguration> CreateKbConfigAsync(KbConfiguration entity)
        {
            if (entity == null)
            {
                throw new ArgumentNullException("entity");
            }

            await this.EnsureInitializedAsync();
            TableOperation insertOrMergeOperation = TableOperation.InsertOrReplace(entity);
            TableResult result = await this.configurationCloudTable.ExecuteAsync(insertOrMergeOperation);
            return result.Result as KbConfiguration;
        }

        /// <summary>
        /// Create crowdsourcerconfig table if it does not exist.
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
