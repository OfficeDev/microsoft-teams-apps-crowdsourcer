// <copyright file="ObjectIdToNameMapper.cs" company="Microsoft">
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
    /// Name Id mapping storage provider class.
    /// </summary>
    public class ObjectIdToNameMapper : IObjectIdToNameMapper
    {
        private const string NameIdMappingTableName = "crowdsourcernames";
        private readonly Lazy<Task> initializeTask;
        private CloudTableClient cloudTableClient;
        private CloudTable configurationCloudTable;

        /// <summary>
        /// Initializes a new instance of the <see cref="ObjectIdToNameMapper"/> class.
        /// </summary>
        /// <param name="configuration">configuration settings.</param>
        public ObjectIdToNameMapper(IConfiguration configuration)
        {
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(configuration["StorageConnectionString"]));
        }

        /// <summary>
        /// This method is used to add or update the aad object id and name mapping.
        /// </summary>
        /// <param name="entity">table entity.</param>
        /// <returns>table entity inserted or updated.</returns>
        public async Task<NameIdMapping> UpdateNameMappingAsync(NameIdMapping entity)
        {
            if (entity == null)
            {
                throw new ArgumentNullException("entity");
            }

            await this.EnsureInitializedAsync();
            TableOperation insertOrMergeOperation = TableOperation.InsertOrReplace(entity);
            TableResult result = await this.configurationCloudTable.ExecuteAsync(insertOrMergeOperation);
            return result.Result as NameIdMapping;
        }

        /// <summary>
        /// This method is used to get name based on aad object id.
        /// </summary>
        /// <param name="objectId">aad object Id.</param>
        /// <returns>name.</returns>
        public async Task<string> GetNameAsync(string objectId)
        {
            await this.EnsureInitializedAsync();
            TableOperation retrieveOperation = TableOperation.Retrieve<NameIdMapping>(NameIdMapping.NameIdMappingPartitionkey, objectId);
            TableResult result = await this.configurationCloudTable.ExecuteAsync(retrieveOperation);
            var mappingEntity = result?.Result as NameIdMapping;
            return mappingEntity.Name;
        }

        /// <summary>
        /// Create name mapping table if it doesnt exists.
        /// </summary>
        /// <param name="connectionString">storage account connection string.</param>
        /// <returns><see cref="Task"/> representing the asynchronous operation task which represents table is created if its not existing.</returns>
        private async Task<CloudTable> InitializeAsync(string connectionString)
        {
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
            this.cloudTableClient = storageAccount.CreateCloudTableClient();
            this.configurationCloudTable = this.cloudTableClient.GetTableReference(NameIdMappingTableName);
            if (!await this.configurationCloudTable.ExistsAsync())
            {
                await this.configurationCloudTable.CreateIfNotExistsAsync();
            }

            return this.configurationCloudTable;
        }

        /// <summary>
        /// This method is called to ensure InitializeAsync method is called before any storage operation.
        /// </summary>
        /// <returns>Task.</returns>
        private async Task EnsureInitializedAsync()
        {
            await this.initializeTask.Value;
        }
    }
}
