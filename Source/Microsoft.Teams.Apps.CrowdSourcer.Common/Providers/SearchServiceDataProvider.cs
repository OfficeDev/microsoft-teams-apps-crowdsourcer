// <copyright file="SearchServiceDataProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker.Models;
    using Microsoft.Teams.Apps.CrowdSourcer.Models;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.Blob;
    using Newtonsoft.Json;

    /// <summary>
    /// azure search service blob storage data provider.
    /// </summary>
    public class SearchServiceDataProvider : ISearchServiceDataProvider
    {
        private readonly IQnaServiceProvider qnaServiceProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="SearchServiceDataProvider"/> class.
        /// </summary>
        /// <param name="qnaServiceProvider">qna ServiceProvider.</param>
        public SearchServiceDataProvider(IQnaServiceProvider qnaServiceProvider)
        {
            this.qnaServiceProvider = qnaServiceProvider;
        }

        /// <summary>
        /// this method downloads the knowledgebase and stores the json string to blob storage.
        /// </summary>
        /// <param name="kbId">knowledgebase id.</param>
        /// <returns>task.</returns>
        public async Task SetupAzureSearchDataAsync(string kbId)
        {
            IEnumerable<QnADTO> qnaDocuments = await this.qnaServiceProvider.DownloadKnowledgebaseAsync(kbId);
            string azureJson = this.GenerateFormattedJson(qnaDocuments);
            await this.AddDatatoBlobStorage(azureJson);
        }

        /// <summary>
        /// Function to convert input JSON to align with Schema Definition.
        /// </summary>
        /// <param name="qnaDocuments">qna documents.</param>
        /// <returns>create json format for search.</returns>
        private string GenerateFormattedJson(IEnumerable<QnADTO> qnaDocuments)
        {
            List<AzureSearchEntity> searchEntityList = new List<AzureSearchEntity>();
            foreach (var item in qnaDocuments)
            {
                var createdDate = item.Metadata.Where(prop => prop.Name == Constants.MetadataCreatedAt).FirstOrDefault();
                var updatedDate = item.Metadata.Where(prop => prop.Name == Constants.MetadataUpdatedAt).FirstOrDefault();
                var teamId = item.Metadata.Where(prop => prop.Name == Constants.MetadataTeamId).First().Value.ToString();

                searchEntityList.Add(
                        new AzureSearchEntity()
                        {
                            Id = item.Id.ToString(),
                            Source = item.Source,
                            Questions = item.Questions,
                            Answer = item.Answer,
                            CreatedDate = createdDate != null ? new DateTimeOffset(new DateTime(Convert.ToInt64(createdDate.Value))) : default(DateTimeOffset),
                            UpdatedDate = updatedDate != null ? new DateTimeOffset(new DateTime(Convert.ToInt64(updatedDate.Value))) : default(DateTimeOffset),
                            TeamId = teamId,
                            Metadata = item.Metadata,
                        });
            }

            return JsonConvert.SerializeObject(searchEntityList);
        }

        /// <summary>
        /// This method is used to store json to blob storage.
        /// </summary>
        /// <param name="jsonData">knowledgebase jsonData string.</param>
        /// <returns>task.</returns>
        private async Task AddDatatoBlobStorage(string jsonData)
        {
            // Retrieve storage account from connection string.
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(Environment.GetEnvironmentVariable("AzureWebJobsStorage"));

            // Create the blob client.
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();

            // Retrieve a reference to a container.
            CloudBlobContainer container = blobClient.GetContainerReference(Constants.StorageContainer);

            // Create the container if it doesn't already exist.
            var result = await container.CreateIfNotExistsAsync();

            // Retrieve reference to blob.
            CloudBlockBlob blockBlob = container.GetBlockBlobReference(Constants.FolderName + "/teamscrowdsourcer.json");
            blockBlob.Properties.ContentType = "application/json";

            // Upload JSON to blob storage.
            await blockBlob.UploadTextAsync(jsonData);
        }
    }
}
