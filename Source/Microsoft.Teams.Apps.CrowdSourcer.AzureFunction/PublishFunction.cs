// <copyright file="PublishFunction.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.PublishFunction
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CrowdSourcer.Common.Models;
    using Microsoft.Teams.Apps.CrowdSourcer.Common.Providers;

    /// <summary>
    /// Azure Function to publish knowledge bases if modified.
    /// </summary>
    public class PublishFunction
    {
        private readonly ITeamKbMappingStorageProvider storageProvider;
        private readonly IQnaServiceProvider qnaServiceProvider;
        private readonly ISearchServiceDataProvider searchServiceDataProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="PublishFunction"/> class.
        /// </summary>
        /// <param name="storageProvider">storage provider.</param>
        /// <param name="qnaServiceProvider">qna service provider.</param>
        /// <param name="searchServiceDataProvider">search service data provider.</param>
        public PublishFunction(ITeamKbMappingStorageProvider storageProvider, IQnaServiceProvider qnaServiceProvider, ISearchServiceDataProvider searchServiceDataProvider)
        {
            this.storageProvider = storageProvider;
            this.qnaServiceProvider = qnaServiceProvider;
            this.searchServiceDataProvider = searchServiceDataProvider;
        }

        /// <summary>
        /// Function to get list of KB and publish KB.
        /// </summary>
        /// <param name="myTimer">Publish frequency.</param>
        /// <param name="log">Log.</param>
        /// <returns>A <see cref="Task"/> representing the result of the asynchronous operation.</returns>
        [FunctionName("PublishFunction")]
        public async Task Run([TimerTrigger("0 */15 * * * *")]TimerInfo myTimer, ILogger log)
        {
            try
            {
                IEnumerable<TeamKbMapping> teamKbMappings = await this.storageProvider.GetAllKbMappingsAsync();
                var knowledgebaseList = teamKbMappings.Select(x => x.KbId).Distinct();
                foreach (string kb in knowledgebaseList)
                {
                    bool toBePublished = await this.qnaServiceProvider.GetPublishStatusAsync(kb);
                    log.LogInformation("To be Published - " + toBePublished);
                    log.LogInformation("KbId - " + kb);
                    log.LogInformation("QnAMakerApiUrl - " + Environment.GetEnvironmentVariable("QnAMakerApiUrl"));
                    if (toBePublished)
                    {
                        await this.qnaServiceProvider.PublishKnowledgebaseAsync(kb);
                    }

                    await this.searchServiceDataProvider.SetupAzureSearchDataAsync(kb);
                }
            }
            catch (Exception ex)
            {
                log.LogError("Error: " + ex.Message); // Exception logging.
                log.LogError(ex.ToString());
            }
        }
    }
}
