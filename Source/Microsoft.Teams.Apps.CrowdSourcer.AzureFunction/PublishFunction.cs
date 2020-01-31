// <copyright file="PublishFunction.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.PublishFunction
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CrowdSourcer.Common.Models;
    using Microsoft.Teams.Apps.CrowdSourcer.Common.Providers;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Azure Function to create and publish knowledge bases.
    /// </summary>
    public class PublishFunction
    {
        /// <summary>
        /// Retry policy with jitter, Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </summary>
        private static RetryPolicy retryPolicy = Policy.Handle<Exception>()
            .WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(1000), 2));

        private readonly IConfigurationStorageProvider configurationStorageProvider;
        private readonly IQnaServiceProvider qnaServiceProvider;
        private readonly ISearchServiceDataProvider searchServiceDataProvider;
        private readonly ISearchService searchService;

        /// <summary>
        /// Initializes a new instance of the <see cref="PublishFunction"/> class.
        /// </summary>
        /// <param name="searchServiceDataProvider">search service data provider.</param>
        /// <param name="qnaServiceProvider">qna service provider.</param>
        /// <param name="configurationStorageProvider">configuration storage provider.</param>
        /// <param name="searchService">search service.</param>
        public PublishFunction(IConfigurationStorageProvider configurationStorageProvider, IQnaServiceProvider qnaServiceProvider, ISearchServiceDataProvider searchServiceDataProvider, ISearchService searchService)
        {
            this.configurationStorageProvider = configurationStorageProvider;
            this.qnaServiceProvider = qnaServiceProvider;
            this.searchServiceDataProvider = searchServiceDataProvider;
            this.searchService = searchService;
        }

        /// <summary>
        /// Function to get knowledge base Id, create knowledge base if not exist and publish knowledge base. Also setup the azure search service dependencies.
        /// </summary>
        /// <param name="myTimer">Publish frequency.</param>
        /// <param name="log">Log.</param>
        /// <returns>A <see cref="Task"/> representing the result of the asynchronous operation.</returns>
        [FunctionName("PublishFunction")]
        public async Task Run([TimerTrigger("0 */15 * * * *", RunOnStartup = true)]TimerInfo myTimer, ILogger log)
        {
            try
            {
                KbConfiguration kbConfiguration = await this.configurationStorageProvider.GetKbConfigAsync();
                if (kbConfiguration == null)
                {
                    log.LogInformation("Creating knowledge base");
                    string kbId = await this.qnaServiceProvider.CreateKnowledgeBaseAsync();

                    log.LogInformation("Publishing knowledge base");
                    await this.qnaServiceProvider.PublishKnowledgebaseAsync(kbId);

                    KbConfiguration kbConfigurationEntity = new KbConfiguration()
                    {
                        KbId = kbId,
                    };

                    log.LogInformation("Storing knowledgebase Id in storage " + kbId);
                    try
                    {
                        await retryPolicy.ExecuteAsync(async () =>
                            {
                                kbConfiguration = await this.configurationStorageProvider.CreateKbConfigAsync(kbConfigurationEntity);
                            });
                    }
                    catch (Exception ex)
                    {
                        log.LogError("Error: " + ex.ToString());
                        log.LogWarning("Failed to store knowledgebase Id in storage. Deleting " + kbId);
                        await this.qnaServiceProvider.DeleteKnowledgebaseAsync(kbId);
                        return;
                    }

                    log.LogInformation("Setup Azure Search Data");
                    await this.searchServiceDataProvider.SetupAzureSearchDataAsync(kbId);
                    log.LogInformation("Update Azure Search service");
                    await this.searchService.InitializeSearchServiceDependency();
                }
                else
                {
                    bool toBePublished = await this.qnaServiceProvider.GetPublishStatusAsync(kbConfiguration.KbId);
                    log.LogInformation("To be Published - " + toBePublished);
                    log.LogInformation("KbId - " + kbConfiguration.KbId);

                    if (toBePublished)
                    {
                        log.LogInformation("Publishing knowledgebase");
                        await this.qnaServiceProvider.PublishKnowledgebaseAsync(kbConfiguration.KbId);
                        log.LogInformation("Setup Azure Search Data");
                        await this.searchServiceDataProvider.SetupAzureSearchDataAsync(kbConfiguration.KbId);
                        log.LogInformation("Update Azure Search service");
                        await this.searchService.InitializeSearchServiceDependency();
                    }
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
