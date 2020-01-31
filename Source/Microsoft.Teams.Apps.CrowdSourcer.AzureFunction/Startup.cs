// <copyright file="Startup.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

using System;
using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Teams.Apps.CrowdSourcer.AzureFunction;
using Microsoft.Teams.Apps.CrowdSourcer.Common.Providers;

[assembly: WebJobsStartup(typeof(Startup))]

namespace Microsoft.Teams.Apps.CrowdSourcer.AzureFunction
{
    /// <summary>
    /// Azure function Startup Class.
    /// </summary>
    public class Startup : IWebJobsStartup
    {
        /// <summary>
        /// Application startup configuration.
        /// </summary>
        /// <param name="builder">webjobs builder.</param>
        public void Configure(IWebJobsBuilder builder)
        {
            builder.Services.AddSingleton<IConfigurationStorageProvider, ConfigurationStorageProvider>();
            IQnAMakerClient qnaMakerClient = new QnAMakerClient(new ApiKeyServiceClientCredentials(Environment.GetEnvironmentVariable("QnAMakerSubscriptionKey"))) { Endpoint = Environment.GetEnvironmentVariable("QnAMakerApiUrl") };
            builder.Services.AddSingleton<IQnaServiceProvider>((provider) => new QnaServiceProvider(
                provider.GetRequiredService<IConfigurationStorageProvider>(),
                qnaMakerClient));
            builder.Services.AddSingleton<ISearchServiceDataProvider, SearchServiceDataProvider>();
            builder.Services.AddSingleton<ISearchService, SearchService>();
        }
    }
}
