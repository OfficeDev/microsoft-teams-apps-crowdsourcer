// <copyright file="SearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Azure.Search;
    using Microsoft.Azure.Search.Models;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.CrowdSourcer.Models;

    /// <summary>
    /// azure blob search service class.
    /// </summary>
    public class SearchService : ISearchService
    {
        private const string IndexName = "teams-crowdsourcer-index";
        private const string DataSourceName = "crowdsourcer-datasource";
        private const int TopCount = 50;

        private readonly IConfiguration configuration;
        private readonly SearchIndexClient searchIndexClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="SearchService"/> class.
        /// </summary>
        /// <param name="configuration">configuration settings.</param>
        public SearchService(IConfiguration configuration)
        {
            this.configuration = configuration;
            this.searchIndexClient = new SearchIndexClient(
                configuration["SearchServiceName"],
                IndexName,
                new SearchCredentials(configuration["SearchServiceKey"]));
        }

        /// <summary>
        /// This method gives search result(Konwledgebase QnA pairs) based on teamId and search query for specific messaging extension command Id.
        /// </summary>
        /// <param name="searchQuery">searchQuery.</param>
        /// <param name="commandId">commandId.</param>
        /// <param name="teamId">team id.</param>
        /// <returns>search result list.</returns>
        public async Task<IList<AzureSearchEntity>> GetAzureSearchEntitiesAsync(string searchQuery, string commandId, string teamId)
        {
            IList<AzureSearchEntity> qnaPairs = new List<AzureSearchEntity>();
            SearchParameters searchParameters = default(SearchParameters);
            string searchFilter = string.Empty;
            if (string.IsNullOrWhiteSpace(searchQuery))
            {
                switch (commandId)
                {
                    case Constants.CreatedCommandId:
                        searchParameters = new SearchParameters()
                        {
                            OrderBy = new[] { "createddate desc" },
                            Top = TopCount,
                            Filter = $"answer ne '{Constants.Unanswered}' and teamid eq '{teamId}'",
                        };
                        break;
                    case Constants.EditedCommandId:
                        searchParameters = new SearchParameters()
                        {
                            OrderBy = new[] { "updateddate desc" },
                            Top = TopCount,
                            Filter = $"answer ne '{Constants.Unanswered}' and teamid eq '{teamId}'",
                        };
                        break;
                    case Constants.UnansweredCommandId:
                        searchParameters = new SearchParameters()
                        {
                            OrderBy = new[] { "createddate desc" },
                            Top = TopCount,
                            Filter = $"answer eq '{Constants.Unanswered}' and teamid eq '{teamId}'",
                        };
                        break;
                    default:
                        break;
                }
            }
            else
            {
                searchParameters = new SearchParameters()
                {
                    OrderBy = new[] { "search.score() desc" },
                    Top = TopCount,
                    Filter = $"teamid eq '{teamId}'",
                };
                searchFilter = searchQuery;
            }

            var documents = await this.searchIndexClient.Documents.SearchAsync<AzureSearchEntity>(searchFilter + "*", searchParameters);
            if (documents != null)
            {
                foreach (SearchResult<AzureSearchEntity> searchResult in documents.Results)
                {
                    qnaPairs.Add(searchResult.Document);
                }
            }

            return qnaPairs;
        }

        /// <summary>
        /// Creates Index, Data Source and Indexer for search service.
        /// </summary>
        /// <returns>task.</returns>
        public async Task InitializeSearchServiceDependency()
        {
            ISearchServiceClient searchClient = this.CreateSearchServiceClient();
            await this.CreateSearchIndexAsync(searchClient);
            await this.CreateDataSourceAsync(searchClient);
            await this.CreateOrRunIndexerAsync(searchClient);
        }

        /// <summary>
        /// this methoda creates search service client.
        /// </summary>
        /// <returns>search client.</returns>
        private SearchServiceClient CreateSearchServiceClient()
        {
            SearchServiceClient serviceClient = new SearchServiceClient(this.configuration["SearchServiceName"], new SearchCredentials(this.configuration["SearchServiceKey"]));
            return serviceClient;
        }

        /// <summary>
        /// Creates new SearchIndex with INDEX_NAME provided, if already exists then delete the index and create again.
        /// </summary>
        /// <param name="searchClient">search client.</param>
        private async Task CreateSearchIndexAsync(ISearchServiceClient searchClient)
        {
            if (await searchClient.Indexes.ExistsAsync(IndexName))
            {
                await searchClient.Indexes.DeleteAsync(IndexName);
            }

            var definition = new Index()
            {
                Name = IndexName,
                Fields = FieldBuilder.BuildForType<AzureSearchEntity>(),
            };
            await searchClient.Indexes.CreateAsync(definition);
        }

        /// <summary>
        ///  Creates new DataSource with DATASOURCE_NAME provided, if already exists no change happen.
        /// </summary>
        /// <param name="searchClient">search client.</param>
        private async Task CreateDataSourceAsync(ISearchServiceClient searchClient)
        {
            if (await searchClient.DataSources.ExistsAsync(DataSourceName))
            {
                return;
            }

            var dataSourceConfig = new DataSource()
            {
                Name = DataSourceName,
                Container = new DataContainer(Constants.StorageContainer, Constants.FolderName),
                Credentials = new DataSourceCredentials(this.configuration["StorageConnectionString"]),
                Type = DataSourceType.AzureBlob,
            };

            await searchClient.DataSources.CreateAsync(dataSourceConfig);
        }

        /// <summary>
        /// Creates new Indexer or run if it already exists.
        /// </summary>
        /// <param name="searchClient">search client.</param>
        private async Task CreateOrRunIndexerAsync(ISearchServiceClient searchClient)
        {
            if (await searchClient.Indexers.ExistsAsync(IndexName))
            {
                await searchClient.Indexers.RunAsync(IndexName);
                return;
            }

            var parseConfig = new Dictionary<string, object>();
            parseConfig.Add("parsingMode", "jsonArray");

            var indexerConfig = new Indexer()
            {
                Name = IndexName,
                DataSourceName = DataSourceName,
                TargetIndexName = IndexName,
                Parameters = new IndexingParameters()
                {
                    Configuration = parseConfig,
                },
            };
            await searchClient.Indexers.CreateAsync(indexerConfig);
        }
    }
}
