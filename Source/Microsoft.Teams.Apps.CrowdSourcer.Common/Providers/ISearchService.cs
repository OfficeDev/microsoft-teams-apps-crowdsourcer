// <copyright file="ISearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Providers
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CrowdSourcer.Models;

    /// <summary>
    /// azure blob search service interface.
    /// </summary>
    public interface ISearchService
    {
        /// <summary>
        /// This method gives search result(Konwledgebase QnA pairs) based on teamId and search query for specific messaging extension command Id.
        /// </summary>
        /// <param name="searchQuery">searchQuery.</param>
        /// <param name="commandId">messaging extension commandId.</param>
        /// <param name="teamId">team Id.</param>
        /// <returns>search result list.</returns>
        Task<IList<AzureSearchEntity>> GetAzureSearchEntitiesAsync(string searchQuery, string commandId, string teamId);

        /// <summary>
        /// Creates Index, Data Source and Indexer for search service.
        /// </summary>
        /// <returns>task.</returns>
        Task InitializeSearchServiceDependency();
    }
}