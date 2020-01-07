// <copyright file="ISearchServiceDataProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Providers
{
    using System.Threading.Tasks;

    /// <summary>
    /// Search service data provider interface.
    /// </summary>
    public interface ISearchServiceDataProvider
    {
        /// <summary>
        /// This method downloads the knowledgebase and stores the json string to blob storage.
        /// </summary>
        /// <param name="kbId">knowledgebase id.</param>
        /// <returns>task.</returns>
        Task SetupAzureSearchDataAsync(string kbId);
    }
}