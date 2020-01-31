// <copyright file="IConfigurationStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Providers
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CrowdSourcer.Common.Models;

    /// <summary>
    /// configuration storage provider interface.
    /// </summary>
    public interface IConfigurationStorageProvider
    {
        /// <summary>
        /// get knowledge base Id from storage.
        /// </summary>
        /// <returns>Kb configuration.</returns>
        Task<KbConfiguration> GetKbConfigAsync();

        /// <summary>
        /// create knowledge base Id configuration in storage.
        /// </summary>
        /// <param name="entity">knowledge base configuration entity.</param>
        /// <returns>Knowledge base configuration details.</returns>
        Task<KbConfiguration> CreateKbConfigAsync(KbConfiguration entity);
    }
}