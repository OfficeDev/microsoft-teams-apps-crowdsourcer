// <copyright file="ITeamKbMappingStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Providers
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CrowdSourcer.Common.Models;

    /// <summary>
    /// Team kb mapping storage provider interface.
    /// </summary>
    public interface ITeamKbMappingStorageProvider
    {
        /// <summary>
        /// Insert or merge the table entity.
        /// </summary>
        /// <param name="entity">TeamKbMapping table entity.</param>
        /// <returns>TeamKbMapping table entity to be inserted or updated.</returns>
        Task<TeamKbMapping> UpdateTeamKbMappingAsync(TeamKbMapping entity);

        /// <summary>
        /// This method gives the kb details based on team id.
        /// </summary>
        /// <param name="teamId">team id.</param>
        /// <returns>entity.</returns>
        Task<TeamKbMapping> GetKbMappingAsync(string teamId);

        /// <summary>
        /// Get kb mapping details.
        /// </summary>
        /// <returns>TeamKbMappingResponse.</returns>
        Task<List<TeamKbMapping>> GetAllKbMappingsAsync();
    }
}
