// <copyright file="IObjectIdToNameMapper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Providers
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CrowdSourcer.Common.Models;

    /// <summary>
    /// Object Id name mapping storage provider interface.
    /// </summary>
    public interface IObjectIdToNameMapper
    {
        /// <summary>
        /// method is used to get name based on aad object id.
        /// </summary>
        /// <param name="objectId">row key of the table.</param>
        /// <returns>name.</returns>
        Task<string> GetNameAsync(string objectId);

        /// <summary>
        /// This method is used to add or update the aad object id and name mapping.
        /// </summary>
        /// <param name="entity">table entity.</param>
        /// <returns>table entity inserted or updated.</returns>
        Task<NameIdMapping> UpdateNameMappingAsync(NameIdMapping entity);
    }
}