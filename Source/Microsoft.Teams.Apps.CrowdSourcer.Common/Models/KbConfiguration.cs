// <copyright file="KbConfiguration.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Models
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Knowledgebase configuration storage entity.
    /// </summary>
    public class KbConfiguration : TableEntity
    {
        /// <summary>
        /// Constant value used as a partition key in storage.
        /// </summary>
        public const string KbConfigurationPartitionKey = "msteams";

        /// <summary>
        /// Constant value used as a row key in storage.
        /// </summary>
        public const string KbConfigurationRowKey = "knowledgebaseId";

        /// <summary>
        /// Initializes a new instance of the <see cref="KbConfiguration"/> class.
        /// </summary>
        public KbConfiguration()
        {
            this.PartitionKey = KbConfigurationPartitionKey;
            this.RowKey = KbConfigurationRowKey;
        }

        /// <summary>
        /// Gets or sets KbId.
        /// </summary>
        public string KbId { get; set; }
    }
}
