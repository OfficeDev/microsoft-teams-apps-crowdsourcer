// <copyright file="TeamKbMapping.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Models
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Teams kb mapping storage entity.
    /// </summary>
    public class TeamKbMapping : TableEntity
    {
        /// <summary>
        /// Constant value used as a partition key in storage.
        /// </summary>
        public const string TeamKbMappingPartitionKey = "msteams";

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamKbMapping"/> class.
        /// </summary>
        public TeamKbMapping()
        {
            this.PartitionKey = TeamKbMappingPartitionKey;
        }

        /// <summary>
        /// Gets or sets TeamId.
        /// </summary>
        public string TeamId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets KbId.
        /// </summary>
        public string KbId { get; set; }
    }
}
