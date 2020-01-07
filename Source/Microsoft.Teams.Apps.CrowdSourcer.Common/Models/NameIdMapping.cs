// <copyright file="NameIdMapping.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Models
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Name Id Mapping table entity.
    /// </summary>
    public class NameIdMapping : TableEntity
    {
        /// <summary>
        /// Constant value used as a partition key for name id mapping storage.
        /// </summary>
        public const string NameIdMappingPartitionkey = "username";

        /// <summary>
        /// Initializes a new instance of the <see cref="NameIdMapping"/> class.
        /// </summary>
        public NameIdMapping()
        {
            this.PartitionKey = NameIdMappingPartitionkey; // constant
        }

        /// <summary>
        /// Gets or sets ObjectId.
        /// </summary>
        public string ObjectId
        {
            get
            {
                return this.RowKey;
            }

            set
            {
                this.RowKey = value;
            }
        }

        /// <summary>
        /// Gets or sets Name.
        /// </summary>
        public string Name { get; set; }
    }
}
