// <copyright file="AdaptiveSubmitActionData.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Models
{
    using Microsoft.Bot.Schema;
    using Newtonsoft.Json;

    /// <summary>
    /// Adaptive Card Action class.
    /// </summary>
    public class AdaptiveSubmitActionData
    {
        /// <summary>
        /// Gets or sets Msteams object.
        /// </summary>
        [JsonProperty("msteams")]
        public CardAction MsTeams { get; set; }

        /// <summary>
        ///  Gets or sets details.
        /// </summary>
        [JsonProperty("details")]
        public Details Details { get; set; }

        /// <summary>
        /// Gets or sets Updated question.
        /// </summary>
        [JsonProperty("question")]
        public string Question { get; set; }

        /// <summary>
        /// Gets or sets Answer.
        /// </summary>
        [JsonProperty("answer")]
        public string Answer { get; set; }
    }
}