// <copyright file="Constants.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common
{
    /// <summary>
    /// constants.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Unanswered name.
        /// </summary>
        public const string Unanswered = "#$unanswered$#";

        /// <summary>
        /// Action name.
        /// </summary>
        public const string SubmitAddCommand = "submit/add";

        /// <summary>
        /// save command.
        /// </summary>
        public const string SaveCommand = "save";

        /// <summary>
        /// delete command.
        /// </summary>
        public const string DeleteCommand = "delete";

        /// <summary>
        /// no command.
        /// </summary>
        public const string NoCommand = "no";

        /// <summary>
        /// add command text.
        /// </summary>
        public const string AddCommand = "add question";

        /// <summary>
        /// qna metadata team id name.
        /// </summary>
        public const string MetadataTeamId = "teamid";

        /// <summary>
        /// qna metadata createdat name.
        /// </summary>
        public const string MetadataCreatedAt = "createdat";

        /// <summary>
        /// qna metadata createdby name.
        /// </summary>
        public const string MetadataCreatedBy = "createdby";

        /// <summary>
        /// qna metadata conversationid name.
        /// </summary>
        public const string MetadataConversationId = "conversationid";

        /// <summary>
        /// qna metadata updatedat name.
        /// </summary>
        public const string MetadataUpdatedAt = "updatedat";

        /// <summary>
        /// qna metadata updatedby name.
        /// </summary>
        public const string MetadataUpdatedBy = "updatedby";

        /// <summary>
        /// MessagingExtension recently created command id.
        /// </summary>
        public const string CreatedCommandId = "created";

        /// <summary>
        /// MessagingExtension recently edited command id.
        /// </summary>
        public const string EditedCommandId = "edited";

        /// <summary>
        /// MessagingExtension unanswered command id.
        /// </summary>
        public const string UnansweredCommandId = "unanswered";

        /// <summary>
        /// Blob Container Name.
        /// </summary>
        public const string StorageContainer = "crowdsourcer-search-container";

        /// <summary>
        /// Folder inside blob container.
        /// </summary>
        public const string FolderName = "crowdsourcer-metadata";
    }
}
