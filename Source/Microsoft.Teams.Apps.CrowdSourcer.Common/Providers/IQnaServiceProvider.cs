// <copyright file="IQnaServiceProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Providers
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker.Models;

    /// <summary>
    /// qna maker service provider interface.
    /// </summary>
    public interface IQnaServiceProvider
    {
        /// <summary>
        /// this method is used to add QnA pair in Kb.
        /// </summary>
        /// <param name="question">question text.</param>
        /// <param name="answer">answer text.</param>
        /// <param name="createdBy">created by user.</param>
        /// <param name="teamId">team id.</param>
        /// <param name="conversationId">conversation id.</param>
        /// <returns>task.</returns>
        Task AddQnaAsync(string question, string answer, string createdBy, string teamId, string conversationId);

        /// <summary>
        /// this method is used to create knowledgebase.
        /// </summary>
        /// <returns>kb id.</returns>
        Task<string> CreateKnowledgeBaseAsync();

        /// <summary>
        /// this method is used to delete Qna pair from KB.
        /// </summary>
        /// <param name="questionId">question id.</param>
        /// <param name="teamId">team id.</param>
        /// <returns>delete response.</returns>
        Task DeleteQnaAsync(int questionId, string teamId);

        /// <summary>
        /// get answer from kb for a given question.
        /// </summary>
        /// <param name="isTest">prod or test.</param>
        /// <param name="question">question text.</param>
        /// <param name="teamId">team id.</param>
        /// <returns>qnaSearchResult response.</returns>
        Task<QnASearchResultList> GenerateAnswerAsync(bool isTest, string question, string teamId);

        /// <summary>
        /// this method can be used to publish the Kb.
        /// </summary>
        /// <param name="kbId">kb id.</param>
        /// <returns>task.</returns>
        Task PublishKnowledgebaseAsync(string kbId);

        /// <summary>
        /// this method is used to update Qna pair in Kb.
        /// </summary>
        /// <param name="questionId">question id.</param>
        /// <param name="answer">answer text.</param>
        /// <param name="updatedBy">updated by user.</param>
        /// <param name="updatedQuestion">updated question text.</param>
        /// <param name="question">original question text.</param>
        /// <param name="teamId">team id.</param>
        /// <returns>task.</returns>
        Task UpdateQnaAsync(int questionId, string answer, string updatedBy, string updatedQuestion, string question, string teamId);

        /// <summary>
        /// Checks whether knowledge base need to be published.
        /// </summary>
        /// <param name="kbId">Knowledgebase ID.</param>
        /// <returns>boolean variable to publish or not.</returns>
        Task<bool> GetPublishStatusAsync(string kbId);

        /// <summary>
        /// this method returns the downloaded kb documents.
        /// </summary>
        /// <param name="kbId">knowledgebase Id.</param>
        /// <returns>json string.</returns>
        Task<IEnumerable<QnADTO>> DownloadKnowledgebaseAsync(string kbId);

        /// <summary>
        /// This method is used to delete a knowledgebase.
        /// </summary>
        /// <param name="kbId">knowledgebase Id.</param>
        /// <returns>task.</returns>
        Task DeleteKnowledgebaseAsync(string kbId);
    }
}
