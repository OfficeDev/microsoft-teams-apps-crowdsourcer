// <copyright file="QnaServiceProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Common.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Web;
    using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker;
    using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker.Models;
    using Microsoft.Extensions.Configuration;

    /// <summary>
    /// qna maker service provider class.
    /// </summary>
    public class QnaServiceProvider : IQnaServiceProvider
    {
        /// <summary>
        /// The amount of time delay before checking the operation status details again.
        /// </summary>
        private const int OperationDelayMilliseconds = 5000;

        /// <summary>
        /// Retry count to check the operation status, if it is 'NotStarted' or 'Running'.
        /// </summary>
        private const int OperationRetryCount = 10;
        private const string DummyQuestion = "dummyquestion";
        private const string DummyAnswer = "dummyanswer";
        private const string DummyMetadataTeamId = "dummy";
        private const string Source = "Bot";
        private const string KbName = "teamscrowdsourcer";
        private const string Environment = "Prod";

        private readonly double scoreThreshold;
        private readonly IQnAMakerClient qnaMakerClient;
        private readonly IQnAMakerRuntimeClient qnaMakerRuntimeClient;
        private readonly IConfigurationStorageProvider configurationStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="QnaServiceProvider"/> class.
        /// </summary>
        /// <param name="configurationStorageProvider">storage provider.</param>
        /// <param name="configuration">configuration.</param>
        /// <param name="qnaMakerClient">qna service client.</param>
        /// <param name="qnaMakerRuntimeClient">qna service runtime client.</param>
        public QnaServiceProvider(IConfigurationStorageProvider configurationStorageProvider, IConfiguration configuration, IQnAMakerClient qnaMakerClient, IQnAMakerRuntimeClient qnaMakerRuntimeClient)
        {
            this.configurationStorageProvider = configurationStorageProvider;
            this.qnaMakerClient = qnaMakerClient;
            this.qnaMakerRuntimeClient = qnaMakerRuntimeClient;
            this.scoreThreshold = Convert.ToDouble(configuration["ScoreThreshold"]);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="QnaServiceProvider"/> class.
        /// </summary>
        /// <param name="configurationStorageProvider">storage provider.</param>
        /// <param name="qnaMakerClient">qna client.</param>
        public QnaServiceProvider(IConfigurationStorageProvider configurationStorageProvider, IQnAMakerClient qnaMakerClient)
        {
            this.configurationStorageProvider = configurationStorageProvider;
            this.qnaMakerClient = qnaMakerClient;
        }

        /// <summary>
        /// this method is used to add QnA pair in Kb.
        /// </summary>
        /// <param name="question">question text.</param>
        /// <param name="answer">answer text.</param>
        /// <param name="createdBy">created by user AAD object Id.</param>
        /// <param name="teamId">team id.</param>
        /// <param name="conversationId">conversation id.</param>
        /// <returns>task.</returns>
        public async Task AddQnaAsync(string question, string answer, string createdBy, string teamId, string conversationId)
        {
            var kb = await this.configurationStorageProvider.GetKbConfigAsync();

            // Update kb
            var updateKbOperation = await this.qnaMakerClient.Knowledgebase.UpdateAsync(kb.KbId, new UpdateKbOperationDTO
            {
                // Create JSON of changes.
                Add = new UpdateKbOperationDTOAdd
                {
                    QnaList = new List<QnADTO>
                    {
                         new QnADTO
                         {
                            Questions = new List<string> { question },
                            Answer = answer,
                            Source = Source,
                            Metadata = new List<MetadataDTO>()
                            {
                                new MetadataDTO() { Name = Constants.MetadataCreatedAt, Value = DateTime.UtcNow.Ticks.ToString("G", CultureInfo.InvariantCulture) },
                                new MetadataDTO() { Name = Constants.MetadataCreatedBy, Value = createdBy },
                                new MetadataDTO() { Name = Constants.MetadataTeamId, Value = HttpUtility.UrlEncode(teamId) },
                                new MetadataDTO() { Name = Constants.MetadataConversationId, Value = HttpUtility.UrlEncode(conversationId) },
                            },
                         },
                    },
                },
            });
        }

        /// <summary>
        /// this method is used to update Qna pair in Kb.
        /// </summary>
        /// <param name="questionId">question id.</param>
        /// <param name="answer">answer text.</param>
        /// <param name="updatedBy">updated by user AAD object Id.</param>
        /// <param name="updatedQuestion">updated question text.</param>
        /// <param name="question">original question text.</param>
        /// <param name="teamId">team id.</param>
        /// <returns>task.</returns>
        public async Task UpdateQnaAsync(int questionId, string answer, string updatedBy, string updatedQuestion, string question, string teamId)
        {
            var kb = await this.configurationStorageProvider.GetKbConfigAsync();
            var questions = default(UpdateQnaDTOQuestions);
            if (!string.IsNullOrEmpty(updatedQuestion))
            {
                questions = (updatedQuestion == question) ? null
                    : new UpdateQnaDTOQuestions()
                    {
                        Add = new List<string> { updatedQuestion },
                        Delete = new List<string> { question },
                    };
            }

            if (string.IsNullOrEmpty(answer))
            {
                answer = Constants.Unanswered;
            }

            // Update kb
            var updateKbOperation = await this.qnaMakerClient.Knowledgebase.UpdateAsync(kb.KbId, new UpdateKbOperationDTO
            {
                // Create JSON of changes.
                Update = new UpdateKbOperationDTOUpdate()
                {
                    QnaList = new List<UpdateQnaDTO>()
                    {
                        new UpdateQnaDTO()
                        {
                            Id = questionId,
                            Source = Source,
                            Answer = answer,
                            Questions = questions,
                            Metadata = new UpdateQnaDTOMetadata()
                            {
                                Add = new List<MetadataDTO>()
                                {
                                    new MetadataDTO() { Name = Constants.MetadataUpdatedAt, Value = DateTime.UtcNow.Ticks.ToString("G", CultureInfo.InvariantCulture) },
                                    new MetadataDTO() { Name = Constants.MetadataUpdatedBy, Value = updatedBy },
                                },
                            },
                        },
                    },
                },
            });
        }

        /// <summary>
        /// this method is used to delete Qna pair from KB.
        /// </summary>
        /// <param name="questionId">question id.</param>
        /// <param name="teamId">team id.</param>
        /// <returns>task.</returns>
        public async Task DeleteQnaAsync(int questionId, string teamId)
        {
            var kb = await this.configurationStorageProvider.GetKbConfigAsync();

            // to delete a qna based on id.
            var updateKbOperation = await this.qnaMakerClient.Knowledgebase.UpdateAsync(kb.KbId, new UpdateKbOperationDTO
            {
                // Create JSON of changes.
                Delete = new UpdateKbOperationDTODelete()
                {
                    Ids = new List<int?>() { questionId },
                },
            });
        }

        /// <summary>
        /// get answer from kb for a given question.
        /// </summary>
        /// <param name="isTest">prod or test.</param>
        /// <param name="question">question text.</param>
        /// <param name="teamId">team id.</param>
        /// <returns>qnaSearchResult response.</returns>
        public async Task<QnASearchResultList> GenerateAnswerAsync(bool isTest, string question, string teamId)
        {
            var kb = await this.configurationStorageProvider.GetKbConfigAsync();

            QnASearchResultList qnaSearchResult = await this.qnaMakerRuntimeClient.Runtime.GenerateAnswerAsync(kb?.KbId, new QueryDTO()
            {
                IsTest = isTest,
                Question = question,
                ScoreThreshold = this.scoreThreshold,
                StrictFilters = new List<MetadataDTO> { new MetadataDTO() { Name = Constants.MetadataTeamId, Value = HttpUtility.UrlEncode(teamId) } },
            });

            return qnaSearchResult;
        }

        /// <summary>
        /// this method can be used to publish the Kb.
        /// </summary>
        /// <param name="kbId">kb id.</param>
        /// <returns>task.</returns>
        public async Task PublishKnowledgebaseAsync(string kbId)
        {
            await this.qnaMakerClient.Knowledgebase.PublishAsync(kbId);
        }

        /// <summary>
        /// this method is used to create knowledgebase.
        /// </summary>
        /// <returns>kb id.</returns>
        public async Task<string> CreateKnowledgeBaseAsync()
        {
            // Adding one qna pair is mandatory for creating a knowledgebase. So giving dummy values.
            // We filter the qna pairs from knowledgebase based on teamid metadata value, so the dummy entry will be ignored.
            var createOp = await this.qnaMakerClient.Knowledgebase.CreateAsync(new CreateKbDTO()
            {
                Name = KbName,
                QnaList = new List<QnADTO>()
                {
                    new QnADTO()
                    {
                        Answer = DummyAnswer,
                        Questions = new List<string>() { DummyQuestion },
                        Source = Source,
                        Metadata = new List<MetadataDTO>()
                        {
                            new MetadataDTO()
                            {
                                Name = Constants.MetadataTeamId,
                                Value = DummyMetadataTeamId,
                            },
                        },
                    },
                },
            });

            createOp = await this.MonitorOperationAsync(createOp);
            return createOp?.ResourceLocation?.Split('/').Last();
        }

        /// <summary>
        /// Checks whether knowledge base need to be published.
        /// </summary>
        /// <param name="kbId">Knowledgebase ID.</param>
        /// <returns>boolean variable to publish or not.</returns>
        public async Task<bool> GetPublishStatusAsync(string kbId)
        {
            KnowledgebaseDTO knowledgebaseDetails = await this.qnaMakerClient.Knowledgebase.GetDetailsAsync(kbId);
            if (knowledgebaseDetails.LastChangedTimestamp != null && knowledgebaseDetails.LastPublishedTimestamp != null)
            {
                return Convert.ToDateTime(knowledgebaseDetails.LastChangedTimestamp) > Convert.ToDateTime(knowledgebaseDetails.LastPublishedTimestamp);
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// this method returns the downloaded kb documents.
        /// </summary>
        /// <param name="kbId">knowledgebase Id.</param>
        /// <returns>json string.</returns>
        public async Task<IEnumerable<QnADTO>> DownloadKnowledgebaseAsync(string kbId)
        {
            var qnaDocuments = await this.qnaMakerClient.Knowledgebase.DownloadAsync(kbId, environment: Environment);
            return qnaDocuments.QnaDocuments;
        }

        /// <summary>
        /// This method is used to delete a knowledgebase.
        /// </summary>
        /// <param name="kbId">knowledgebase Id.</param>
        /// <returns>task.</returns>
        public async Task DeleteKnowledgebaseAsync(string kbId)
        {
            await this.qnaMakerClient.Knowledgebase.DeleteAsync(kbId);
        }

        /// <summary>
        /// this method can be used to monitor any qnamaker operation.
        /// </summary>
        /// <param name="operation">operation details.</param>
        /// <returns>operation.</returns>
        private async Task<Operation> MonitorOperationAsync(Operation operation)
        {
            // Loop while operation is success
            for (int i = 0;
                i < OperationRetryCount && (operation.OperationState == OperationStateType.NotStarted || operation.OperationState == OperationStateType.Running);
                i++)
            {
                await Task.Delay(OperationDelayMilliseconds);
                operation = await this.qnaMakerClient.Operations.GetDetailsAsync(operation.OperationId);
            }

            if (operation.OperationState != OperationStateType.Succeeded)
            {
                throw new Exception($"Operation {operation.OperationId} failed. Error {operation.ErrorResponse.Error.Message}");
            }

            return operation;
        }
    }
}
