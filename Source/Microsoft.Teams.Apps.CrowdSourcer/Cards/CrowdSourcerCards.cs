// <copyright file="CrowdSourcerCards.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CrowdSourcer.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Web;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.CrowdSourcer.Common;
    using Microsoft.Teams.Apps.CrowdSourcer.Common.Providers;
    using Microsoft.Teams.Apps.CrowdSourcer.Models;
    using Microsoft.Teams.Apps.CrowdSourcer.Resources;

    /// <summary>
    /// create crowdsourcer cards.
    /// </summary>
    public class CrowdSourcerCards
    {
        private const int TruncateThresholdLength = 50;
        private const int QuestionMaxInputLength = 100;
        private const int AnswerMaxInputLength = 500;
        private readonly IObjectIdToNameMapper nameMappingStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="CrowdSourcerCards"/> class.
        /// </summary>
        /// <param name="nameMappingStorageProvider">name id mapping storage provider.</param>
        public CrowdSourcerCards(IObjectIdToNameMapper nameMappingStorageProvider)
        {
            this.nameMappingStorageProvider = nameMappingStorageProvider;
        }

        /// <summary>
        /// returns the messaging extension attachment of all answers.
        /// </summary>
        /// <param name="qnaDocuments">all qnaDocuments.</param>
        /// <returns>returns the list of all answered questions.</returns>
        public async Task<List<MessagingExtensionAttachment>> MessagingExtensionCardListAsync(IList<AzureSearchEntity> qnaDocuments)
        {
            var messagingExtensionAttachments = new List<MessagingExtensionAttachment>();

            foreach (var qnaDoc in qnaDocuments)
            {
                DateTime createdAt = (qnaDoc.Metadata.Count > 1) ? new DateTime(long.Parse(qnaDoc.Metadata.Where(s => s.Name == Constants.MetadataCreatedAt).First().Value)) : default;
                string dateString = default;
                string createdBy = default;
                string conversationId = default;

                if (qnaDoc.Metadata?.Count > 1)
                {
                    string objectId = qnaDoc.Metadata.Where(x => x.Name == Constants.MetadataCreatedBy).First().Value;
                    createdBy = await this.nameMappingStorageProvider.GetNameAsync(objectId);
                    conversationId = qnaDoc.Metadata.Where(s => s.Name == Constants.MetadataConversationId).First().Value;
                    dateString = string.Format(Strings.DateFormat, "{{DATE(" + createdAt.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'") + ", SHORT)}}", "{{TIME(" + createdAt.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'") + ")}}");
                }

                string answer = qnaDoc.Answer.Equals(Constants.Unanswered) ? string.Empty : qnaDoc.Answer;

                var card = new AdaptiveCard("1.0")
                {
                    Body = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Text = $"**{Strings.QuestionTitle}**: {qnaDoc.Questions[0]}",
                            Size = AdaptiveTextSize.Default,
                            Wrap = true,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = string.IsNullOrEmpty(answer) ? string.Empty : $"**{Strings.AnswerTitle}**: {answer}",
                            Size = AdaptiveTextSize.Default,
                            Wrap = true,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = $"{createdBy} | {dateString}",
                            Wrap = true,
                        },
                    },
                };

                // If conversation id is not "#" whose url decode value is "%23" then create go to thread url.
                if (!conversationId.Equals("%23"))
                {
                    conversationId = HttpUtility.UrlDecode(conversationId);
                    string[] threadAndMessageId = conversationId.Split(";");
                    var threadId = threadAndMessageId[0];
                    var messageId = threadAndMessageId[1].Split("=")[1];

                    card.Actions.Add(
                        new AdaptiveOpenUrlAction()
                        {
                            Title = Strings.GoToThread,
                            Url = new Uri($"https://teams.microsoft.com/l/message/{threadId}/{messageId}"),
                        });
                }

                string truncatedAnswer = answer.Length <= TruncateThresholdLength ? answer : answer.Substring(0, 45) + "...";

                ThumbnailCard previewCard = new ThumbnailCard
                {
                    Title = $"<b>{HttpUtility.HtmlEncode(qnaDoc.Questions[0])}</b>",
                    Text = $"{HttpUtility.HtmlEncode(truncatedAnswer)} <br/>{HttpUtility.HtmlEncode(createdBy)} | {HttpUtility.HtmlEncode(createdAt)} <br/>",
                };

                messagingExtensionAttachments.Add(new Attachment
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = card,
                }.ToMessagingExtensionAttachment(previewCard.ToAttachment()));
            }

            return messagingExtensionAttachments;
        }

        /// <summary>
        /// no answer found card.
        /// </summary>
        /// <param name="question">question.</param>
        /// <returns>attachment.</returns>
        public Attachment NoAnswerCard(string question)
        {
            AdaptiveCard card = new AdaptiveCard("1.0");
            var container = new AdaptiveContainer()
            {
                Items = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Size = AdaptiveTextSize.Default,
                            Wrap = true,
                            Text = Strings.AnswerNotFound,
                        },
                    },
            };
            card.Body.Add(container);
            card.Actions.Add(
                new AdaptiveShowCardAction()
                {
                    Title = Strings.AddEntryTitle,
                    Card = this.UpdateEntry(question),
                });
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };

            return adaptiveCardAttachment;
        }

        /// <summary>
        /// welcome card.
        /// </summary>
        /// <returns>attachment.</returns>
        public Attachment WelcomeCard()
        {
            var card = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Text = Strings.WelcomeMessage,
                            Size = AdaptiveTextSize.Default,
                            Wrap = true,
                        },
                    },
            };

            card.Actions.Add(
               new AdaptiveSubmitAction()
               {
                   Title = Strings.AskQuestion,
                   Data = new AdaptiveSubmitActionData
                   {
                       MsTeams = new CardAction
                       {
                           Type = "task/fetch",
                       },
                   },
               });

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
        }

        /// <summary>
        /// updated answer card.
        /// </summary>
        /// <param name="question">question.</param>
        /// <param name="answer">answer.</param>
        /// <param name="editedBy">editedby.</param>
        /// <param name="isTest">boolean environment.</param>
        /// <returns>attachment.</returns>
        public Attachment AddedAnswer(string question, string answer, string editedBy, bool isTest)
        {
            if (!string.IsNullOrWhiteSpace(answer))
            {
                answer = answer.Equals(Constants.Unanswered) ? string.Empty : answer;
            }

            AdaptiveCard card = new AdaptiveCard("1.0");
            var container = new AdaptiveContainer()
            {
                Items = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Size = AdaptiveTextSize.Default,
                            Wrap = true,
                            Text = $"**{Strings.QuestionTitle}:** {question}",
                        },
                        new AdaptiveTextBlock
                        {
                            Size = AdaptiveTextSize.Default,
                            Wrap = true,
                            Text = string.IsNullOrWhiteSpace(answer) ? answer : $"**{Strings.AnswerTitle}:** {answer}",
                        },
                        new AdaptiveTextBlock
                        {
                            Size = AdaptiveTextSize.Small,
                            Wrap = true,
                            Text = string.Format(Strings.LastEdited, editedBy),
                        },
                    },
            };

            if (isTest)
            {
                container.Items.Add(new AdaptiveTextBlock
                {
                    Size = AdaptiveTextSize.Small,
                    Wrap = true,
                    Text = Strings.WaitMessageAnswer,
                });
            }

            card.Body.Add(container);

            card.Actions.Add(
                new AdaptiveShowCardAction()
                {
                    Title = Strings.UpdateEntryTitle,
                    Card = this.UpdateEntry(question, answer),
                });

            card.Actions.Add(
                new AdaptiveShowCardAction()
                {
                    Title = Strings.DeleteEntryTitle,
                    Card = this.DeleteEntry(question),
                });

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };

            return adaptiveCardAttachment;
        }

        /// <summary>
        /// update toggle card.
        /// </summary>
        /// <param name="question">question.</param>
        /// <returns>card.</returns>
        public AdaptiveCard AddQuestionAnswer()
        {
            AdaptiveCard card = new AdaptiveCard("1.0");
            var container = new AdaptiveContainer()
            {
                Items = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Text = Strings.QuestionTitle,
                            Size = AdaptiveTextSize.Small,
                        },
                        new AdaptiveTextInput
                        {
                            Id = "question",
                            Placeholder = Strings.PlaceholderQuestion,
                            MaxLength = QuestionMaxInputLength,
                            Style = AdaptiveTextInputStyle.Text,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = Strings.AnswerTitle,
                            Size = AdaptiveTextSize.Small,
                        },
                        new AdaptiveTextInput
                        {
                            Id = "answer",
                            Placeholder = Strings.PlaceholderAnswer,
                            IsMultiline = true,
                            MaxLength = AnswerMaxInputLength,
                            Style = AdaptiveTextInputStyle.Text,
                        },
                    },
            };
            card.Body.Add(container);

            card.Actions.Add(
               new AdaptiveSubmitAction()
               {
                   Title = Strings.Save,
                   Data = new AdaptiveSubmitActionData
                   {
                       MsTeams = new CardAction
                       {
                           Type = ActionTypes.MessageBack,
                           Text = Constants.SubmitAddCommand,
                       },
                   },
               });

            return card;
        }

        /// <summary>
        /// update toggle card.
        /// </summary>
        /// <param name="question">question.</param>
        /// <param name="answer">answer.</param>
        /// <returns>card.</returns>
        public AdaptiveCard UpdateEntry(string question, string answer = "")
        {
            AdaptiveCard card = new AdaptiveCard("1.0");
            var container = new AdaptiveContainer()
            {
                Items = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Text = Strings.QuestionTitle,
                            Size = AdaptiveTextSize.Small,
                        },
                        new AdaptiveTextInput
                        {
                            Id = "question",
                            MaxLength = QuestionMaxInputLength,
                            Placeholder = Strings.PlaceholderQuestion,
                            Style = AdaptiveTextInputStyle.Text,
                            Value = question,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = Strings.AnswerTitle,
                            Size = AdaptiveTextSize.Small,
                        },
                        new AdaptiveTextInput
                        {
                            Id = "answer",
                            Placeholder = Strings.PlaceholderAnswer,
                            IsMultiline = true,
                            MaxLength = AnswerMaxInputLength,
                            Style = AdaptiveTextInputStyle.Text,
                            Value = answer,
                        },
                    },
            };
            card.Body.Add(container);

            card.Actions.Add(
               new AdaptiveSubmitAction()
               {
                   Title = Strings.Save,
                   Data = new AdaptiveSubmitActionData
                   {
                       MsTeams = new CardAction
                       {
                           Type = ActionTypes.MessageBack,
                           Text = Constants.SaveCommand,
                       },
                       Details = new Details() { Question = question },
                   },
               });

            return card;
        }

        /// <summary>
        /// delete toggle card.
        /// </summary>
        /// <param name="question">question.</param>
        /// <returns>card.</returns>
        public AdaptiveCard DeleteEntry(string question)
        {
            AdaptiveCard card = new AdaptiveCard("1.0");
            var container = new AdaptiveContainer()
            {
                Items = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                           Text = Strings.DeleteConfirmation,
                           Wrap = true,
                        },
                    },
            };
            card.Body.Add(container);

            card.Actions.Add(
               new AdaptiveSubmitAction()
               {
                   Title = Strings.Yes,
                   Data = new AdaptiveSubmitActionData
                   {
                       MsTeams = new CardAction
                       {
                           Type = ActionTypes.MessageBack,
                           Text = Constants.DeleteCommand,
                       },
                       Details = new Details() { Question = question },
                   },
               });

            card.Actions.Add(
              new AdaptiveSubmitAction()
              {
                  Title = Strings.No,
                  Data = new AdaptiveSubmitActionData
                  {
                      MsTeams = new CardAction
                      {
                          Type = ActionTypes.MessageBack,
                          Text = Constants.NoCommand,
                      },
                  },
              });

            return card;
        }

        /// <summary>
        /// deleted item card.
        /// </summary>
        /// <param name="question">question.</param>
        /// <param name="answer">answer.</param>
        /// <param name="deletedBy">deleted by user.</param>
        /// <returns>card.</returns>
        public Attachment DeletedEntry(string question, string answer, string deletedBy)
        {
            AdaptiveCard card = new AdaptiveCard("1.0");
            var container = new AdaptiveContainer()
            {
                Items = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Size = AdaptiveTextSize.Default,
                            Wrap = true,
                            Text = $"**{Strings.QuestionTitle}:** {question}",
                        },
                        new AdaptiveTextBlock
                        {
                            Size = AdaptiveTextSize.Default,
                            Wrap = true,
                            Text = $"**{Strings.AnswerTitle}:** {answer}",
                        },
                        new AdaptiveTextBlock
                        {
                            Size = AdaptiveTextSize.Small,
                            Wrap = true,
                            Text = string.Format(Strings.DeletedQnaPair, deletedBy),
                        },
                        new AdaptiveTextBlock
                        {
                            Size = AdaptiveTextSize.Small,
                            Wrap = true,
                            Text = Strings.WaitMessageAnswer,
                        },
                    },
            };
            card.Body.Add(container);

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };

            return adaptiveCardAttachment;
        }

        /// <summary>
        /// Add question card task module.
        /// </summary>
        /// <param name="isValid">validation flag.</param>
        /// <returns>card.</returns>
        public Attachment AddQuestionActionCard(bool isValid)
        {
            AdaptiveCard card = new AdaptiveCard("1.0");
            var container = new AdaptiveContainer()
            {
                Items = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Text = Strings.QuestionTitle,
                            Size = AdaptiveTextSize.Small,
                        },
                        new AdaptiveTextInput
                        {
                            Id = "question",
                            Placeholder = Strings.PlaceholderQuestion,
                            MaxLength = QuestionMaxInputLength,
                            Style = AdaptiveTextInputStyle.Text,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = Strings.AnswerTitle,
                            Size = AdaptiveTextSize.Small,
                        },
                        new AdaptiveTextInput
                        {
                            Id = "answer",
                            Placeholder = Strings.PlaceholderAnswer,
                            IsMultiline = true,
                            MaxLength = AnswerMaxInputLength,
                            Style = AdaptiveTextInputStyle.Text,
                        },
                    },
            };

            if (!isValid)
            {
                container.Items.Add(new AdaptiveTextBlock
                {
                    Text = Strings.EmptyQnaValidation,
                    Size = AdaptiveTextSize.Small,
                    Color = AdaptiveTextColor.Attention,
                });
            }

            card.Body.Add(container);

            card.Actions.Add(
                new AdaptiveSubmitAction()
                {
                    Title = Strings.SubmitTitle,
                });

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };

            return adaptiveCardAttachment;
        }
    }
}
