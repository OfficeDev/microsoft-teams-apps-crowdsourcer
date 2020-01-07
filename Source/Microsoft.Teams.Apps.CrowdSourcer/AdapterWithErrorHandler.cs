// <copyright file="AdapterWithErrorHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
// Licensed under the MIT License.
// Generated with Bot Builder V4 SDK Template for Visual Studio CoreBot v4.5.0

namespace Microsoft.Teams.Apps.CrowdSourcer
{
    using System;
    using Microsoft.ApplicationInsights;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.CrowdSourcer.Resources;

    /// <summary>
    /// Log any leaked exception from the application.
    /// </summary>
    public class AdapterWithErrorHandler : BotFrameworkHttpAdapter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AdapterWithErrorHandler"/> class.
        /// </summary>
        /// <param name="configuration">configuration.</param>
        /// <param name="telemetryClient">telemetry client.</param>
        /// <param name="conversationState">conversation state.</param>
        public AdapterWithErrorHandler(IConfiguration configuration, TelemetryClient telemetryClient, ConversationState conversationState = null)
            : base(configuration)
        {
            this.OnTurnError = async (turnContext, exception) =>
            {
                // Log any leaked exception from the application.
                telemetryClient.TrackException(exception);

                // Send a catch-all apology to the user.
                var errorMessage = MessageFactory.Text(Strings.ErrorMsgText, Strings.ErrorMsgText, InputHints.ExpectingInput);
                await turnContext.SendActivityAsync(errorMessage);

                if (conversationState != null)
                {
                    try
                    {
                        // Delete the conversationState for the current conversation to prevent the
                        // bot from getting stuck in a error-loop caused by being in a bad state.
                        // ConversationState should be thought of as similar to "cookie-state" in a Web pages.
                        await conversationState.DeleteAsync(turnContext);
                    }
                    catch (Exception e)
                    {
                        telemetryClient.TrackException(e);
                    }
                }
            };
        }
    }
}
