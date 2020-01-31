// <copyright file="Startup.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
// Licensed under the MIT License.
// Generated with Bot Builder V4 SDK Template for Visual Studio CoreBot v4.5.0

namespace Microsoft.Teams.Apps.CrowdSourcer
{
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Teams.Apps.CrowdSourcer.Bots;
    using Microsoft.Teams.Apps.CrowdSourcer.Cards;
    using Microsoft.Teams.Apps.CrowdSourcer.Common.Providers;

    /// <summary>
    /// This a Startup class for this Bot.
    /// </summary>
    public class Startup
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Startup"/> class.
        /// </summary>
        /// <param name="configuration">Startup Configuration.</param>
        public Startup(Extensions.Configuration.IConfiguration configuration)
        {
            this.Configuration = configuration;
        }

        /// <summary>
        /// Gets Configurations Interfaces.
        /// </summary>
        public Extensions.Configuration.IConfiguration Configuration { get; }

        /// <summary>
        /// This method gets called by the runtime. Use this method to add services to the container.
        /// </summary>
        /// <param name="services"> Service Collection Interface.</param>
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_1);
            services.AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();
            services.AddSingleton<IBotFrameworkHttpAdapter, BotFrameworkHttpAdapter>();
            services.AddSingleton(new MicrosoftAppCredentials(this.Configuration["MicrosoftAppId"], this.Configuration["MicrosoftAppPassword"]));
            services.AddSingleton<IConfigurationStorageProvider>(new ConfigurationStorageProvider(this.Configuration));
            services.AddSingleton<IObjectIdToNameMapper>(new ObjectIdToNameMapper(this.Configuration));

            IQnAMakerClient qnaMakerClient = new QnAMakerClient(new ApiKeyServiceClientCredentials(this.Configuration["QnAMakerSubscriptionKey"])) { Endpoint = this.Configuration["QnAMakerApiUrl"] };
            string endpointKey = Task.Run(() => qnaMakerClient.EndpointKeys.GetKeysAsync()).Result.PrimaryEndpointKey;

            services.AddSingleton<IQnaServiceProvider>((provider) => new QnaServiceProvider(
                provider.GetRequiredService<IConfigurationStorageProvider>(),
                this.Configuration,
                qnaMakerClient,
                new QnAMakerRuntimeClient(new EndpointKeyServiceClientCredentials(endpointKey)) { RuntimeEndpoint = this.Configuration["QnAMakerHostUrl"] }));

            services.AddSingleton<ISearchService, SearchService>();
            services.AddSingleton<CrowdSourcerCards>();

            services.AddTransient<IBot>((provider) => new CrowdSourcerBot(
                provider.GetRequiredService<TelemetryClient>(),
                provider.GetRequiredService<IQnaServiceProvider>(),
                this.Configuration,
                provider.GetRequiredService<IConfigurationStorageProvider>(),
                provider.GetRequiredService<IObjectIdToNameMapper>(),
                provider.GetRequiredService<ISearchService>(),
                provider.GetRequiredService<CrowdSourcerCards>()));

            services.AddApplicationInsightsTelemetry();
            services.AddLocalization();
        }

        /// <summary>
        /// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        /// </summary>
        /// <param name="app">Application Builder.</param>
        /// <param name="env">Hosting Environment.</param>
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseHsts();
            }

            app.UseDefaultFiles();
            app.UseStaticFiles();
            app.UseMvc();
        }
    }
}
