// <copyright file="Program.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
// Licensed under the MIT License.
// Generated with Bot Builder V4 SDK Template for Visual Studio CoreBot v4.5.0

namespace Microsoft.Teams.Apps.CrowdSourcer
{
    using Microsoft.AspNetCore;
    using Microsoft.AspNetCore.Hosting;

    /// <summary>
    /// Main class.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// main method of project.
        /// </summary>
        /// <param name="args">arguments.</param>
        public static void Main(string[] args)
        {
            CreateWebHostBuilder(args).Build().Run();
        }

        /// <summary>
        ///  Create WebHostBuilder and initialize startup.
        /// </summary>
        /// <param name="args">arguments.</param>
        /// <returns>WebHostBuilder.</returns>
        public static IWebHostBuilder CreateWebHostBuilder(string[] args) =>
            WebHost.CreateDefaultBuilder(args)
                .UseStartup<Startup>();
    }
}
