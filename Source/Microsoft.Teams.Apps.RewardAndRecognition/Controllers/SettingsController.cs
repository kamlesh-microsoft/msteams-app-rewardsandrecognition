// <copyright file="SettingsController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Controllers
{
    using System;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// Controller to get bot settings related request.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize]
    public class SettingsController : ControllerBase
    {
        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly IConfiguration configuration;

        /// <summary>
        /// Initializes a new instance of the <see cref="SettingsController"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="configuration">configuration.</param>
        public SettingsController(ILogger<SettingsController> logger, IConfiguration configuration)
        {
            this.logger = logger;
            this.configuration = configuration;
        }

        /// <summary>
        /// Get bot setting to client application.
        /// </summary>
        /// <returns>Bot id.</returns>
        [HttpGet("botsettings")]
        public IActionResult GetBotSettings()
        {
            try
            {
                return this.Ok(new
                {
                    botId = this.configuration["MicrosoftAppId"],
                    instrumentationKey = this.configuration["ApplicationInsights:InstrumentationKey"],
                });
            }
            #pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                this.logger.LogError(ex, "Error while fetching bot setting.");
                throw;
            }
        }
    }
}