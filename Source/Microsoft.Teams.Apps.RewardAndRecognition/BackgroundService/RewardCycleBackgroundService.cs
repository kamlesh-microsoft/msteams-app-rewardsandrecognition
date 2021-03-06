// <copyright file="RewardCycleBackgroundService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.BackgroundService
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Cronos;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// BackgroundService class that inherits IHostedService and implements the methods related to background tasks for updating award cycle one times a day.
    /// </summary>
    public class RewardCycleBackgroundService : IHostedService, IDisposable
    {
        /// <summary>
        /// Provides a parser and scheduler for Daily cron expression.
        /// </summary>
        private readonly CronExpression expression;

        /// <summary>
        /// Represents any time zone in the world.
        /// </summary>
        private readonly TimeZoneInfo timeZoneInfo;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<RewardCycleBackgroundService> logger;

        /// <summary>
        /// Instance of reward cycle helper which helps in updating reward cycle state.
        /// </summary>
        private readonly IRewardCycleHelper rewardCycleHelper;

        /// <summary>
        /// Instance of Timer for executing the service at particular interval.
        /// </summary>
        private System.Timers.Timer timer;

        /// <summary>
        /// Counter for number of times the service is executing.
        /// </summary>
        private int executionCount = 0;

        /// <summary>
        /// Flag to check whether Dispose is already called or not.
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="RewardCycleBackgroundService"/> class.
        /// BackgroundService class that inherits IHostedService and implements the methods related to award cycle.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="rewardCycleHelper">Helper to start/stop a reward cycle.</param>
        public RewardCycleBackgroundService(ILogger<RewardCycleBackgroundService> logger, IRewardCycleHelper rewardCycleHelper)
        {
            this.logger = logger;
            this.expression = CronExpression.Parse("0 */4 * * *"); // scheduled to run every 4 hrs.
            this.timeZoneInfo = TimeZoneInfo.Utc;
            this.rewardCycleHelper = rewardCycleHelper;
        }

        /// <summary>
        /// Method to start the background task when application starts.
        /// </summary>
        /// <param name="cancellationToken">Signals cancellation to the executing method.</param>
        /// <returns>A task instance.</returns>
        public async Task StartAsync(CancellationToken cancellationToken)
        {
            try
            {
                this.logger.LogInformation("RewardCycle Hosted Service is running.");
                await this.ScheduleCycleAsync();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while running the background service to send update award cycle): {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Triggered when the host is performing a graceful shutdown.
        /// </summary>
        /// <param name="cancellationToken">Signals cancellation to the executing method.</param>
        /// <returns>A task instance.</returns>
        public async Task StopAsync(CancellationToken cancellationToken)
        {
            this.logger.LogInformation("RewardCycle Hosted Service is stopping.");
            await Task.CompletedTask;
        }

        /// <summary>
        /// This code added to correctly implement the disposable pattern.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Protected implementation of Dispose pattern.
        /// </summary>
        /// <param name="disposing">True if already disposed else false.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (this.disposed)
            {
                return;
            }

            if (disposing)
            {
                this.timer.Dispose();
            }

            this.disposed = true;
        }

        /// <summary>
        /// Set the timer and enqueue send notification task if timer matched as per Cron expression.
        /// </summary>
        /// <returns>A task that Enqueue sends notification task.</returns>
        private async Task ScheduleCycleAsync()
        {
            var count = Interlocked.Increment(ref this.executionCount);
            this.logger.LogInformation("RewardCycle Hosted Service is working. Count: {Count}", count);

            var next = this.expression.GetNextOccurrence(DateTimeOffset.UtcNow, this.timeZoneInfo);
            if (next.HasValue)
            {
                var delay = next.Value - DateTimeOffset.UtcNow;
                this.timer = new System.Timers.Timer(delay.TotalMilliseconds);
                #pragma warning disable AvoidAsyncVoid // Avoid async void
                this.timer.Elapsed += async (sender, args) =>
                #pragma warning restore AvoidAsyncVoid // Avoid async void
                {
                    this.logger.LogInformation($"Timer matched to send notification at timer value : {this.timer}");
                    this.timer.Stop();  // reset timer

                    await this.UpdateCycleAsync();
                    await this.ScheduleCycleAsync();
                };
                this.timer.Start();
            }
        }

        /// <summary>
        /// Method invokes to check and update award cycle state.
        /// </summary>
        /// <returns>A task that sends notification in channel for group activity.</returns>
        private async Task UpdateCycleAsync()
        {
            this.logger.LogInformation("Check and update reward cycle");
            await this.rewardCycleHelper.CheckOrUpdateCycleStatusAsync();
            this.logger.LogInformation("updated reward cycle");
        }
    }
}
