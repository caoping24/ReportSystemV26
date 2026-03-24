using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using CenterReport.Repository.IServices;
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using ReportServer.Services.IUserService;

namespace ReportServer.Services
{
    /// <summary>
    /// Hosted background service that runs at HH:00:01 every hour and invokes ICollectWinccDatas.ReadAndSaveDataAsync().
    /// </summary>
    public class HourlyHostedService : BackgroundService
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly SemaphoreSlim _semaphore = new(1, 1);

        public HourlyHostedService(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            try
            {
                while (!stoppingToken.IsCancellationRequested)
                {
                    DateTime now = DateTime.Now;

                    //Calculate next occurrence at HH: 00:01 strictly after 'now'.
                    DateTime next = new(now.Year, now.Month, now.Day, now.Hour, 0, 1);
                    if (now >= next)
                    {
                        // If we're at or past the trigger moment, schedule the next hour.
                        // This ensures we do NOT run immediately on startup even if the clock
                        // is exactly at HH:00:01 when the app starts.
                        next = next.AddHours(1);
                    }
                    // 每10s执行一次
                    //DateTime next = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second);
                    //if (now >= next)
                    //{
                    //    // If we're at or past the trigger moment, schedule the next hour.
                    //    // This ensures we do NOT run immediately on startup even if the clock
                    //    // is exactly at HH:00:01 when the app starts.
                    //    next = next.AddSeconds(10);
                    //}

                    Debug.WriteLine($"HourlyHostedService: now={now:O}, next={next:O}, delaySeconds={(next - now).TotalSeconds}");

                    if (now >= next)
                    {
                        next = next.AddSeconds(10);
                    }

                    var delay = next - now;
                    try
                    {
                        await Task.Delay(delay, stoppingToken);
                    }
                    catch (TaskCanceledException)
                    {
                        break;
                    }

                    if (stoppingToken.IsCancellationRequested)
                        break;

                    if (!await _semaphore.WaitAsync(0, stoppingToken))
                    {
                        Debug.WriteLine("HourlyHostedService: previous run still in progress, skipping this scheduled run.");
                        continue;
                    }

                    try
                    {
                        using var scope = _serviceProvider.CreateScope();
                        ICollectWinccDatas? collector = scope.ServiceProvider.GetService(typeof(ICollectWinccDatas)) as ICollectWinccDatas;
                        if (collector != null)
                        {
                            await collector.ReadAndSaveDataAsync();
                        }
                        else
                        {
                            Debug.WriteLine("HourlyHostedService: ICollectWinccDatas not available in scope.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"HourlyHostedService execution error: {ex}");
                    }
                    finally
                    {
                        try { _semaphore.Release(); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"HourlyHostedService fatal error: {ex}");
            }
        }

        public override void Dispose()
        {
            _semaphore.Dispose();
            base.Dispose();
        }
    }
}
