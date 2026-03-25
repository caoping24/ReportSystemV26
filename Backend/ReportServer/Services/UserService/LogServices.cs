using System.Collections.Concurrent;
using System.IO;
using ReportServer.Models;

namespace ReportServer.Services.UserService
{
    public class LogServices
    {

        /// <summary>
        /// WPF异步日志工具类（后台调用，线程安全）
        /// </summary>
        public static class AsyncLogHelper
        {
            // 线程安全队列：缓存待写入的日志消息
            private static readonly ConcurrentQueue<string> _logQueue = new();
            // 信号量：控制并发写入，避免多个线程同时写文件
            private static readonly SemaphoreSlim _semaphore = new(1, 1);
            // 日志文件存储路径（应用程序目录下的Logs文件夹）
            private static readonly string _logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");

            /// <summary>
            /// 异步写入日志（核心方法，后台直接调用）
            /// </summary>
            /// <param name="message">日志内容</param>
            /// <param name="level">日志级别</param>
            /// <param name="token">取消令牌（可选）</param>
            public static async Task LogAsync(string message, LogLevel level = LogLevel.Info, CancellationToken token = default)
            {
                try
                {
                    // 1. 格式化日志消息（时间 + 级别 + 内容）
                    var logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level.ToString().ToUpper()}] {message}";

                    // 2. 将日志消息加入线程安全队列
                    _logQueue.Enqueue(logMessage);

                    // 3. 异步处理队列中的日志写入文件
                    await ProcessLogQueueAsync(token);
                }
                catch (Exception ex)
                {
                    // 日志写入失败时的兜底处理（避免崩溃）
                    Console.WriteLine($"日志写入异常：{ex.Message}");
                }
            }

            /// <summary>
            /// 异步处理日志队列，写入文件
            /// </summary>
            private static async Task ProcessLogQueueAsync(CancellationToken token)
            {
                // 如果队列为空，直接返回
                if (_logQueue.IsEmpty) return;

                try
                {
                    // 等待信号量（确保同一时间只有一个线程写入文件）
                    await _semaphore.WaitAsync(token);

                    // 确保日志目录存在
                    if (!Directory.Exists(_logDirectory))
                    {
                        Directory.CreateDirectory(_logDirectory);
                    }

                    // 日志文件名：按日期分割（例如：2026-03-23.log）
                    var logFileName = $"{DateTime.Now:yyyy-MM-dd}.log";
                    var logFilePath = Path.Combine(_logDirectory, logFileName);

                    // 循环处理队列中的所有日志消息
                    while (_logQueue.TryDequeue(out var logMessage) && !token.IsCancellationRequested)
                    {
                        // 异步写入文件（追加模式，UTF8编码）
                        await File.AppendAllTextAsync(logFilePath, logMessage + Environment.NewLine, token);
                    }
                }
                finally
                {
                    // 释放信号量，允许其他线程继续写入
                    _semaphore.Release();
                }
            }

            /// <summary>
            /// 便捷方法：异步写入错误日志
            /// </summary>
            public static async Task LogErrorAsync(string message, CancellationToken token = default)
            {
                await LogAsync(message, LogLevel.Error, token);
            }

            /// <summary>
            /// 便捷方法：异步写入警告日志
            /// </summary>
            public static async Task LogWarningAsync(string message, CancellationToken token = default)
            {
                await LogAsync(message, LogLevel.Warning, token);
            }

            /// <summary>
            /// 便捷方法：异步写入信息日志
            /// </summary>
            public static async Task LogInfoAsync(string message, CancellationToken token = default)
            {
                await LogAsync(message, LogLevel.Info, token);
            }
        }
    }
}
