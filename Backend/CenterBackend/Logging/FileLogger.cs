using System.Text;

namespace CenterBackend.Logging
{
    public class FileLogger : IAppLogger, IDisposable
    {
        private readonly string _logDirectory;
        private readonly SemaphoreSlim _semaphore = new(1, 1);
        private bool _disposed;

        public FileLogger(IWebHostEnvironment env)
        {
            // 使用 ContentRootPath 定位到项目根目录，确保在 WPF 嵌入场景也能正确定位
            _logDirectory = Path.Combine(env.ContentRootPath ?? AppContext.BaseDirectory, "wwwroot", "log");
            Directory.CreateDirectory(_logDirectory);
        }

        public async Task LogAsync(string message, string level = "INFO")
        {
            if (_disposed) throw new ObjectDisposedException(nameof(FileLogger));
            string fileName = DateTime.Now.ToString("yyyy-MM-dd") + ".log";
            string filePath = Path.Combine(_logDirectory, fileName);
            string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] {message}{Environment.NewLine}";

            await _semaphore.WaitAsync();
            try
            {
                // 使用 AppendAllTextAsync 保证简单可靠
                await File.AppendAllTextAsync(filePath, line, Encoding.UTF8);
            }
            finally
            {
                _semaphore.Release();
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _semaphore.Dispose();
            _disposed = true;
        }
    }
}