namespace CenterBackend.Logging
{
    public interface IAppLogger
    {
        Task LogAsync(string message, string level = "INFO");
        Task LogInfoAsync(string message) => LogAsync(message, "INFO");
        Task LogWarnAsync(string message) => LogAsync(message, "WARN");
        Task LogErrorAsync(string message) => LogAsync(message, "ERROR");
    }
}