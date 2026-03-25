using Microsoft.Extensions.Configuration;
using System.IO;
namespace ReportServer.Models
{
    public class RemoteWinccTags
    {
        private static IConfiguration? _configuration;
        private const string HomePageUrl = "http://localhost:5260/user/login"; // 主页地址（常量，便于修改）
        
        public static void Initialize()// 初始化配置
        {
            if (_configuration != null)
            {
                throw new InvalidOperationException("RemoteWinccTags 已经初始化过了。");
            }
            var basePath = Directory.GetCurrentDirectory();

            _configuration = new ConfigurationBuilder()
                .SetBasePath(basePath)
                .AddJsonFile("config\\appsettings.json", optional: false, reloadOnChange: true)
                .Build();
        }
        /// <summary>
        /// 获取服务器前缀
        /// </summary>
        public static string ServerPrefix => GetConfigValue("Perfix");
        /// <summary>
        /// 获取S7连接标签名称
        /// </summary>
        public static string WinccS7ConnectionTagName => GetConfigValue("S7ConnectionTagName");

        // 【优化2】内部辅助方法，减少重复代码
        private static string GetConfigValue(string key)
        {
            if (_configuration == null)
            {
                throw new InvalidOperationException("RemoteWinccTags 尚未初始化，请先调用 Initialize()。");
            }
            return _configuration.GetSection("WinccRemoteServerTag")[key] ?? string.Empty;
        }

    }
}
