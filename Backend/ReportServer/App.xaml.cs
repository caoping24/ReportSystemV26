using CenterReport.Repository;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Services;
using Microsoft.AspNetCore.Builder;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ReportServer.Models;
using ReportServer.Services;
using ReportServer.Services.IUserService;
using ReportServer.Services.UserService;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Resources;

namespace ReportServer
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        private WebApplication? _apiApp;
        private NotifyIcon? _notifyIcon;
        private ToolStripMenuItem? _startMenuItem;
        private ToolStripMenuItem? _stopMenuItem;
        private ToolStripMenuItem? _openMainWindow;
        private const string HomePageUrl = "http://localhost:5260/user/login"; // 主页地址（常量，便于修改）
        private readonly object _apiLock = new();
        private Icon? _iconRunning; // 服务运行时图标（图标A）
        private Icon? _iconStopped; // 服务停止时图标（图标B）

        private IHost? _host;
        // 应用启动入口（最重要）
        protected override async void OnStartup(StartupEventArgs e)
        {
            try
            {
                var builder = Host.CreateDefaultBuilder()
                    .ConfigureAppConfiguration((context, config) =>
                    {
                        config.SetBasePath(Directory.GetCurrentDirectory())
                              .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
                    })
                    .ConfigureServices((context, services) =>
                    {
                        var configuration = context.Configuration;
                        var defaultConnection = configuration.GetConnectionString("DefaultConnection") ?? string.Empty;

                        // 注册仓储
                        services.AddScoped(typeof(IReportRepository<>), typeof(ReportRepository<>));
                        services.AddScoped(typeof(IReportRecordRepository<>), typeof(ReportRecordRepository<>));
                        services.AddScoped(typeof(IReportUnitOfWork), typeof(ReportUnitOfWork));

                        services.AddScoped<ITagReadServices, TagReadServices>();
                        services.AddScoped<ITagDataConverter, TagDataConverter>();
                        services.AddScoped<ICollectWinccDatas, CollectWinccDatas>();

                        services.AddDbContext<CenterReportDbContext>(options => options.UseSqlServer(defaultConnection));
                        //// ViewModel
                        //services.AddScoped<MainViewModel>();
                        // Hosted service
                        services.AddHostedService<HourlyHostedService>();
                    });
                _host = builder.Build();
                _host.Start();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"应用启动失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown();
            }
            RemoteWinccTags.Initialize();//初始化读取json
            //
            base.OnStartup(e);
            InitializeTray();// 初始化全局托盘
            await StartEmbeddedApiAsync();// 自动启动后端服务
            MainWindow = new MainWindow();// 创建主窗口（默认隐藏）
        }
        private void InitializeTray()
        {
            if (_notifyIcon != null) return;// 如果已初始化，跳过
            var menu = new ContextMenuStrip();// 创建托盘菜单

            _startMenuItem = new ToolStripMenuItem("启动后端");
            _startMenuItem.Click += async (_, __) => await StartEmbeddedApiAsync();
            menu.Items.Add(_startMenuItem);

            _stopMenuItem = new ToolStripMenuItem("停止后端");
            _stopMenuItem.Click += async (_, __) => await StopEmbeddedApiAsync();
            menu.Items.Add(_stopMenuItem);

            menu.Items.Add(new ToolStripSeparator());

            _openMainWindow = new ToolStripMenuItem("系统信息");
            _openMainWindow.Click += (_, __) => Current.Dispatcher.Invoke(() =>
                       {
                           if (MainWindow is MainWindow window) window.ShowAndActivateWindow();
                       });
            menu.Items.Add(_openMainWindow);
            //查看日志菜单
            menu.Items.Add(new ToolStripSeparator());
            var logMenuItem = new ToolStripMenuItem("查看日志");
            logMenuItem.Click += async (_, __) => await OpenLogFolder();
            menu.Items.Add(logMenuItem);
            //测试归档
            menu.Items.Add(new ToolStripSeparator());
            var testComItem = new ToolStripMenuItem("测试连接");
            testComItem.Click += async (_, __) => await TestWinccComAsync();
            menu.Items.Add(testComItem);

            var testMenuItem = new ToolStripMenuItem("测试归档");
            testMenuItem.Click += async (_, __) => await TestWinccDataWriteAsync();
            menu.Items.Add(testMenuItem);

            var exitMenu = new ToolStripMenuItem("退出");
            exitMenu.Click += async (_, __) => await ExitApplicationAsync();
            menu.Items.Add(exitMenu);

            //加载图标
            _iconRunning = LoadIconFromResource("pack://application:,,,/AppIco/SL_Icon_Green.ico");
            _iconStopped = LoadIconFromResource("pack://application:,,,/AppIco/SL_Icon_Gray.ico");
            //异常回退：使用系统图标兜底
            if (_iconRunning == null) _iconRunning = SystemIcons.Shield;
            if (_iconStopped == null) _iconStopped = SystemIcons.Application;

            _notifyIcon = new NotifyIcon
            {
                Icon = _iconStopped!, // 初始状态：服务未启动
                Text = "ReportServer RT",
                ContextMenuStrip = menu,
                Visible = true
            };
            _notifyIcon.DoubleClick += (_, __) => Dispatcher.Invoke(OpenBrowserToHomePage);// 双击托盘显示窗口
            UpdateMenuState();
        }
        private Icon? LoadIconFromResource(string packUri)// 从资源加载图标
        {
            try
            {
                Uri uri = new Uri(packUri, UriKind.Absolute);
                StreamResourceInfo resourceInfo = GetResourceStream(uri);
                if (resourceInfo?.Stream != null)
                {
                    return new Icon(resourceInfo.Stream, 32, 32); // 固定32x32适配托盘
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                    System.Windows.MessageBox.Show($"加载图标失败：{ex.Message}", "警告", MessageBoxButton.OK, MessageBoxImage.Warning));
            }
            return null;
        }
        private void OpenBrowserToHomePage()//打开系统默认浏览器并访问主页
        {
            if (string.IsNullOrEmpty(HomePageUrl))
            {
                System.Windows.MessageBox.Show("主页地址未配置！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try
            {
                Process.Start(new ProcessStartInfo(HomePageUrl)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>// 异常处理：提示用户手动访问
                    System.Windows.MessageBox.Show($"打开浏览器失败：{ex.Message}\n请手动访问主页：{HomePageUrl}", "访问失败", MessageBoxButton.OK, MessageBoxImage.Warning
                    )
                );
            }
        }
        private async Task StartEmbeddedApiAsync()
        {
            lock (_apiLock)
            {
                if (_apiApp != null) return; // 已经启动
            }
            try // 直接用程序集目录
            {
                string webApiProjectDir = Path.GetDirectoryName(typeof(CenterBackend.Program).Assembly.Location) ?? AppContext.BaseDirectory;
                string contentRootPath = Path.GetFullPath(webApiProjectDir);
                int port = 5260;
                var app = await CenterBackend.Program.BuildWebApplicationAsync(Array.Empty<string>(), contentRootPath, port);// 传入正确的 contentRootPath
                await app.StartAsync();
                lock (_apiLock) _apiApp = app;
                UpdateMenuState();
            }
            catch (Exception ex)// 输出详细异常信息（包含内部异常和调用栈）
            {
                string errorMsg = $"启动服务失败：{ex.Message}\n" + $"内部异常：{ex.InnerException?.Message}\n";
                System.Windows.MessageBox.Show(errorMsg, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                await ExitApplicationAsync();

            }
        }
        private async Task StopEmbeddedApiAsync()
        {
            WebApplication? appToStop = null;
            lock (_apiLock)
            {
                if (_apiApp == null) return;
                appToStop = _apiApp;
                _apiApp = null;
            }
            try
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5)); // StopAsync 需要 CanCellationToken
                await appToStop!.StopAsync(cts.Token);
                await appToStop.DisposeAsync();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"停止服务失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                UpdateMenuState();
            }
        }
        //打开日志文件夹
        private async Task OpenLogFolder()
        {
            await Task.CompletedTask; // 适配 async 语法
            try
            {
                string appRootPath = AppContext.BaseDirectory;
                string logFolder = Path.Combine(appRootPath, "Logs");

                if (!Directory.Exists(logFolder))
                {
                    Directory.CreateDirectory(logFolder);
                    Dispatcher.Invoke(() =>
                        System.Windows.MessageBox.Show($"日志文件夹不存在，已自动创建！\n路径：{logFolder}", "提示", MessageBoxButton.OK, MessageBoxImage.Information));
                }

                Process.Start(new ProcessStartInfo(logFolder)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"打开日志文件夹失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private async Task TestWinccComAsync()
        {
            try
            {
                if (_host == null || _host.Services == null)// 校验DI容器是否就绪
                {
                    System.Windows.MessageBox.Show("DI容器未初始化，无法执行测试！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                using var scope = _host.Services.CreateScope(); // 创建作用域（适配Scoped生命周期）
                var tagReadServices = scope.ServiceProvider.GetRequiredService<ITagReadServices>();
                bool result = await tagReadServices.GetConnectStatus();
                System.Windows.MessageBox.Show(result ? "s7连接正常!" : "s7连接断开!", "测试结果", MessageBoxButton.OK,
                    result ? MessageBoxImage.Information : MessageBoxImage.Warning
                );// 反馈执行结果
            }
            catch (Exception ex)// 异常兜底
            {
                System.Windows.MessageBox.Show($"测试执行异常：{ex.Message}\n{ex.StackTrace}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task TestWinccDataWriteAsync()
        {
            try
            {
                if (_host == null || _host.Services == null)// 校验DI容器是否就绪
                {
                    System.Windows.MessageBox.Show("DI容器未初始化，无法执行测试！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                using var scope = _host.Services.CreateScope(); // 创建作用域（适配Scoped生命周期）
                var collectWinccDatas = scope.ServiceProvider.GetRequiredService<ICollectWinccDatas>();
                bool result = await collectWinccDatas.ReadAndSaveDataAsync();// 执行数据写入逻辑
                System.Windows.MessageBox.Show(result ? "Ok！" : "Error!！", "测试结果", MessageBoxButton.OK,
                    result ? MessageBoxImage.Information : MessageBoxImage.Warning
                );// 反馈执行结果
            }
            catch (Exception ex)// 异常兜底
            {
                System.Windows.MessageBox.Show($"测试执行异常：{ex.Message}\n{ex.StackTrace}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void UpdateMenuState()
        {
            if (!Dispatcher.CheckAccess())// 必须在 UI 线程执行（服务启动/停止是异步操作，可能触发非 UI 线程调用）
            {
                Dispatcher.Invoke(UpdateMenuState);
                return;
            }
            bool isServiceRunning = _apiApp != null;
            if (_startMenuItem != null) _startMenuItem.Enabled = !isServiceRunning;
            if (_stopMenuItem != null) _stopMenuItem.Enabled = isServiceRunning;

            if (_notifyIcon != null)// 根据服务状态切换托盘图标
            {
                _notifyIcon.Icon = isServiceRunning ? _iconRunning! : _iconStopped!;
                _notifyIcon.Text = isServiceRunning ? "ReportServer（服务运行中）" : "ReportServer（服务已停止）";
            }
        }
        private async Task ExitApplicationAsync()
        {
            try
            {
                await StopEmbeddedApiAsync();
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                    System.Windows.MessageBox.Show($"退出应用失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error));
            }
            Dispatcher.Invoke(() =>
            {
                if (_notifyIcon != null)
                {
                    _notifyIcon.Visible = false; // 先隐藏
                    _notifyIcon.Dispose();      // 释放资源
                    _notifyIcon = null;
                }
            });
            _iconStopped?.Dispose();//手动释放图标资源
            _iconRunning?.Dispose();

            await Task.Delay(200);// 延迟一小段时间再关闭，给系统处理图标移除的时间
            Current.Shutdown();
        }
        //应用退出兜底
        protected override async void OnExit(ExitEventArgs e)
        {
            _notifyIcon?.Dispose();
            _apiApp?.DisposeAsync();
            if (_host != null)
            {
                try
                {
                    await _host.StopAsync(TimeSpan.FromSeconds(5));
                    _host.Dispose();
                }
                catch { }
            }
            base.OnExit(e);
        }
    }

}
