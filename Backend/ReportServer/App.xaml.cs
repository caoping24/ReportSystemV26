using Microsoft.AspNetCore.Builder;
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

        // 应用启动入口（最重要）
        protected override async void OnStartup(StartupEventArgs e)
        {
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

            var exitMenu = new ToolStripMenuItem("退出");
            exitMenu.Click += async (_, __) => await ExitApplicationAsync();
            menu.Items.Add(exitMenu);

            // 1. 加载图标
            _iconRunning = LoadIconFromResource("pack://application:,,,/AppIco/SL_Icon_Green.ico");
            _iconStopped = LoadIconFromResource("pack://application:,,,/AppIco/SL_Icon_Gray.ico");
            // 2. 异常回退：使用系统图标兜底
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
        //打开系统默认浏览器并访问主页
        private void OpenBrowserToHomePage()
        {
            if (string.IsNullOrEmpty(HomePageUrl))
            {
                //Process.Start(new ProcessStartInfo(App.HomePageUrl) { UseShellExecute = true });
                Dispatcher.Invoke(() =>
                    System.Windows.MessageBox.Show("主页地址未配置！", "错误", MessageBoxButton.OK, MessageBoxImage.Error));
                return;
            }
            try
            {
                Process.Start(new ProcessStartInfo(HomePageUrl)// 调用系统默认浏览器打开URL
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>// 异常处理：提示用户手动访问
                    System.Windows.MessageBox.Show(
                        $"打开浏览器失败：{ex.Message}\n请手动访问主页：{HomePageUrl}",
                        "访问失败",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning
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

            try
            {
                // 直接用程序集目录
                string webApiProjectDir = Path.GetDirectoryName(typeof(CenterBackend.Program).Assembly.Location) ?? AppContext.BaseDirectory;
                string contentRootPath = Path.GetFullPath(webApiProjectDir);
                int port = 5260;
                // 传入正确的 contentRootPath
                var app = CenterBackend.Program.BuildWebApplication(Array.Empty<string>(), contentRootPath, port);
                await app.StartAsync();

                lock (_apiLock) _apiApp = app;
                UpdateMenuState();
                //Dispatcher.Invoke(UpdateMenuState);// 更新托盘菜单状态（在UI线程）
            }
            catch (Exception ex)
            {
                // 输出详细异常信息（包含内部异常和调用栈）
                string errorMsg = $"启动服务失败：{ex.Message}\n" +
                                 $"内部异常：{ex.InnerException?.Message}\n";
                Dispatcher.Invoke(() =>
                    System.Windows.MessageBox.Show(errorMsg, "错误", MessageBoxButton.OK, MessageBoxImage.Error));
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
                Dispatcher.Invoke(() =>
                    System.Windows.MessageBox.Show($"停止服务失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error));
            }
            finally
            {
                UpdateMenuState();
                //Dispatcher.Invoke(UpdateMenuState);
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
        // 应用退出兜底
        //protected override void OnExit(ExitEventArgs e)
        //{
        //    _notifyIcon?.Dispose();
        //    _apiApp?.DisposeAsync();
        //    base.OnExit(e);
        //}
    }

}
