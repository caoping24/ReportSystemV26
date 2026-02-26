using Microsoft.AspNetCore.Builder;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Resources;
using System.Windows.Threading;

namespace ReportServer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
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

        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += MainWindow_Loaded;
            this.Closing += MainWindow_Closing;
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {

            //this.Hide();            // 隐藏窗口
            this.ShowInTaskbar = false;
            InitializeTray();// 初始化托盘图标与菜单
            await StartEmbeddedApiAsync();// 可选：启动时自动启动后端
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
            _openMainWindow.Click += (_, __) => Dispatcher.Invoke(ShowAndActivateWindow);
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
                StreamResourceInfo resourceInfo = System.Windows.Application.GetResourceStream(uri);
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
        private void ShowAndActivateWindow()
        {
            if (this.IsVisible == false)
            {
                this.Show();
                this.WindowState = WindowState.Normal;
                this.ShowInTaskbar = true;
            }
            this.Activate();
        }

        // 封装：打开系统默认浏览器并访问主页
        private void OpenBrowserToHomePage()
        {
            if (string.IsNullOrEmpty(HomePageUrl))
            {
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

        private async void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            e.Cancel = true;
            Hide();
            ShowInTaskbar = false;
            Topmost = false; // 隐藏时重置置顶状态
        }

        private void Window_GotFocus(object sender, RoutedEventArgs e)// 获得焦点时置顶
        {
            Topmost = true;
        }
        private void Window_LostFocus(object sender, RoutedEventArgs e)//失去焦点时取消置顶
        {
            if (IsVisible)
            {
                Topmost = false;
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
                lock (_apiLock)
                {
                    _apiApp = app;
                }
                Dispatcher.Invoke(UpdateMenuState);// 更新托盘菜单状态（在UI线程）
            }
            catch (Exception ex)
            {
                //Dispatcher.Invoke(() =>// 如果启动失败，提示用户（UI线程）
                //    System.Windows.MessageBox.Show($"启动服务失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error));
                //await ExitApplicationAsync();

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
                Dispatcher.Invoke(UpdateMenuState);
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
            //手动释放图标资源
            _iconStopped?.Dispose();
            _iconRunning?.Dispose();

            await Task.Delay(200);// 延迟一小段时间再关闭，给系统处理图标移除的时间
            System.Windows.Application.Current.Shutdown();
        }
    }
}