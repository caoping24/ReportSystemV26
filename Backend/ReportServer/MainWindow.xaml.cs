
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Threading;
namespace ReportServer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            //await _collectWinccDatas.ReadAndSaveDataAsync();
            InitializeComponent();
            this.Closing += MainWindow_Closing;
            this.GotFocus += Window_GotFocus;
            this.LostFocus += Window_LostFocus;
        }

        public void ShowAndActivateWindow()
        {
            if (!IsVisible)
            {
                Show();
                WindowState = WindowState.Normal;
                ShowInTaskbar = true;
            }
            Activate();
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
            if (IsVisible) Topmost = false;
        }

    }
}