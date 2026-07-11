using System.Threading.Tasks;
using System.Windows;

namespace CheckIn.Client;

/// <summary>
/// 应用程序入口类，负责启动流程控制
/// 先显示启动画面（SplashScreen），初始化主窗口后关闭启动画面
/// </summary>
public partial class App : Application
{
    /// <summary>
    /// 应用程序启动事件处理
    /// 显示启动画面 -> 初始化主窗口（等待最少2秒让动画完整播放）-> 显示主窗口并关闭启动画面
    /// </summary>
    private async void Application_Startup(object sender, StartupEventArgs e)
    {
        // 显示启动画面（带呼吸动画和进度条）
        var splash = new SplashScreen();
        splash.Show();

        // 初始化主窗口，同时等待至少 2 秒确保启动动画完整展示
        var mainWindow = new MainWindow();
        await Task.Delay(2000);

        // 显示主窗口并关闭启动画面
        mainWindow.Show();
        splash.Close();
    }
}
