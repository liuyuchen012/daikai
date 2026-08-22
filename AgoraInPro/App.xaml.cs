using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace CheckIn.Client;

/// <summary>
/// 应用程序入口类，负责启动流程控制
/// 先显示启动画面（SplashScreen），初始化主窗口后关闭启动画面
/// </summary>
public partial class App : Application
{
    /// <summary>日志目录（应用目录下 logs 文件夹）</summary>
    internal static readonly string LogDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");

    /// <summary>
    /// 应用程序启动事件处理
    /// 显示启动画面 -> 初始化主窗口 -> 显示主窗口并关闭启动画面
    /// L1：移除硬编码的 2 秒等待，仅保留极短延迟确保启动画面完成首帧渲染
    /// </summary>
    private async void Application_Startup(object sender, StartupEventArgs e)
    {
        // 全局异常兜底：任何未处理异常都写入日志，避免程序直接闪退后无从排查
        DispatcherUnhandledException += (_, args) =>
        {
            LogError("DispatcherUnhandledException", args.Exception);
            MessageBox.Show($"程序发生未处理的异常：\n{args.Exception.Message}\n\n详细信息已写入日志文件：\n{Path.Combine(LogDir, "error.log")}",
                "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            args.Handled = true;
        };
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            LogError("AppDomainUnhandledException", args.ExceptionObject as Exception);
        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            LogError("UnobservedTaskException", args.Exception);
            args.SetObserved();
        };

        // 诊断模式：--selftest 直接构造控制中心，验证是否抛异常，结果写入 selftest.log
        if (e.Args.Length > 0 && e.Args.Contains("--selftest"))
        {
            var log = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "selftest.log");
            try
            {
                _ = new ControlCenterView();
                File.WriteAllText(log, $"OK: ControlCenterView constructed at {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n");
            }
            catch (Exception ex)
            {
                File.WriteAllText(log, $"ERR: {ex}\n");
            }
            Shutdown();
            return;
        }

        ShutdownMode = ShutdownMode.OnExplicitShutdown;

        // 显示启动画面（带呼吸动画和进度条）
        var splash = new SplashScreen();
        splash.Show();

        await Task.Delay(50);

        splash.Close();

        Window mainWindow = new MainWindow();
        MainWindow = mainWindow;
        mainWindow.Show();
        ShutdownMode = ShutdownMode.OnMainWindowClose;
    }

    /// <summary>记录异常到日志文件</summary>
    internal static void LogError(string tag, Exception? ex)
    {
        try
        {
            Directory.CreateDirectory(LogDir);
            File.AppendAllText(Path.Combine(LogDir, "error.log"),
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{tag}] {ex}\r\n\r\n");
        }
        catch { }
    }
}
