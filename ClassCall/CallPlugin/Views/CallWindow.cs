using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using CallPlugin.Models;

namespace CallPlugin.Views;

/// <summary>
/// 呼叫显示窗口（Avalonia，运行在 ClassIsland 2.x 主程序进程内）
/// 两种模式配色：紧急呼叫=红 / 普通通知=蓝，置顶显示
/// </summary>
public static class CallWindow
{
    public static void Show(CallMessage call)
    {
        Dispatcher.UIThread.Post(() =>
        {
            var (accent, icon, label) = call.Type switch
            {
                "urgent" => (new SolidColorBrush(Color.FromRgb(0xe5, 0x39, 0x35)), "🚨", "紧急呼叫"),
                _ => (new SolidColorBrush(Color.FromRgb(0x25, 0x63, 0xeb)), "📢", "呼叫通知")
            };

            var win = new Window
            {
                Title = $"{label} - {call.Title}",
                Width = 720,
                Height = 420,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Topmost = true,
                CanResize = false,
                FontSize = 16,
                Background = Brushes.White
            };

            var stack = new StackPanel { Margin = new Thickness(30), Spacing = 12 };

            // 顶部：图标 + 类型标签
            stack.Children.Add(new TextBlock
            {
                Text = $"{icon} {label}",
                FontSize = 18,
                FontWeight = FontWeight.Bold,
                Foreground = accent
            });

            // 标题
            stack.Children.Add(new TextBlock
            {
                Text = call.Title,
                FontSize = 26,
                FontWeight = FontWeight.Bold,
                Foreground = Brushes.Black,
                TextWrapping = TextWrapping.Wrap
            });

            // 内容
            if (!string.IsNullOrWhiteSpace(call.Message))
            {
                stack.Children.Add(new TextBlock
                {
                    Text = call.Message,
                    FontSize = 17,
                    TextWrapping = TextWrapping.Wrap,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x42, 0x42, 0x42))
                });
            }

            // 发送者
            if (!string.IsNullOrWhiteSpace(call.Sender))
            {
                stack.Children.Add(new TextBlock
                {
                    Text = $"发送人：{call.Sender}    {call.CreatedAt:HH:mm:ss}",
                    FontSize = 12,
                    Foreground = Brushes.Gray
                });
            }

            var ok = new Button
            {
                Content = "知道了",
                Width = 110,
                Height = 36,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            ok.Click += (_, _) => win.Close();
            stack.Children.Add(ok);

            win.Content = stack;
            win.Show();
        });
    }
}
