using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using AgoraIn.ClassIslandPlugin.Models;

namespace AgoraIn.ClassIslandPlugin.Views;

/// <summary>
/// 呼叫显示窗口（Avalonia，运行在 ClassIsland 2.x 主程序进程内）
/// 三种模式配色：待下课=蓝 / 上课应急=红 / 下课传唤=橙，置顶显示
/// </summary>
public static class CallWindow
{
    public static void Show(CallMessage call)
    {
        Dispatcher.UIThread.Post(() =>
        {
            var (accent, icon, label) = call.Type switch
            {
                "emergency" => (new SolidColorBrush(Color.FromRgb(0xe5, 0x39, 0x35)), "🚨", "上课应急通知"),
                "summon" => (new SolidColorBrush(Color.FromRgb(0xfb, 0x8c, 0x00)), "📢", "下课传唤"),
                _ => (new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)), "⏰", "待下课时段通知")
            };

            var win = new Window
            {
                Title = $"{label} - {call.Title}",
                Width = 780,
                Height = 470,
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

            // 传唤名单
            if (call.Type == "summon" && !string.IsNullOrWhiteSpace(call.StudentNames))
            {
                var names = string.Join("、", call.StudentNames.Split(
                    new[] { '\r', '\n', ',', '，' },
                    StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
                stack.Children.Add(new TextBlock
                {
                    Text = $"传唤学生：{names}",
                    FontSize = 18,
                    FontWeight = FontWeight.SemiBold,
                    Foreground = accent,
                    TextWrapping = TextWrapping.Wrap
                });
            }

            // 发送者
            if (!string.IsNullOrWhiteSpace(call.Sender))
            {
                stack.Children.Add(new TextBlock
                {
                    Text = $"发送人：{call.Sender}    {DateTime.Now:HH:mm:ss}",
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
