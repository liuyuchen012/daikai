using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using System.Text;
using System.Speech.Synthesis;
using AgoraIn.ClassIslandPlugin.Models;

namespace AgoraIn.ClassIslandPlugin.Views;

/// <summary>
/// 呼叫提示窗口（Avalonia，运行在 ClassIsland 2.x 主程序进程内）：
/// 1. 收到呼叫 → 弹出对应颜色「打卡提示框」（应急红 / 传唤橙 / 待下课蓝），
///    附带约 2.6 秒醒目脉冲特效（呼吸光边框 + 顶部色带闪烁）
/// 2. 提示框置顶、常驻、无自动关闭——主界面保持醒目状态，直到用户点击「知道了，关闭」
/// 3. 显示后中文 TTS 朗读（System.Speech，自动选择中文语音，同一呼叫不重复朗读）
/// </summary>
public static class CallWindow
{
    /// <summary>最近一次朗读的话语（避免多窗口/重复轮询时重读）</summary>
    private static string? _lastSpoken;

    private static (Color Accent, string Icon, string Label, Color Dark) ThemeFor(string type) => type switch
    {
        "emergency" => (Color.FromRgb(0xe5, 0x39, 0x35), "🚨", "上课应急通知", Color.FromRgb(0x8a, 0x1c, 0x1c)),
        "summon" => (Color.FromRgb(0xfb, 0x8c, 0x00), "📢", "下课传唤", Color.FromRgb(0x9a, 0x52, 0x00)),
        _ => (Color.FromRgb(0x42, 0x85, 0xf4), "⏰", "待下课时段通知", Color.FromRgb(0x18, 0x3a, 0x7a)),
    };

    public static void Show(CallMessage call)
    {
        Dispatcher.UIThread.Post(() =>
        {
            var (accent, icon, label, dark) = ThemeFor(call.Type);
            var accentBrush = new SolidColorBrush(accent);

            // ── 打卡提示框：同色渐变大字、置顶、无自动关闭，点「知道了，关闭」才消失 ──
            var win = new Window
            {
                Title = $"{label} - {call.Title}",
                Width = 920,
                Height = 560,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Topmost = true,
                CanResize = false,
                FontSize = 16,
                Background = new LinearGradientBrush
                {
                    StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                    EndPoint = new RelativePoint(1, 1, RelativeUnit.Relative),
                    GradientStops =
                    {
                        new GradientStop(dark, 0),
                        new GradientStop(accent, 1),
                    }
                }
            };

            var root = new Grid();
            // 半透明深色底提供对比
            root.Children.Add(new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(0xE6, 0x10, 0x10, 0x16)),
                CornerRadius = new CornerRadius(24),
                Margin = new Thickness(18)
            });

            // 醒目脉冲光边框（特效载体）
            var pulseBorder = new Border
            {
                Background = Brushes.Transparent,
                BorderBrush = new SolidColorBrush(accent),
                BorderThickness = new Thickness(8),
                CornerRadius = new CornerRadius(24),
                Margin = new Thickness(6)
            };
            root.Children.Add(pulseBorder);

            // 顶部色带闪烁条（特效载体）
            var banner = new Border
            {
                Background = new SolidColorBrush(accent),
                Height = 8,
                VerticalAlignment = VerticalAlignment.Top,
                Opacity = 0.9
            };
            root.Children.Add(banner);

            var stack = new StackPanel
            {
                Margin = new Thickness(46),
                Spacing = 14,
                VerticalAlignment = VerticalAlignment.Center
            };

            // 顶部：图标 + 类型标签（大号醒目标题）
            stack.Children.Add(new TextBlock
            {
                Text = $"{icon} {label}",
                FontSize = 34,
                FontWeight = FontWeight.ExtraBold,
                Foreground = accentBrush,
                TextWrapping = TextWrapping.Wrap
            });

            // 标题（超大字号）
            stack.Children.Add(new TextBlock
            {
                Text = call.Title,
                FontSize = 44,
                FontWeight = FontWeight.ExtraBold,
                Foreground = Brushes.White,
                TextWrapping = TextWrapping.Wrap
            });

            // 内容
            if (!string.IsNullOrWhiteSpace(call.Message))
            {
                stack.Children.Add(new TextBlock
                {
                    Text = call.Message,
                    FontSize = 24,
                    TextWrapping = TextWrapping.Wrap,
                    Foreground = new SolidColorBrush(Color.FromRgb(0xf1, 0xf5, 0xf9))
                });
            }

            // 传唤名单（醒目黄色）
            if (call.Type == "summon" && !string.IsNullOrWhiteSpace(call.StudentNames))
            {
                var names = string.Join("、", call.StudentNames.Split(
                    new[] { '\r', '\n', ',', '，' },
                    StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
                stack.Children.Add(new TextBlock
                {
                    Text = $"传唤学生：{names}",
                    FontSize = 28,
                    FontWeight = FontWeight.Bold,
                    Foreground = new SolidColorBrush(Color.FromRgb(0xff, 0xd5, 0x4d)),
                    TextWrapping = TextWrapping.Wrap
                });
            }

            // 发送者
            if (!string.IsNullOrWhiteSpace(call.Sender))
            {
                stack.Children.Add(new TextBlock
                {
                    Text = $"发送人：{call.Sender}    {DateTime.Now:HH:mm:ss}",
                    FontSize = 14,
                    Foreground = new SolidColorBrush(Color.FromRgb(0xc9, 0xd1, 0xd9))
                });
            }

            // 底部：关闭按钮（醒目大按钮；窗口无自动关闭，收到呼叫后主界面保持醒目直到点击）
            var ok = new Button
            {
                Content = "知道了，关闭",
                Width = 220,
                Height = 52,
                FontSize = 18,
                FontWeight = FontWeight.Bold,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            ok.Click += (_, _) => win.Close();
            stack.Children.Add(ok);

            root.Children.Add(stack);
            win.Content = root;
            win.Show();

            // ── 醒目脉冲特效：约 2.6 秒，光边框呼吸 + 色带闪烁（3 次脉冲） ──
            var frames = 0;
            var totalFrames = 17;
            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
            timer.Tick += (_, _) =>
            {
                frames++;
                var t = frames / (double)totalFrames;
                var pulse = Math.Abs(Math.Sin(t * Math.PI * 3));
                pulseBorder.BorderThickness = new Thickness(4 + pulse * 14);
                pulseBorder.Opacity = 0.35 + pulse * 0.65;
                banner.Opacity = 0.25 + pulse * 0.75;
                if (frames >= totalFrames)
                {
                    timer.Stop();
                    pulseBorder.Opacity = 1;
                    pulseBorder.BorderThickness = new Thickness(2);
                    banner.Opacity = 0.9;
                }
            };
            timer.Start();

            SpeakAsync(call, label);
        });
    }

    /// <summary>
    /// 朗读呼叫：拼接「类型、标题、内容、名单」后经 Windows 中文语音合成
    /// SpeakAsync 在语音合成线程执行，不阻塞 UI；同一呼叫不重复朗读
    /// </summary>
    private static void SpeakAsync(CallMessage call, string label)
    {
        var sb = new StringBuilder();
        sb.Append(label).Append('，');
        sb.Append(call.Title);
        if (!string.IsNullOrWhiteSpace(call.Message))
            sb.Append('。').Append(call.Message);
        if (call.Type == "summon" && !string.IsNullOrWhiteSpace(call.StudentNames))
            sb.Append('。').Append("请以下同学到办公室：").Append(call.StudentNames.Replace('\n', '、'));
        var text = sb.ToString();

        if (text == _lastSpoken) return;
        _lastSpoken = text;

        Task.Run(() =>
        {
            try
            {
                using var synth = new SpeechSynthesizer();
                // 优先选用中文语音（如 Microsoft Huihui Desktop），找不到则用默认语音
                var zh = synth.GetInstalledVoices()
                    .FirstOrDefault(v => v.VoiceInfo.Culture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase))
                    ?.VoiceInfo.Name;
                if (!string.IsNullOrEmpty(zh)) synth.SelectVoice(zh);
                synth.Rate = 1;
                synth.Speak(text);
            }
            catch
            {
                // 无 TTS 环境（如未安装语音包）时静默降级：只看弹窗提示
            }
        });
    }
}
