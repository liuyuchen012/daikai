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
/// 呼叫提示（Avalonia，运行在 ClassIsland 2.x 主程序进程内，v2.3.0.0）：
/// 1. 呼叫到达 → 在 CLASSISLAND 主界面上播放颜色提示动画（叠加在主窗口顶层，按类型配色：
///    应急红 / 传唤橙 / 待下课蓝），约 3.2 秒：颜色呼吸 + 中心波纹扩散 + 顶部横幅滑入；
/// 2. 动画结束后 → 弹出消息框（同色渐变大窗、置顶、常驻、无自动关闭），
///    主界面保持提醒状态直到用户点击「知道了，关闭」；
/// 3. 消息框出现时中文 TTS 朗读（System.Speech，自动选择中文语音，同一呼叫不重复朗读）。
/// </summary>
public static class CallWindow
{
    /// <summary>最近一次朗读的话语（避免多窗口/重复轮询时重读）</summary>
    private static string? _lastSpoken;

    /// <summary>主界面动画播放锁：一次只播一个动画，避免多屏幕叠加</summary>
    private static readonly object _fxLock = new();
    private static bool _fxPlaying;

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

            // ① ClassIsland 主界面提示动画（按类型颜色）
            PlayMainOverlayFx(call, accent, icon, label, dark);

            // ② 动画结束后弹消息框（常驻，点击才关闭）并朗读
            var delay = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(3400) };
            delay.Tick += (_, _) =>
            {
                delay.Stop();
                ShowMessageBox(call, accent, icon, label, dark);
            };
            delay.Start();
        });
    }

    /// <summary>
    /// 在 ClassIsland 主窗口 OverlayLayer 上叠加颜色提示动画：
    /// 颜色呼吸遮罩 + 中心波纹扩散 + 顶部横幅滑入（约 3.2 秒后自动移除）
    /// </summary>
    private static void PlayMainOverlayFx(CallMessage call, Color accent, string icon, string label, Color dark)
    {
        if (_fxPlaying) return;
        Window? main = GetMainWindow();
        if (main == null) return;
        _fxPlaying = true;

        var panel = new Panel { IsHitTestVisible = false };

        // ① 颜色呼吸遮罩
        var veil = new Border
        {
            Background = new SolidColorBrush(accent),
            Opacity = 0,
            CornerRadius = new CornerRadius(0)
        };
        panel.Children.Add(veil);

        // ② 中心波纹（扩散圆环 + 中心光斑）
        var ring = new Avalonia.Controls.Shapes.Ellipse
        {
            Stroke = new SolidColorBrush(accent),
            StrokeThickness = 10,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
        panel.Children.Add(ring);
        var glow = new Avalonia.Controls.Shapes.Ellipse
        {
            Fill = new SolidColorBrush(Color.FromArgb(0x55, accent.R, accent.G, accent.B)),
            Width = 200,
            Height = 200,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Opacity = 1
        };
        panel.Children.Add(glow);

        // ③ 顶部横幅（从上方滑入：类型 + 标题，白字彩底）
        var bannerTranslate = new TranslateTransform { Y = -92 };
        var banner = new Border
        {
            Background = new LinearGradientBrush
            {
                StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                EndPoint = new RelativePoint(1, 0, RelativeUnit.Relative),
                GradientStops =
                {
                    new GradientStop(dark, 0),
                    new GradientStop(accent, 1),
                }
            },
            Height = 92,
            VerticalAlignment = VerticalAlignment.Top,
            RenderTransform = bannerTranslate
        };
        var bannerStack = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 14, Margin = new Thickness(28), VerticalAlignment = VerticalAlignment.Center };
        bannerStack.Children.Add(new TextBlock { Text = icon, FontSize = 30, VerticalAlignment = VerticalAlignment.Center });
        var bannerText = new StackPanel { VerticalAlignment = VerticalAlignment.Center, Spacing = 2 };
        bannerText.Children.Add(new TextBlock { Text = label, FontSize = 15, FontWeight = FontWeight.Bold, Foreground = Brushes.White });
        bannerText.Children.Add(new TextBlock { Text = call.Title, FontSize = 24, FontWeight = FontWeight.ExtraBold, Foreground = Brushes.White, TextWrapping = TextWrapping.Wrap });
        bannerStack.Children.Add(bannerText);
        banner.Child = bannerStack;
        panel.Children.Add(banner);

        // 优先叠加到主窗口内容（Grid 类布局上铺满且 zIndex 置顶；
        // 非 Grid 布局时用独立覆盖窗口兜底，同样盖在主界面之上）
        // 后添加的 Grid 子项绘制在上层，无需显式 ZIndex
        var attached = false;
        Grid? mainGrid = main.Content as Grid;
        Canvas? mainCanvas = main.Content is Canvas c ? c : null;
        if (mainGrid != null)
        {
            mainGrid.Children.Add(panel);
            attached = true;
        }
        else if (mainCanvas != null)
        {
            mainCanvas.Children.Add(panel);
            panel.Width = main.Bounds.Width;
            panel.Height = main.Bounds.Height;
            attached = true;
        }

        Window? fxHost = main;
        if (!attached)
        {
            fxHost = new Window
            {
                SystemDecorations = SystemDecorations.None,
                ShowInTaskbar = false,
                IsHitTestVisible = false,
                ShowActivated = false,
                Topmost = true,
                Width = main.Width,
                Height = main.Height,
                Content = panel
            };
            fxHost.Show();
        }

        // 动画：22 帧 × 150ms ≈ 3.3s
        var frames = 0;
        const int totalFrames = 22;
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
        timer.Tick += (_, _) =>
        {
            frames++;
            var t = frames / (double)totalFrames;
            var pulse = Math.Abs(Math.Sin(t * Math.PI * 3));

            veil.Opacity = 0.08 + pulse * 0.22;
            var maxSize = Math.Max(main.Bounds.Width, main.Bounds.Height) * 0.95;
            var size = 140 + t * maxSize;
            ring.Width = size;
            ring.Height = size;
            ring.StrokeThickness = 6 + (1 - t) * 20;
            ring.Opacity = Math.Max(0.12, pulse) * 0.95;
            glow.Opacity = 0.25 + pulse * 0.75;
            glow.Width = 180 + pulse * 140;
            glow.Height = 180 + pulse * 140;

            var slideIn = t < 0.25 ? t / 0.25 : 1.0;
            bannerTranslate.Y = -92 * (1 - slideIn);
            if (frames >= totalFrames)
            {
                timer.Stop();
                if (attached)
                {
                    if (mainGrid != null) mainGrid.Children.Remove(panel);
                    if (mainGrid == null && mainCanvas != null) mainCanvas.Children.Remove(panel);
                }
                else
                {
                    fxHost?.Close();
                }
                _fxPlaying = false;
            }
        };
        timer.Start();
    }

    /// <summary>
    /// 获取 ClassIsland 主窗口：优先 AppBase.MainWindow（反射），
    /// 兜底取进程内可见且面积最大的窗口（此时尚无插件弹窗，即主窗口）
    /// </summary>
    private static Window? GetMainWindow()
    {
        try
        {
            var app = ClassIsland.Core.AppBase.Current;
            var mw = app?.GetType().GetProperty("MainWindow")?.GetValue(app) as Window;
            if (mw != null) return mw;
        }
        catch { }
        // 兜底：ClassIsland.Core.AppBase 的 MainWindow（反射各种属性名）
        try
        {
            var app = ClassIsland.Core.AppBase.Current;
            foreach (var propName in new[] { "MainWindow", "Window", "Current" })
            {
                if (app?.GetType().GetProperty(propName)?.GetValue(app) is Window w && w.IsVisible)
                    return w;
            }
        }
        catch { }
        return null;
    }

    /// <summary>常驻消息框：同色渐变大窗、置顶、无自动关闭，点「知道了，关闭」才消失</summary>
    private static void ShowMessageBox(CallMessage call, Color accent, string icon, string label, Color dark)
    {
        var accentBrush = new SolidColorBrush(accent);
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
        root.Children.Add(new Border
        {
            Background = new SolidColorBrush(Color.FromArgb(0xE6, 0x10, 0x10, 0x16)),
            CornerRadius = new CornerRadius(24),
            Margin = new Thickness(18)
        });

        var pulseBorder = new Border
        {
            Background = Brushes.Transparent,
            BorderBrush = new SolidColorBrush(accent),
            BorderThickness = new Thickness(8),
            CornerRadius = new CornerRadius(24),
            Margin = new Thickness(6)
        };
        root.Children.Add(pulseBorder);

        var stack = new StackPanel
        {
            Margin = new Thickness(46),
            Spacing = 14,
            VerticalAlignment = VerticalAlignment.Center
        };

        stack.Children.Add(new TextBlock
        {
            Text = $"{icon} {label}",
            FontSize = 34,
            FontWeight = FontWeight.ExtraBold,
            Foreground = accentBrush,
            TextWrapping = TextWrapping.Wrap
        });
        stack.Children.Add(new TextBlock
        {
            Text = call.Title,
            FontSize = 44,
            FontWeight = FontWeight.ExtraBold,
            Foreground = Brushes.White,
            TextWrapping = TextWrapping.Wrap
        });
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
        if (!string.IsNullOrWhiteSpace(call.Sender))
        {
            stack.Children.Add(new TextBlock
            {
                Text = $"发送人：{call.Sender}    {DateTime.Now:HH:mm:ss}",
                FontSize = 14,
                Foreground = new SolidColorBrush(Color.FromRgb(0xc9, 0xd1, 0xd9))
            });
        }

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

        // 消息框闪光（约 2.6 秒）
        var frames = 0;
        const int totalFrames = 17;
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
        timer.Tick += (_, _) =>
        {
            frames++;
            var t = frames / (double)totalFrames;
            var pulse = Math.Abs(Math.Sin(t * Math.PI * 3));
            pulseBorder.BorderThickness = new Thickness(4 + pulse * 14);
            pulseBorder.Opacity = 0.35 + pulse * 0.65;
            if (frames >= totalFrames)
            {
                timer.Stop();
                pulseBorder.Opacity = 1;
                pulseBorder.BorderThickness = new Thickness(2);
            }
        };
        timer.Start();

        SpeakAsync(call, label);
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
                var zh = synth.GetInstalledVoices()
                    .FirstOrDefault(v => v.VoiceInfo.Culture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase))
                    ?.VoiceInfo.Name;
                if (!string.IsNullOrEmpty(zh)) synth.SelectVoice(zh);
                synth.Rate = 1;
                synth.Speak(text);
            }
            catch
            {
                // 无 TTS 环境时静默降级：只看弹窗提示
            }
        });
    }
}
