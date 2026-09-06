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
/// 呼叫提示（Avalonia，运行在 ClassIsland 2.x 主程序进程内，v2.5.0）：
/// 1. 「标准提醒」：ClassIsland 风格右上角通知卡片（类型色条 + 两行文本，约 6 秒自动淡出）
/// 2. 「新窗口」：呼叫内容大消息框（双行展示：类型·标题 / 内容·名单），
///    置顶、无自动关闭，用户手动点击「知道了，关闭」才消失
/// 3. 「主界面常驻提示」：主窗口顶部横幅（呼叫类型+标题）在动画结束后常驻显示，
///    直到用户在主界面点击「我知道了」才关闭（动画期间按朗读节奏脉冲 ×3）
/// 4. 朗读：中文 TTS 仅播放一遍（其余两遍节奏无声音，仅界面动画）
/// </summary>
public static class CallWindow
{
    private static string? _lastSpoken;
    private static readonly object _fxLock = new();
    private static bool _fxPlaying;
    private const int RepeatPulse = 3;   // 朗读节奏次数（3 遍：第 1 遍出声，其余无声）
    private const double PulseSeconds = 2.8;

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
            Color accent, dark; string icon, label;
            try
            {
                (accent, icon, label, dark) = ThemeFor(call.Type);
            }
            catch (Exception ex) { LogEx("Theme", ex); return; }

            // ① 标准提醒通知卡片（自动淡出）
            try { StandardNotice.Show(call, accent, icon, label); }
            catch (Exception ex) { LogEx("StandardNotice", ex); }

            // ② 主界面动画（3 次脉冲节奏）→ 动画结束后顶部常驻横幅（直到点「我知道了」）
            try { PlayMainOverlayFx(call, accent, icon, label, dark); }
            catch (Exception ex) { LogEx("PlayMainOverlayFx", ex); }

            // ③ 立即弹出「新窗口」（单行展示，点「知道了，关闭」才消失）
            try { ShowMessageBox(call, accent, icon, label, dark); }
            catch (Exception ex) { LogEx("ShowMessageBox", ex); }

            // ④ 中文朗读仅一遍（其余节奏由主界面脉冲承担）
            try { SpeakOnce(call, label); }
            catch (Exception ex) { LogEx("SpeakOnce", ex); }
        });
    }

    /// <summary>
    /// 主界面效果：
    /// 阶段一（约 3×Pulse）全屏动画：呼吸遮罩 + 波纹扩散 + 顶部横幅滑入，按朗读节奏脉冲 3 次；
    /// 阶段二：动画停止后横幅常驻（呼叫类型 + 标题 + 「我知道了」按钮），直到用户点击关闭。
    /// </summary>
    private static void PlayMainOverlayFx(CallMessage call, Color accent, string icon, string label, Color dark)
    {
        Window? main = GetMainWindow();
        if (main == null) return;
        if (_fxPlaying) return;
        _fxPlaying = true;

        var panel = new Panel { IsHitTestVisible = false };

        // 中心波纹
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
            Width = 200, Height = 200,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Opacity = 1
        };
        panel.Children.Add(glow);

        // 顶部横幅（滑入；动画由内容驱动）
        var bannerTranslate = new TranslateTransform { Y = -64 };
        var banner = BuildStickyBanner(call, accent, icon, label, dark, bannerTranslate, panel);
        panel.Children.Add(banner);

        // 装入主窗（Grid 叠加或覆盖窗兜底）
        var attached = false;
        var fxHost = main;
        if (main.Content is Grid g)
        {
            g.Children.Add(panel);
            attached = true;
        }
        else
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

        // 阶段一：脉冲动画（RepeatPulse 次呼吸），随后进入阶段二常驻
        var totalFrames = (int)(RepeatPulse * PulseSeconds * 6.67); // ≈每脉冲 2.8s
        var frames = 0;
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
        timer.Tick += (_, _) =>
        {
            frames++;
            var total = (double)totalFrames;
            var t = frames / total;
            var pulse = Math.Abs(Math.Sin(t * Math.PI * RepeatPulse));
            var size = 200 + t * Math.Max(main.Width, main.Height) * 0.9;

            ring.Width = size; ring.Height = size;
            ring.StrokeThickness = 6 + (1 - t) * 22;
            ring.Opacity = Math.Max(0.1, pulse) * 0.95;
            glow.Opacity = 0.2 + pulse * 0.75;
            var slideIn = t < 0.12 ? t / 0.12 : 1.0;
            bannerTranslate.Y = -64 * (1 - slideIn);

            if (frames >= totalFrames)
            {
                timer.Stop();
                // 阶段二：动画元素淡出，横幅转为常驻（带「我知道了」按钮）
                ring.Width = 0; ring.Height = 0; ring.Opacity = 0; glow.Opacity = 0;
                bannerTranslate.Y = 0;
                _fxPlaying = false;
                // 横幅已显示按钮，由按钮点击移除 panel（含覆盖窗/主窗子项）
                banner.Opacity = 1;
            }
        };
        timer.Start();

        // 记录容器用于移除
        panel.Tag = new object[] { attached, fxHost, main };
    }

    /// <summary>构建常驻横幅（类型色渐变 + 图标/类型 + 标题 + 「我知道了」按钮）</summary>
    private static Border BuildStickyBanner(CallMessage call, Color accent, string icon, string label, Color dark,
        TranslateTransform translate, Panel hostPanel)
    {
        var banner = new Border
        {
            Height = 58,
            VerticalAlignment = VerticalAlignment.Top,
            RenderTransform = translate,
            Opacity = 0.98,
            Background = new LinearGradientBrush
            {
                StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                EndPoint = new RelativePoint(1, 0, RelativeUnit.Relative),
                GradientStops =
                {
                    new GradientStop(dark, 0),
                    new GradientStop(accent, 1),
                }
            }
        };
        var row = new Grid();
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        row.ColumnDefinitions.Add(new ColumnDefinition());
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var iconBlock = new TextBlock { Text = icon, FontSize = 22, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(16, 0, 10, 0) };
        Grid.SetColumn(iconBlock, 0);
        row.Children.Add(iconBlock);

        var textStack = new StackPanel { VerticalAlignment = VerticalAlignment.Center, Spacing = 2 };
        textStack.Children.Add(new TextBlock
        {
            Text = $"{label}（呼叫类型）",
            FontSize = 11, FontWeight = FontWeight.Bold, Foreground = Brushes.White
        });
        textStack.Children.Add(new TextBlock
        {
            Text = call.Title,
            FontSize = 18, FontWeight = FontWeight.Bold, Foreground = Brushes.White,
            TextTrimming = TextTrimming.CharacterEllipsis,
            MaxWidth = 1100
        });
        Grid.SetColumn(textStack, 1);
        row.Children.Add(textStack);

        var okBtn = new Button
        {
            Content = "我知道了",
            MinWidth = 110, Height = 38, FontSize = 13, FontWeight = FontWeight.Bold,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(12, 0, 22, 0),
            Background = new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)),
            Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)),
            HorizontalAlignment = HorizontalAlignment.Right
        };
        okBtn.Click += (_, _) =>
        {
            if (hostPanel.Tag is object[] bag)
            {
                var attached = (bool)bag[0];
                var fxHost = (Window?)bag[1];
                var main = (Window?)bag[2];
                if (attached)
                {
                    if (main?.Content is Grid g3) g3.Children.Remove(hostPanel);
                    if (main?.Content is Canvas cv3) cv3.Children.Remove(hostPanel);
                }
                else fxHost?.Close();
            }
        };
        Grid.SetColumn(okBtn, 2);
        row.Children.Add(okBtn);

        banner.Child = row;
        return banner;
    }

    /// <summary>标准提醒通知卡片（ClassIsland 风格，约 6 秒后自动淡出）</summary>
    public static class StandardNotice
    {
        public static void Show(CallMessage call, Color accent, string icon, string label)
        {
            bool shown = false;
            var saved = new Border();
            Dispatcher.UIThread.Post(() =>
            {
                if (shown) return;
                shown = true;
                var card = new Border
                {
                    Width = 390,
                    CornerRadius = new CornerRadius(14),
                    ClipToBounds = true,
                    BorderBrush = new SolidColorBrush(Color.FromArgb(0x88, 0xff, 0xff, 0xff)),
                    BorderThickness = new Thickness(1),
                    Background = new SolidColorBrush(Color.FromArgb(0xF2, 0x1d, 0x1f, 0x26)),
                    IsHitTestVisible = false,
                    Opacity = 0
                };
                var grid = new Grid();
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                var colorBar = new Border { Height = 6, Background = new SolidColorBrush(accent), CornerRadius = new CornerRadius(14, 14, 0, 0) };
                Grid.SetRow(colorBar, 0);
                grid.Children.Add(colorBar);

                var line1 = new TextBlock
                {
                    Text = $"{icon} {label} · {call.Title}",
                    FontSize = 16, FontWeight = FontWeight.Bold, Foreground = Brushes.White,
                    Margin = new Thickness(16, 10, 16, 0),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    MaxWidth = 360
                };
                var line2 = new TextBlock
                {
                    Text = string.IsNullOrWhiteSpace(call.Message) ? (call.Type == "summon" ? call.StudentNames.Replace('\n', '、') : "呼叫内容") : call.Message,
                    FontSize = 13,
                    Foreground = new SolidColorBrush(Color.FromRgb(0xd7, 0xde, 0xe6)),
                    Margin = new Thickness(16, 2, 16, 12),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    MaxWidth = 360
                };
                var stack = new StackPanel();
                stack.Children.Add(line1);
                stack.Children.Add(line2);
                Grid.SetRow(stack, 1);
                grid.Children.Add(stack);
                card.Child = grid;

                var host = GetMainWindow();
                if (host == null) return;
                var fx = new Window
                {
                    SystemDecorations = SystemDecorations.None,
                    ShowInTaskbar = false,
                    IsHitTestVisible = false,
                    ShowActivated = false,
                    Topmost = true,
                    Width = 404,
                    Height = 96,
                    Content = card
                };
                var screen = host.Bounds;
                fx.Position = new PixelPoint((int)screen.Width - 420, 84);
                fx.Show();

                // 淡入 → 6 秒 → 淡出
                var ticks = 0;
                const int fadeIn = 4, hold = 40, fadeOut = 4;
                var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
                timer.Tick += (_, _) =>
                {
                    ticks++;
                    if (ticks <= fadeIn) card.Opacity = ticks / (double)fadeIn;
                    else if (ticks <= fadeIn + hold) card.Opacity = 1;
                    else if (ticks <= fadeIn + hold + fadeOut) card.Opacity = 1 - (ticks - fadeIn - hold) / (double)fadeOut;
                    else { timer.Stop(); fx.Close(); }
                };
                timer.Start();
            });
        }
    }

    private static Window? GetMainWindow()
    {
        try
        {
            var app = ClassIsland.Core.AppBase.Current;
            var mw = app?.GetType().GetProperty("MainWindow")?.GetValue(app) as Window;
            if (mw != null) return mw;
        }
        catch { }
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

    /// <summary>新窗口消息框：双行展示（类型·标题 / 内容·名单），点「知道了，关闭」才消失</summary>
    private static void ShowMessageBox(CallMessage call, Color accent, string icon, string label, Color dark)
    {        var accentBrush = new SolidColorBrush(accent);
        var win = new Window
        {
            Title = $"{label} - {call.Title}",
            Width = 920,
            Height = 520,
            WindowStartupLocation = WindowStartupLocation.Manual,
            Topmost = true,
            CanResize = false,
            ShowActivated = false,
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

        var stack = new StackPanel { Margin = new Thickness(46), Spacing = 16, VerticalAlignment = VerticalAlignment.Center };
        // 双行展示：第一行 = 类型·标题；第二行 = 内容（或传唤名单），超长单行省略（不滚动）
        stack.Children.Add(new TextBlock
        {
            Text = $"{icon} {label} — {call.Title}",
            FontSize = 34,
            FontWeight = FontWeight.ExtraBold,
            Foreground = Brushes.White,
            TextTrimming = TextTrimming.CharacterEllipsis,
            TextWrapping = TextWrapping.NoWrap
        });
        var contentLine = string.IsNullOrWhiteSpace(call.Message)
            ? (call.Type == "summon" ? $"传唤学生：{call.StudentNames.Replace('\n', '、')}" : "（无附加内容）")
            : call.Message;
        stack.Children.Add(new TextBlock
        {
            Text = contentLine,
            FontSize = 24,
            FontWeight = FontWeight.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(0xf1, 0xf5, 0xf9)),
            TextTrimming = TextTrimming.CharacterEllipsis,
            TextWrapping = TextWrapping.NoWrap
        });
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
            Width = 220, Height = 52, FontSize = 18, FontWeight = FontWeight.Bold,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        ok.Click += (_, _) => win.Close();
        stack.Children.Add(ok);

        root.Children.Add(stack);
        win.Content = root;
        // 显式定位到主屏中央：必须用屏幕尺寸（主窗口可能是 60px 顶部条模式，
        // 用主窗 Bounds 会把窗口算到屏幕外）
        try
        {
            var scr = win.Screens?.Primary?.Bounds;
            var sw = scr?.Width ?? 1920;
            var sh = scr?.Height ?? 1080;
            win.Position = new PixelPoint((int)((sw - 920) / 2), (int)((sh - 520) / 2));
        }
        catch { }        win.Show();
        // 窗口脉冲（2.6s）
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
    }

    /// <summary>插件异常日志（%LOCALAPPDATA%\AgoraIn\plugin.log）</summary>
    internal static void LogEx(string tag, Exception ex)
    {
        try
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AgoraIn");
            Directory.CreateDirectory(dir);
            File.AppendAllText(Path.Combine(dir, "plugin.log"),
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{tag}] {ex}\r\n\r\n");
        }
        catch { }
    }

    /// <summary>中文朗读仅一遍（语音一次；第二、三遍由主界面脉冲承担）</summary>
    private static void SpeakOnce(CallMessage call, string label)
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
                // 无 TTS 环境时静默降级
            }
        });
    }
}
