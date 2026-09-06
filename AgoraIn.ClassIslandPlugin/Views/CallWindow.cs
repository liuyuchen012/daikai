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

    private static (Color Accent, string Icon, string Label, Color Dark) ThemeFor(string type) => type switch
    {
        "emergency" => (Color.FromRgb(0xe5, 0x39, 0x35), "🚨", "上课应急通知", Color.FromRgb(0x8a, 0x1c, 0x1c)),
        "summon" => (Color.FromRgb(0xfb, 0x8c, 0x00), "📢", "下课传唤", Color.FromRgb(0x9a, 0x52, 0x00)),
        _ => (Color.FromRgb(0x42, 0x85, 0xf4), "⏰", "待下课时段通知", Color.FromRgb(0x18, 0x3a, 0x7a)),
    };

    public static void Show(CallMessage call)
    {
        try { Console.WriteLine("[AgoraIn] SHOW-ENTER " + call.Type); } catch { }
        Dispatcher.UIThread.Post(() =>
        {
            Color accent, dark; string icon, label;
            try
            {
                (accent, icon, label, dark) = ThemeFor(call.Type);
            }
            catch (Exception ex) { LogEx("Theme", ex); return; }

            try { Console.WriteLine("[AgoraIn] POST-RUN"); } catch { }
            // ① 顶部提示栏（WinForms 置顶条，15 秒倒计时自动关闭）
            try { ShowStickyBanner(call, accent, icon, label, dark); }
            catch (Exception ex) { LogEx("ShowStickyBanner", ex); }

            // ② 中文朗读一遍
            try { SpeakOnce(call, label); }
            catch (Exception ex) { LogEx("SpeakOnce", ex); }
        });
    }

    /// <summary>
    /// 顶部提示栏（WinForms 置顶条）：
    /// 单行展示「类型 · 标题 · 内容」，右侧倒计时「N 秒后自动关闭」，到 0 自动关闭。同时只保留最新一条。
    /// </summary>
    private static System.Windows.Forms.Form? _sticky;

    private static void ShowStickyBanner(CallMessage call, Color accent, string icon, string label, Color dark)
    {
        try
        {
            // 旧的先关，保持只有最新呼叫横幅
            _sticky?.Close();
            var scr = System.Windows.Forms.Screen.PrimaryScreen;
            if (scr == null) return;
            var f = new System.Windows.Forms.Form
            {
                Text = $"{label} - {call.Title}（常驻）",
                FormBorderStyle = System.Windows.Forms.FormBorderStyle.None,
                StartPosition = System.Windows.Forms.FormStartPosition.Manual,
                Location = new System.Drawing.Point(scr.Bounds.Left, scr.Bounds.Top),
                Size = new System.Drawing.Size(scr.Bounds.Width, 58),
                TopMost = true,
                ShowInTaskbar = false,
                BackColor = System.Drawing.Color.FromArgb(dark.R, dark.G, dark.B),
            };
            f.Paint += (_, p) =>
            {
                var rect = new System.Drawing.Rectangle(0, 0, f.Width, f.Height);
                using var br = new System.Drawing.Drawing2D.LinearGradientBrush(rect,
                    System.Drawing.Color.FromArgb(dark.R, dark.G, dark.B),
                    System.Drawing.Color.FromArgb(accent.R, accent.G, accent.B), 0f);
                p.Graphics.FillRectangle(br, rect);
                var contentLine = string.IsNullOrWhiteSpace(call.Message)
                    ? (call.Type == "summon" ? $"传唤学生：{call.StudentNames.Replace('\n', '、')}" : "（无附加内容）")
                    : call.Message;
                var full = $"{icon} {label} · {call.Title} · {contentLine}";
                using var font = new System.Drawing.Font("Microsoft YaHei", 15f, System.Drawing.FontStyle.Bold);
                var maxW = f.Width - 260;
                // 单行省略：测量裁剪
                var shown = full;
                while (System.Windows.Forms.TextRenderer.MeasureText(shown, font).Width > maxW && shown.Length > 4)
                    shown = shown.Substring(0, shown.Length - 2);
                if (shown != full) shown = shown + "...";
                System.Windows.Forms.TextRenderer.DrawText(p.Graphics, shown, font,
                    new System.Drawing.Point(28, (f.Height - 26) / 2),
                    System.Drawing.Color.White,
                    System.Windows.Forms.TextFormatFlags.Left |
                    System.Windows.Forms.TextFormatFlags.VerticalCenter |
                    System.Windows.Forms.TextFormatFlags.NoPadding |
                    System.Windows.Forms.TextFormatFlags.NoPrefix);
            };
            var ok = new System.Windows.Forms.Label
            {
                Text = $"{AutoCloseSeconds} 秒后自动关闭",
                AutoSize = false,
                Size = new System.Drawing.Size(150, 58),
                BackColor = System.Drawing.Color.Transparent,
                ForeColor = System.Drawing.Color.White,
                Font = new System.Drawing.Font("Microsoft YaHei", 12f, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            ok.Location = new System.Drawing.Point(f.Width - ok.Width - 24, 0);
            f.Controls.Add(ok);

            // 15 秒倒计时自动关闭
            var remain = AutoCloseSeconds;
            var timer = new System.Windows.Forms.Timer { Interval = 1000 };
            timer.Tick += (_, _) =>
            {
                remain--;
                if (remain <= 0)
                {
                    timer.Stop();
                    f.Close();
                }
                else ok.Text = $"{remain} 秒后自动关闭";
            };
            f.Shown += (_, _) => timer.Start();
            f.FormClosed += (_, _) => { timer.Stop(); if (_sticky == f) _sticky = null; };
            f.Show();
            _sticky = f;
        }
        catch { }
    }

    /// <summary>提示栏自动关闭倒计时（秒）</summary>
    private const int AutoCloseSeconds = 15;

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

                var contentLine = string.IsNullOrWhiteSpace(call.Message)
                    ? (call.Type == "summon" ? call.StudentNames.Replace('\n', '、') : "呼叫内容")
                    : call.Message;
                // 单行：类型 · 标题 · 内容（超长省略号）
                var line1 = new TextBlock
                {
                    Text = $"{icon} {label} · {call.Title} · {contentLine}",
                    FontSize = 14, FontWeight = FontWeight.Bold, Foreground = Brushes.White,
                    Margin = new Thickness(16, 10, 16, 12),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    TextWrapping = TextWrapping.NoWrap,
                    MaxWidth = 360
                };
                var stack = new StackPanel();
                stack.Children.Add(line1);
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

    /// <summary>新窗口消息框：WinForms 原生窗口（不经过 Avalonia，GDI 绘制，任何环境都保证弹出）</summary>
    private static void ShowMessageBox(CallMessage call, Color accent, string icon, string label, Color dark)
    {
        System.Windows.Forms.Form f = new System.Windows.Forms.Form
        {
            Text = $"{label} - {call.Title}",
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.None,
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
            TopMost = true,
            ShowInTaskbar = false,
            Size = new System.Drawing.Size(920, 520),
            BackColor = System.Drawing.Color.FromArgb(dark.R, dark.G, dark.B),
        };
        f.Paint += (_, p) =>
        {
            // 渐变背景（dark -> accent）
            var rect = new System.Drawing.Rectangle(0, 0, f.Width, f.Height);
            using var br = new System.Drawing.Drawing2D.LinearGradientBrush(rect,
                System.Drawing.Color.FromArgb(dark.R, dark.G, dark.B),
                System.Drawing.Color.FromArgb(accent.R, accent.G, accent.B), 35f);
            p.Graphics.FillRectangle(br, rect);
            // 双层深色卡片底
            p.Graphics.FillRectangle(new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(230, 16, 16, 22)),
                new System.Drawing.Rectangle(62, 62, f.Width - 124, f.Height - 124));
            // 单行：类型 · 标题 · 内容
            var contentLine = string.IsNullOrWhiteSpace(call.Message)
                ? (call.Type == "summon" ? $"传唤学生：{call.StudentNames.Replace('\n', '、')}" : "（无附加内容）")
                : call.Message;
            var full = $"{icon} {label} · {call.Title} · {contentLine}";
            using var font = new System.Drawing.Font("Microsoft YaHei", 23f, System.Drawing.FontStyle.Bold);
            var maxW = f.Width - 220;
            // 单行省略：测量裁剪
            var shown = full;
            while (System.Windows.Forms.TextRenderer.MeasureText(shown, font).Width > maxW && shown.Length > 4)
                shown = shown.Substring(0, shown.Length - 2);
            if (shown != full) shown = shown + "...";
            System.Windows.Forms.TextRenderer.DrawText(p.Graphics, shown, font,
                new System.Drawing.Rectangle(90, (f.Height - 60) / 2 - 28, f.Width - 180, 70),
                System.Drawing.Color.White, System.Windows.Forms.TextFormatFlags.Left |
                System.Windows.Forms.TextFormatFlags.VerticalCenter |
                System.Windows.Forms.TextFormatFlags.NoPadding |
                System.Windows.Forms.TextFormatFlags.NoPrefix);
            // 发送人
            using var sf = new System.Drawing.Font("Microsoft YaHei", 10.5f);
            System.Windows.Forms.TextRenderer.DrawText(p.Graphics,
                $"发送人：{call.Sender}    {DateTime.Now:HH:mm:ss}", sf,
                new System.Drawing.Point(95, f.Height - 90), System.Drawing.Color.FromArgb(200, 210, 218));
        };
        var ok = new System.Windows.Forms.Button
        {
            Text = "知道了，关闭",
            Size = new System.Drawing.Size(220, 52),
            BackColor = System.Drawing.Color.White,
            ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
            Font = new System.Drawing.Font("Microsoft YaHei", 13f, System.Drawing.FontStyle.Bold),
            FlatStyle = System.Windows.Forms.FlatStyle.Flat,
            Cursor = System.Windows.Forms.Cursors.Hand
        };
        ok.FlatAppearance.BorderSize = 0;
        ok.Location = new System.Drawing.Point(f.Width - ok.Width - 70, f.Height - ok.Height - 42);
        ok.Click += (_, _) => { f.Close(); };
        f.Controls.Add(ok);
        f.Show();
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

    /// <summary>插件运行日志（%LOCALAPPDATA%\AgoraIn\plugin.log），失败时静默降级</summary>
    private static void LogEx(string stage, Exception ex)
    {
        try
        {
            var dir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AgoraIn");
            System.IO.Directory.CreateDirectory(dir);
            System.IO.File.AppendAllText(System.IO.Path.Combine(dir, "plugin.log"),
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{stage}] {ex}\r\n");
        }
        catch { }
    }
}
