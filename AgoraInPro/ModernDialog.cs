using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace CheckIn.Client;

/// <summary>
/// 现代化对话框工具类，提供自定义无边框、圆角、带阴影的确认框和提示框
/// 所有对话框窗口均享有 DWM 圆角效果
/// </summary>
public static class ModernDialog
{
    /// <summary>
    /// 显示确认对话框（含确定/取消按钮），返回用户选择
    /// </summary>
    /// <param name="message">提示消息内容</param>
    /// <param name="title">对话框标题</param>
    /// <returns>用户点击"确定"返回 true，否则 false</returns>
    public static bool Confirm(string message, string title = "确认")
    {
        var result = false;
        var win = CreateWindow(title, 400, 200);

        var grid = new Grid { Margin = new Thickness(24) };
        grid.RowDefinitions.Add(new RowDefinition());
        grid.RowDefinitions.Add(new RowDefinition());
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        grid.Children.Add(new TextBlock
        {
            Text = title, FontSize = 16, FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)),
            Margin = new Thickness(0, 0, 0, 12), VerticalAlignment = VerticalAlignment.Top
        });

        grid.Children.Add(new TextBlock
        {
            Text = message, FontSize = 14, TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)),
            Margin = new Thickness(0, 0, 0, 16)
        });
        Grid.SetRow(grid.Children[^1], 1);

        var cancelBtn = NewBtn("取消", new SolidColorBrush(Color.FromRgb(0xf0, 0xf0, 0xf0)),
            new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)));
        var okBtn = NewBtn("确定", new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)),
            new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)));

        cancelBtn.Click += (_, _) => { result = false; win.Close(); };
        okBtn.Click += (_, _) => { result = true; win.Close(); };

        var btnPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        btnPanel.Children.Add(cancelBtn);
        btnPanel.Children.Add(okBtn);
        Grid.SetRow(btnPanel, 2);
        grid.Children.Add(btnPanel);

        win.Content = new Border
        {
            CornerRadius = new CornerRadius(12),
            Background = new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe0, 0xe0)),
            BorderThickness = new Thickness(1),
            Child = grid
        };
        win.ShowDialog();
        return result;
    }

    /// <summary>
    /// 显示提示对话框（仅含确定按钮）
    /// </summary>
    /// <param name="message">提示消息内容</param>
    /// <param name="title">对话框标题</param>
    public static void Alert(string message, string title = "提示")
    {
        var win = CreateWindow(title, 380, 170);

        var grid = new Grid { Margin = new Thickness(24) };
        grid.RowDefinitions.Add(new RowDefinition());
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        grid.Children.Add(new TextBlock
        {
            Text = title, FontSize = 16, FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)),
            Margin = new Thickness(0, 0, 0, 12)
        });

        grid.Children.Add(new TextBlock
        {
            Text = message, FontSize = 14, TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)),
            Margin = new Thickness(0, 0, 0, 16)
        });
        Grid.SetRow(grid.Children[^1], 1);

        var okBtn = NewBtn("确定", new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)),
            new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)));
        okBtn.Click += (_, _) => win.Close();

        var btnPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Bottom
        };
        btnPanel.Children.Add(okBtn);

        Grid.SetRow(btnPanel, 2);
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.Children.Add(btnPanel);

        win.Content = new Border
        {
            CornerRadius = new CornerRadius(12),
            Background = new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe0, 0xe0)),
            BorderThickness = new Thickness(1),
            Child = grid
        };
        win.ShowDialog();
    }

    /// <summary>
    /// 显示"发现新版本"对话框：告知最新版本与当前版本，并提供前往下载按钮
    /// </summary>
    /// <param name="latest">最新版本号（如 v2.8.34）</param>
    /// <param name="current">当前版本号</param>
    /// <param name="downloadUrl">下载/发布页地址</param>
    public static void UpdateAvailable(string latest, string current, string downloadUrl)
    {
        var win = CreateWindow("发现新版本", 420, 250);

        var grid = new Grid { Margin = new Thickness(24) };
        grid.RowDefinitions.Add(new RowDefinition());
        grid.RowDefinitions.Add(new RowDefinition());
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        grid.Children.Add(new TextBlock
        {
            Text = "发现新版本", FontSize = 16, FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)),
            Margin = new Thickness(0, 0, 0, 12)
        });

        grid.Children.Add(new TextBlock
        {
            Text = $"最新版本：{latest}\n当前版本：{current}\n\n建议前往下载更新，以获得最新功能与修复。",
            FontSize = 14, TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)),
            Margin = new Thickness(0, 0, 0, 16)
        });
        Grid.SetRow(grid.Children[^1], 1);

        var laterBtn = NewBtn("稍后", new SolidColorBrush(Color.FromRgb(0xf0, 0xf0, 0xf0)),
            new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)));
        var downloadBtn = NewBtn("前往下载", new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)),
            new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)));

        laterBtn.Click += (_, _) => win.Close();
        downloadBtn.Click += (_, _) =>
        {
            try { Process.Start(new ProcessStartInfo(downloadUrl) { UseShellExecute = true }); } catch { }
            win.Close();
        };

        var btnPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        btnPanel.Children.Add(laterBtn);
        btnPanel.Children.Add(downloadBtn);
        Grid.SetRow(btnPanel, 2);
        grid.Children.Add(btnPanel);

        win.Content = new Border
        {
            CornerRadius = new CornerRadius(12),
            Background = new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe0, 0xe0)),
            BorderThickness = new Thickness(1),
            Child = grid
        };
        win.ShowDialog();
    }

    /// <summary>
    /// 显示"集控平台发现新版本"对话框：告知管理员平台有新版本，并提供前往下载按钮
    /// </summary>
    /// <param name="latestVersion">最新版本号（如 v2.8.35）</param>
    /// <param name="downloadUrl">下载/发布页地址</param>
    public static void ServerUpdateAvailable(string latestVersion, string downloadUrl)
    {
        var win = CreateWindow("集控平台更新", 420, 250);

        var grid = new Grid { Margin = new Thickness(24) };
        grid.RowDefinitions.Add(new RowDefinition());
        grid.RowDefinitions.Add(new RowDefinition());
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        grid.Children.Add(new TextBlock
        {
            Text = "集控平台发现新版本", FontSize = 16, FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)),
            Margin = new Thickness(0, 0, 0, 12)
        });

        grid.Children.Add(new TextBlock
        {
            Text = $"检测到集控管理平台有新版本 {latestVersion}。\n建议通知管理员前往下载更新，以获得最新功能与修复。",
            FontSize = 14, TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)),
            Margin = new Thickness(0, 0, 0, 16)
        });
        Grid.SetRow(grid.Children[^1], 1);

        var laterBtn = NewBtn("稍后", new SolidColorBrush(Color.FromRgb(0xf0, 0xf0, 0xf0)),
            new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)));
        var downloadBtn = NewBtn("前往下载", new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)),
            new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)));

        laterBtn.Click += (_, _) => win.Close();
        downloadBtn.Click += (_, _) =>
        {
            try { Process.Start(new ProcessStartInfo(downloadUrl) { UseShellExecute = true }); } catch { }
            win.Close();
        };

        var btnPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        btnPanel.Children.Add(laterBtn);
        btnPanel.Children.Add(downloadBtn);
        Grid.SetRow(btnPanel, 2);
        grid.Children.Add(btnPanel);

        win.Content = new Border
        {
            CornerRadius = new CornerRadius(12),
            Background = new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe0, 0xe0)),
            BorderThickness = new Thickness(1),
            Child = grid
        };
        win.ShowDialog();
    }

    /// <summary>
    /// 创建统一样式的无边框对话框窗口，支持 DWM 圆角
    /// </summary>
    private static Window CreateWindow(string title, int w, int h)
    {
        var win = new Window
        {
            Title = title, Width = w, Height = h,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false,
            WindowStyle = WindowStyle.None, AllowsTransparency = true,
            Background = Brushes.Transparent
            // L2：移除 Topmost = true，弹窗不再全局置顶遮挡其他窗口；
            // ShowDialog 本身已是模态，无需置顶
        };
        try
        {
            var hwnd = new WindowInteropHelper(win).EnsureHandle();
            int p = 2;
            NativeMethods.DwmSetWindowAttribute(hwnd, 33, ref p, sizeof(int));
        }
        catch { }
        return win;
    }

    /// <summary>
    /// 创建统一样式的圆角按钮，含悬停透明度效果
    /// </summary>
    private static Button NewBtn(string text, Brush bg, Brush fg)
    {
        var btn = new Button
        {
            Content = text, Width = 80, Height = 32, Cursor = Cursors.Hand,
            FontSize = 13, Foreground = fg, Background = bg,
            BorderThickness = new Thickness(0), Margin = new Thickness(5, 0, 0, 0)
        };
        var template = new ControlTemplate(typeof(Button));
        var border = new FrameworkElementFactory(typeof(Border));
        border.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
        border.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
        var presenter = new FrameworkElementFactory(typeof(ContentPresenter));
        presenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
        presenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
        border.AppendChild(presenter);
        template.VisualTree = border;
        var trigger = new Trigger { Property = Button.IsMouseOverProperty, Value = true };
        trigger.Setters.Add(new Setter(Border.OpacityProperty, 0.85));
        template.Triggers.Add(trigger);
        btn.Template = template;
        return btn;
    }
}

/// <summary>
/// 原生 Win32 API 封装，用于调用 DWM（桌面窗口管理器）设置窗口属性
/// </summary>
internal static class NativeMethods
{
    /// <summary>
    /// 设置 DWM 窗口属性（如圆角效果）
    /// </summary>
    [System.Runtime.InteropServices.DllImport("dwmapi.dll", PreserveSig = true)]
    internal static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
