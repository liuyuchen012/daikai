using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using ClassIsland.Core.Abstractions.Controls;
using ClassIsland.Core.Attributes;
using AgoraIn.ClassIslandPlugin.Services;

namespace AgoraIn.ClassIslandPlugin.Views;

/// <summary>
/// AgoraIn 联动插件设置页（注册到 ClassIsland 设置窗口的插件设置页）
/// 说明如何连接集控平台：填写平台地址 / 连接密码 / 设备 UUID，或点击「自动探测」
/// </summary>
[SettingsPageInfo("agorain.classisland.settings", "AgoraIn 联动")]
public class PluginSettingsPage : SettingsPageBase
{
    /// <summary>插件配置目录（由插件入口在启动时注入）</summary>
    public static string ConfigFolder { get; set; } = "";

    private readonly PluginSettings _settings = new();
    private TextBox _serverBox = null!;
    private TextBox _pwdBox = null!;
    private TextBox _uuidBox = null!;
    private TextBlock _statusText = null!;

    public PluginSettingsPage()
    {
        if (!string.IsNullOrEmpty(ConfigFolder))
        {
            var loaded = PluginSettings.Load(ConfigFolder);
            _settings.ServerUrl = loaded.ServerUrl;
            _settings.Password = loaded.Password;
            _settings.DeviceUuid = loaded.DeviceUuid;
            _settings.PrenoticeEnabled = loaded.PrenoticeEnabled;
            _settings.PollIntervalSeconds = loaded.PollIntervalSeconds;
        }

        Content = BuildUi();
    }

    private Control BuildUi()
    {
        var stack = new StackPanel
        {
            Margin = new Thickness(30),
            Spacing = 10,
            MaxWidth = 720,
            HorizontalAlignment = HorizontalAlignment.Left
        };

        // 标题 + 使用说明
        stack.Children.Add(new TextBlock
        {
            Text = "连接集控平台",
            FontSize = 22,
            FontWeight = FontWeight.Bold
        });
        stack.Children.Add(new TextBlock
        {
            Text = "在 AgoraIn 教师端「控制模式」连接集控平台后，本插件即可接收集控平台下发的呼叫" +
                   "（待下课时段通知 / 上课应急通知 / 下课传唤），并在 ClassIsland 上置顶弹窗提醒。\n" +
                   "以下三项与本机 AgoraIn 大屏客户端 config.json 保持一致，点击「自动探测」可自动读取。",
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(0x60, 0x60, 0x60))
        });

        // 服务器地址
        stack.Children.Add(new TextBlock { Text = "集控平台地址", FontWeight = FontWeight.SemiBold });
        _serverBox = new TextBox
        {
            Text = _settings.ServerUrl,
            Watermark = "如 http://192.168.1.100:5250"
        };
        stack.Children.Add(_serverBox);

        // 连接密码
        stack.Children.Add(new TextBlock { Text = "连接密码", FontWeight = FontWeight.SemiBold });
        _pwdBox = new TextBox
        {
            Text = _settings.Password,
            PasswordChar = '●',
            Watermark = "与 config.json 的 ServerPassword 一致"
        };
        stack.Children.Add(_pwdBox);

        // 设备 UUID
        stack.Children.Add(new TextBlock { Text = "设备 UUID", FontWeight = FontWeight.SemiBold });
        _uuidBox = new TextBox
        {
            Text = _settings.DeviceUuid,
            Watermark = "client_uuid.txt 内容（留空则自动生成并注册）"
        };
        stack.Children.Add(_uuidBox);

        // 附加选项
        var prenotice = new CheckBox
        {
            Content = "启用待下课时段通知（结合课表在下课时间前自动提醒）",
            IsChecked = _settings.PrenoticeEnabled,
            Margin = new Thickness(0, 6, 0, 0)
        };
        stack.Children.Add(prenotice);

        var interval = new NumericUpDown
        {
            Value = _settings.PollIntervalSeconds,
            Minimum = 5,
            Maximum = 120,
            Increment = 5,
            FormatString = "0 秒",
            Width = 160
        };
        var intervalRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 8
        };
        intervalRow.Children.Add(new TextBlock
        {
            Text = "呼叫轮询间隔：",
            VerticalAlignment = VerticalAlignment.Center
        });
        intervalRow.Children.Add(interval);
        stack.Children.Add(intervalRow);

        // 按钮区
        var btnRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 10,
            Margin = new Thickness(0, 14, 0, 0)
        };
        var detect = new Button { Content = "自动探测", Width = 110, Height = 36 };
        detect.Click += (_, _) =>
        {
            var s = new PluginSettings();
            PluginSettings.TryAutoDetect(s);
            _serverBox.Text = s.ServerUrl;
            _pwdBox.Text = s.Password;
            _uuidBox.Text = s.DeviceUuid;
            _statusText.Text = string.IsNullOrEmpty(s.ServerUrl)
                ? "未找到本机 AgoraIn 大屏客户端配置，请手动填写。"
                : "已从本机 AgoraIn 大屏客户端自动读取配置。";
        };
        var save = new Button { Content = "保存", Width = 110, Height = 36 };
        save.Click += (_, _) =>
        {
            _settings.ServerUrl = _serverBox.Text?.Trim() ?? "";
            _settings.Password = _pwdBox.Text?.Trim() ?? "";
            _settings.DeviceUuid = _uuidBox.Text?.Trim() ?? "";
            _settings.PrenoticeEnabled = prenotice.IsChecked == true;
            if (interval.Value is decimal v) _settings.PollIntervalSeconds = (int)v;
            if (!string.IsNullOrEmpty(ConfigFolder))
                _settings.Save(ConfigFolder);
            _statusText.Text = "设置已保存，正在重启插件轮询…";
            OnSettingsSaved?.Invoke();
        };
        btnRow.Children.Add(detect);
        btnRow.Children.Add(save);
        stack.Children.Add(btnRow);

        _statusText = new TextBlock
        {
            Text = string.IsNullOrEmpty(_settings.ServerUrl) ? "" : "当前已配置集控平台，正在连接…",
            Foreground = new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xF4)),
            TextWrapping = TextWrapping.Wrap
        };
        stack.Children.Add(_statusText);

        return stack;
    }

    /// <summary>保存设置后触发，供插件重启轮询</summary>
    public static event Action? OnSettingsSaved;
}
