using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using CallPlugin.Services;
using ClassIsland.Core.Abstractions.Controls;
using ClassIsland.Core.Attributes;

namespace CallPlugin.Views;

/// <summary>
/// 插件设置页（ClassIsland 设置窗口 → 插件设置）
/// 配置：服务端地址、设备名称、房间、轮询间隔、朗读开关、重新注册
/// </summary>
[SettingsPageInfo("classcall.settings", "班级呼叫系统")]
public class PluginSettingsPage : SettingsPageBase
{
    /// <summary>插件配置目录（由插件入口注入）</summary>
    public static string ConfigFolder { get; set; } = "";

    /// <summary>设置保存后触发（用于重启轮询）</summary>
    public static event Action? OnSettingsSaved;

    private readonly PluginSettings _settings;
    private readonly TextBox _serverBox = new() { Width = 380 };
    private readonly TextBox _nameBox = new() { Width = 380 };
    private readonly TextBox _roomBox = new() { Width = 380 };
    private readonly NumericUpDown _intervalBox = new() { Minimum = 3, Maximum = 300, Width = 120 };
    private readonly CheckBox _speechBox = new() { Content = "启用朗读（收到呼叫后用 ClassIsland 朗读内容）" };
    private readonly TextBlock _statusText = new() { Foreground = Brushes.Gray, TextWrapping = TextWrapping.Wrap };

    public PluginSettingsPage()
    {
        _settings = PluginSettings.Load(ConfigFolder);
        InitializeControls();
    }

    private void InitializeControls()
    {
        var panel = new StackPanel
        {
            Margin = new Thickness(16),
            Spacing = 10,
            MaxWidth = 520
        };

        panel.Children.Add(new TextBlock
        {
            Text = "班级呼叫系统（被控端）",
            FontSize = 20,
            FontWeight = FontWeight.Bold
        });
        panel.Children.Add(new TextBlock
        {
            Text = "在插件配置目录配置服务器地址后，本机即可接收呼出端（控制端）发起的班级呼叫并朗读。",
            FontSize = 12,
            Foreground = Brushes.Gray,
            TextWrapping = TextWrapping.Wrap
        });

        AddField(panel, "服务端地址", _serverBox,
            "呼出端与服务端共用的服务器地址，例如 http://192.168.1.100:5260");
        AddField(panel, "设备名称", _nameBox,
            "用于区分本设备，建议填写教室名，如「301 教室」；留空则使用计算机名");
        AddField(panel, "所属房间", _roomBox, "可选，仅用于展示");

        var intervalRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        intervalRow.Children.Add(new TextBlock { Text = "轮询间隔（秒）", VerticalAlignment = VerticalAlignment.Center });
        intervalRow.Children.Add(_intervalBox);
        panel.Children.Add(intervalRow);

        panel.Children.Add(_speechBox);

        var saveBtn = new Button { Content = "保存设置", Width = 110, Height = 32 };
        saveBtn.Click += (_, _) => Save();
        panel.Children.Add(saveBtn);

        var registerBtn = new Button { Content = "重新注册设备", Width = 110, Height = 32 };
        registerBtn.Click += (_, _) => ReRegister();
        panel.Children.Add(registerBtn);

        panel.Children.Add(_statusText);

        // 回填当前配置
        _serverBox.Text = _settings.ServerUrl;
        _nameBox.Text = _settings.DeviceName;
        _roomBox.Text = _settings.Room;
        _intervalBox.Value = _settings.PollIntervalSeconds;
        _speechBox.IsChecked = _settings.SpeechEnabled;
        _statusText.Text = _settings.IsConfigured
            ? $"当前状态：已注册（UUID：{_settings.DeviceUuid}）"
            : "当前状态：未配置。填写服务端地址并保存后会自动注册。";

        Content = new ScrollViewer { Content = panel };
    }

    private void AddField(StackPanel panel, string label, Control input, string hint)
    {
        panel.Children.Add(new TextBlock { Text = label, FontWeight = FontWeight.SemiBold });
        panel.Children.Add(input);
        if (!string.IsNullOrEmpty(hint))
            panel.Children.Add(new TextBlock { Text = hint, FontSize = 11, Foreground = Brushes.Gray, TextWrapping = TextWrapping.Wrap });
    }

    private void Save()
    {
        _settings.ServerUrl = _serverBox.Text?.Trim() ?? "";
        _settings.DeviceName = _nameBox.Text?.Trim() ?? "";
        _settings.Room = _roomBox.Text?.Trim() ?? "";
        _settings.PollIntervalSeconds = (int)(_intervalBox.Value ?? 5);
        _settings.SpeechEnabled = _speechBox.IsChecked ?? true;
        _settings.Save(ConfigFolder);
        _statusText.Text = "已保存，正在重启接收服务…";
        OnSettingsSaved?.Invoke();
        _statusText.Text = _settings.IsConfigured
            ? $"已保存并应用。已注册（UUID：{_settings.DeviceUuid}）"
            : "已保存。填写服务端地址后会自动注册。";
    }

    private void ReRegister()
    {
        _settings.DeviceUuid = "";
        _settings.Password = "";
        _settings.Save(ConfigFolder);
        _statusText.Text = "已清除注册信息，正在重新注册…";
        OnSettingsSaved?.Invoke();
    }
}
