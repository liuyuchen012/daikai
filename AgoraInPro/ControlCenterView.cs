using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using CheckIn.Client.Models;
using CheckIn.Client.Services;

namespace CheckIn.Client;

/// <summary>
/// 控制中心（控制模式）：UI 风格与大屏模式保持一致（蓝色标题栏 + 统一状态栏与选中高亮）。
/// 「划课」页直接嵌入课时划消界面（ClassHoursPanelControl），不再弹出独立窗口，
/// 修复「打开划课中心」按钮点击闪退问题，实现进入控制中心即可直接打卡/划课。
/// </summary>
public sealed class ControlCenterView : Grid
{
    private readonly ObservableCollection<ConnectedPlatform> _platforms = new();
    private readonly ObservableCollection<AggregatedDevice> _devices = new();
    private readonly ObservableCollection<RemoteTask> _tasks = new();
    private readonly TextBlock _totalDevices = StatValue();
    private readonly TextBlock _totalTasks = StatValue();
    private readonly TextBlock _onlineDevices = StatValue();
    private readonly TextBlock _status = new() { Foreground = Brushes.Gray, Margin = new Thickness(12, 0, 0, 0) };
    private readonly TextBox _name = new() { Text = "集控平台" };
    private readonly TextBox _url = new() { Text = "http://192.168.1.100:5000" };
    private readonly TextBox _username = new();
    private readonly PasswordBox _password = new();
    private readonly TabControl _navigation = new();
    private readonly ContentControl _content = new();
    private ClassHoursPanelControl? _hoursControl;

    /// <summary>请求切回大屏模式（由 MainWindow 订阅）</summary>
    public event Action? CloseRequested;

    public ControlCenterView()
    {
        Background = Brushes.White;
        RowDefinitions.Add(new RowDefinition { Height = new GridLength(48) });
        RowDefinitions.Add(new RowDefinition { Height = new GridLength(36) });
        RowDefinitions.Add(new RowDefinition());
        RowDefinitions.Add(new RowDefinition { Height = new GridLength(28) });

        // ===== 标题栏（与大屏模式一致的蓝色 #4285f4） =====
        var title = new Grid { Background = new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)) };
        title.ColumnDefinitions.Add(new ColumnDefinition());
        title.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var left = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(12, 0, 0, 0) };
        left.Children.Add(new TextBlock
        {
            Text = "控制中心", FontSize = 13, FontWeight = FontWeights.SemiBold,
            Foreground = Brushes.White, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 16, 0)
        });
        left.Children.Add(SceneButton("大屏模式", () => CloseRequested?.Invoke()));
        left.Children.Add(SceneButton("控制模式（当前）", () => { }, disabled: true));
        title.Children.Add(left);

        // 关闭按钮：回到大屏模式，而不是退出整个应用（修复控制模式 UI 问题）
        var close = new Button
        {
            Content = "✕", Width = 46, Height = 30, Foreground = Brushes.White,
            Background = Brushes.Transparent, BorderThickness = new Thickness(0),
            Cursor = Cursors.Hand, FontSize = 14
        };
        close.Click += (_, _) => CloseRequested?.Invoke();
        Grid.SetColumn(close, 1);
        title.Children.Add(close);
        SetRow(title, 0);
        Children.Add(title);

        // ===== 统计条 =====
        var stats = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Background = new SolidColorBrush(Color.FromRgb(0xf5, 0xf7, 0xfa)),
            VerticalAlignment = VerticalAlignment.Center
        };
        stats.Children.Add(Stat("总设备数量", _totalDevices));
        stats.Children.Add(Stat("总任务数量", _totalTasks));
        stats.Children.Add(Stat("在线设备数量", _onlineDevices));
        SetRow(stats, 1);
        Children.Add(stats);

        // ===== 主体：左侧导航 + 右侧内容 =====
        BuildNavigation();
        var body = new Grid();
        body.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
        body.ColumnDefinitions.Add(new ColumnDefinition());
        Grid.SetColumn(_navigation, 0);
        Grid.SetColumn(_content, 1);
        body.Children.Add(_navigation);
        body.Children.Add(_content);
        SetRow(body, 2);
        Children.Add(body);

        SetRow(_status, 3);
        Children.Add(_status);
    }

    private void BuildNavigation()
    {
        _navigation.Background = Brushes.White;
        _navigation.BorderThickness = new Thickness(0, 0, 1, 0);
        _navigation.TabStripPlacement = Dock.Left;

        // 与大屏标签栏一致的选中高亮（浅蓝背景 + 蓝字 + 半粗体）
        var tabStyle = new Style(typeof(TabItem))
        {
            Setters =
            {
                new Setter(Control.HeightProperty, 42.0),
                new Setter(Control.FontSizeProperty, 13.0),
                new Setter(Control.PaddingProperty, new Thickness(14, 0, 0, 0)),
                new Setter(Control.ForegroundProperty, Brushes.Gray),
                new Setter(Control.CursorProperty, Cursors.Hand),
                new Setter(Control.BackgroundProperty, Brushes.Transparent)
            },
            Triggers =
            {
                new Trigger
                {
                    Property = TabItem.IsSelectedProperty, Value = true,
                    Setters =
                    {
                        new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(0xe8, 0xf0, 0xfe))),
                        new Setter(Control.ForegroundProperty, new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4))),
                        new Setter(Control.FontWeightProperty, FontWeights.SemiBold)
                    }
                },
                new Trigger
                {
                    Property = TabItem.IsMouseOverProperty, Value = true,
                    Setters = { new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(0xf5, 0xf7, 0xfa))) }
                }
            }
        };
        _navigation.Resources.Add(typeof(TabItem), tabStyle);

        _navigation.SelectionChanged += (_, _) =>
        {
            if (_navigation.SelectedItem is TabItem item)
                _content.Content = item.Tag switch
                {
                    "hours" => HoursPanel(),
                    "devices" => DevicePanel(),
                    "tasks" => TaskPanel(),
                    _ => PlatformPanel()
                };
        };

        foreach (var item in new[] { ("划课", "hours"), ("设备列表", "devices"), ("任务中心", "tasks"), ("集控平台列表", "platforms") })
            _navigation.Items.Add(new TabItem { Header = item.Item1, Tag = item.Item2, Content = new Grid() });
        _navigation.SelectedIndex = 0;
    }

    /// <summary>
    /// 「划课」页：直接嵌入课时划消界面（复用大屏同款 ClassHoursPanelControl）。
    /// 不再显示「打开划课中心」按钮、不再弹出独立 ClassHoursWindow ——
    /// 修复划课按钮点击闪退问题，进入控制中心即可直接打卡/划课。
    /// </summary>
    private UIElement HoursPanel()
    {
        _hoursControl ??= new ClassHoursPanelControl();
        return _hoursControl;
    }

    private UIElement DevicePanel()
    {
        var grid = new Grid { Margin = new Thickness(18) };
        var list = new DataGrid { ItemsSource = _devices, AutoGenerateColumns = false, IsReadOnly = true };
        list.Columns.Add(new DataGridTextColumn { Header = "集控平台", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Platform)), Width = 150 });
        list.Columns.Add(new DataGridTextColumn { Header = "设备名称", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Name)), Width = 220 });
        list.Columns.Add(new DataGridTextColumn { Header = "状态", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Status)), Width = 90 });
        list.Columns.Add(new DataGridTextColumn { Header = "UUID", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Uuid)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        grid.Children.Add(list);
        _ = RefreshDevices();
        return grid;
    }

    private UIElement TaskPanel()
    {
        var list = new ListView { ItemsSource = _tasks, Margin = new Thickness(18) };
        var taskText = new FrameworkElementFactory(typeof(TextBlock));
        taskText.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding(nameof(RemoteTask.Subject)));
        list.ItemTemplate = new DataTemplate { VisualTree = taskText };
        _ = RefreshTasks();
        return list;
    }

    private UIElement PlatformPanel()
    {
        var panel = new StackPanel { Margin = new Thickness(22) };
        panel.Children.Add(new TextBlock { Text = "连接集控平台", FontSize = 20, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 14) });
        AddField(panel, "平台名称", _name);
        AddField(panel, "平台地址", _url);
        AddField(panel, "用户名", _username);
        AddField(panel, "密码", _password);
        var connect = new Button { Content = "连接并加入列表", Width = 140, Height = 34, HorizontalAlignment = HorizontalAlignment.Left, Margin = new Thickness(0, 12, 0, 18) };
        connect.Click += async (_, _) => await ConnectPlatform();
        panel.Children.Add(connect);
        panel.Children.Add(new ListBox { ItemsSource = _platforms, DisplayMemberPath = nameof(ConnectedPlatform.Name), Height = 180 });
        return panel;
    }

    private async Task ConnectPlatform()
    {
        try
        {
            var service = new RemoteControlService();
            service.SetBaseUrl(_url.Text.Trim());
            await service.LoginAsync(_username.Text.Trim(), _password.Password);
            _platforms.Add(new ConnectedPlatform { Name = string.IsNullOrWhiteSpace(_name.Text) ? _url.Text.Trim() : _name.Text.Trim(), Url = _url.Text.Trim(), Service = service });
            _status.Text = $"已连接 {_platforms.Count} 个集控平台";
            await RefreshAll();
        }
        catch (Exception ex) { MessageBox.Show(ex.Message, "连接失败", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    private async Task RefreshAll() { await RefreshDevices(); await RefreshTasks(); }

    private async Task RefreshDevices()
    {
        _devices.Clear();
        // M4：并行拉取所有平台的设备，避免逐平台串行等待拖慢刷新；
        // 单个平台异常只影响该平台，不影响其他平台的统计
        var snapshot = _platforms.ToList();
        var results = await Task.WhenAll(snapshot.Select(async platform =>
        {
            try
            {
                var devices = await platform.Service.GetDevicesAsync();
                return (platform, devices, failed: false);
            }
            catch
            {
                return (platform, devices: new List<DeviceItem>(), failed: true);
            }
        }));
        foreach (var (platform, devices, _) in results)
            foreach (var device in devices)
                _devices.Add(new AggregatedDevice { Platform = platform.Name, Name = device.Name, Uuid = device.Uuid, Online = device.Online });
        _totalDevices.Text = _devices.Count.ToString();
        _onlineDevices.Text = _devices.Count(d => d.Online).ToString();
        _status.Text = $"已连接 {_platforms.Count} 个平台，共 {_devices.Count} 台设备";
    }

    private async Task RefreshTasks()
    {
        _tasks.Clear();
        // M4：并行拉取所有平台的任务，单平台失败不影响其他平台
        var snapshot = _platforms.ToList();
        var results = await Task.WhenAll(snapshot.Select(async platform =>
        {
            try { return await platform.Service.GetTasksAsync(); }
            catch { return new List<RemoteTask>(); }
        }));
        foreach (var tasks in results)
            foreach (var task in tasks) _tasks.Add(task);
        _totalTasks.Text = _tasks.Count.ToString();
    }

    private static void AddField(Panel panel, string label, Control control) { control.Height = 30; control.Margin = new Thickness(0, 0, 0, 8); panel.Children.Add(new TextBlock { Text = label, FontSize = 12, Foreground = Brushes.Gray }); panel.Children.Add(control); }
    private static TextBlock StatValue() => new() { Text = "0", FontSize = 18, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)) };
    private static UIElement Stat(string label, TextBlock value) { var p = new StackPanel { Margin = new Thickness(20, 0, 30, 0) }; p.Children.Add(new TextBlock { Text = label, FontSize = 11, Foreground = Brushes.Gray }); p.Children.Add(value); return p; }
    private static Button SceneButton(string text, Action action, bool disabled = false)
    {
        var b = new Button
        {
            Content = text, Height = 30, Padding = new Thickness(10, 0, 10, 0),
            Foreground = Brushes.White, Background = Brushes.Transparent, BorderThickness = new Thickness(0),
            Cursor = disabled ? Cursors.Arrow : Cursors.Hand,
            IsEnabled = !disabled, Opacity = disabled ? 0.65 : 1.0
        };
        b.Click += (_, _) => action();
        return b;
    }
}
