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
/// 「划课」页直接嵌入课时划消界面（ClassHoursPanelControl），进入控制中心即可直接打卡/划课。
/// 标题栏支持拖动窗口；所有按钮统一圆角样式；左侧导航栏加宽至 220px 保证标签文字完整显示。
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

    // 集控平台表单值：控件每次切换页面时重建，值保存在字符串字段中，
    // 避免同一控件实例重复加入视觉树触发"元素已是另一个元素的逻辑子元素"异常（error.log 崩溃根因）
    private string _platformName = "集控平台";
    private string _platformUrl = "http://192.168.1.100:5000";
    private string _platformUsername = "";
    private string _platformPassword = "";

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

        // ===== 标题栏（与大屏模式一致的蓝色 #4285f4，支持拖动窗口） =====
        var title = new Grid { Background = new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)) };
        title.MouseLeftButtonDown += (_, e) =>
        {
            // 点击按钮不触发拖动，其余标题栏区域可拖动窗口
            if (e.ChangedButton == MouseButton.Left && !IsInsideButton(e.OriginalSource))
                Window.GetWindow(this)?.DragMove();
        };
        title.ColumnDefinitions.Add(new ColumnDefinition());
        title.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var left = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(12, 0, 0, 0) };
        left.Children.Add(new TextBlock
        {
            Text = "控制中心", FontSize = 13, FontWeight = FontWeights.SemiBold,
            Foreground = Brushes.White, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 16, 0)
        });
        // 大屏 / 控制模式切换下拉框（与主窗口标题栏一致的圆角白字样式）
        var modeCombo = new ComboBox
        {
            Style = (Style)Application.Current.FindResource("ModeComboBoxStyle"),
            ItemContainerStyle = (Style)Application.Current.FindResource("ModeComboItemStyle"),
            VerticalAlignment = VerticalAlignment.Center,
            SelectedIndex = 1
        };
        modeCombo.Items.Add(new ComboBoxItem { Content = "大屏模式" });
        modeCombo.Items.Add(new ComboBoxItem { Content = "控制模式" });
        modeCombo.SelectionChanged += (_, _) =>
        {
            if (modeCombo.SelectedIndex == 0) CloseRequested?.Invoke();
        };
        left.Children.Add(modeCombo);
        title.Children.Add(left);

        // 关闭按钮：回到大屏模式，而不是退出整个应用
        var close = RoundedButton("✕", () => CloseRequested?.Invoke(), primary: false, foreground: Brushes.White, width: 46, height: 30, padding: new Thickness(0));
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

        // ===== 主体：左侧导航（220px 加宽）+ 右侧内容 =====
        BuildNavigation();
        var body = new Grid();
        body.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(220) });
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

        // 圆角选项卡样式（App.xaml 中定义，避免 FrameworkElementFactory 在 .NET 10 WPF 的 TargetName 解析问题）
        _navigation.Resources.Add(typeof(TabItem), (Style)Application.Current.FindResource("RoundedTabItemStyle"));

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

    /// <summary>「划课」页：直接嵌入课时划消界面（复用大屏同款 ClassHoursPanelControl），进入控制中心即可直接打卡/划课</summary>
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

    /// <summary>集控平台列表页：每次切换重建表单控件（修复重复添加同一控件导致崩溃的问题）</summary>
    private UIElement PlatformPanel()
    {
        var panel = new StackPanel { Margin = new Thickness(22) };
        panel.Children.Add(new TextBlock { Text = "连接集控平台", FontSize = 20, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 14, 14) });
        var name = new TextBox { Text = _platformName };
        var url = new TextBox { Text = _platformUrl };
        var username = new TextBox { Text = _platformUsername };
        var password = new PasswordBox { Password = _platformPassword };
        AddField(panel, "平台名称", name);
        AddField(panel, "平台地址", url);
        AddField(panel, "用户名", username);
        AddField(panel, "密码", password);
        var connect = RoundedButton("连接并加入列表", async () =>
        {
            _platformName = name.Text;
            _platformUrl = url.Text;
            _platformUsername = username.Text;
            _platformPassword = password.Password;
            await ConnectPlatform(name.Text, url.Text, username.Text, password.Password);
        }, primary: true, width: 150);
        connect.Margin = new Thickness(0, 12, 0, 18);
        panel.Children.Add(connect);
        panel.Children.Add(new ListBox { ItemsSource = _platforms, DisplayMemberPath = nameof(ConnectedPlatform.Name), Height = 180 });
        return panel;
    }

    private async Task ConnectPlatform(string name, string url, string username, string password)
    {
        try
        {
            var service = new RemoteControlService();
            service.SetBaseUrl(url);
            await service.LoginAsync(username, password);
            _platforms.Add(new ConnectedPlatform { Name = string.IsNullOrWhiteSpace(name) ? url : name, Url = url, Service = service });
            _status.Text = $"已连接 {_platforms.Count} 个集控平台";
            await RefreshAll();
        }
        catch (Exception ex) { MessageBox.Show(ex.Message, "连接失败", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    private async Task RefreshAll() { await RefreshDevices(); await RefreshTasks(); }

    private async Task RefreshDevices()
    {
        _devices.Clear();
        // 并行拉取所有平台的设备，避免逐平台串行等待拖慢刷新；单个平台异常只影响该平台
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
        // 并行拉取所有平台的任务，单平台失败不影响其他平台
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

    /// <summary>圆角按钮：primary = 蓝色实心；primary = false = 透明背景（可配前景色，用于标题栏白字按钮）</summary>
    private static Button RoundedButton(string text, Action action,
        bool primary = true, bool disabled = false,
        double width = double.NaN, double height = 32,
        Brush? foreground = null, Thickness padding = default)
    {
        var blue = new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4));
        var blueHover = new SolidColorBrush(Color.FromRgb(0x5b, 0x95, 0xf5));
        var bluePressed = new SolidColorBrush(Color.FromRgb(0x33, 0x67, 0xd6));
        var ghostHover = new SolidColorBrush(Color.FromArgb(0x22, 0xff, 0xff, 0xff));
        var ghostPressed = new SolidColorBrush(Color.FromArgb(0x33, 0xff, 0xff, 0xff));
        var disabledBg = new SolidColorBrush(Color.FromRgb(0xcc, 0xcc, 0xcc));

        var b = new Button
        {
            Content = text,
            Height = height,
            Padding = padding == default ? new Thickness(14, 0, 14, 0) : padding,
            FontSize = 13,
            FontWeight = FontWeights.SemiBold,
            Cursor = disabled ? Cursors.Arrow : Cursors.Hand,
            IsEnabled = !disabled,
            Opacity = disabled ? 0.55 : 1.0,
            HorizontalAlignment = HorizontalAlignment.Left,
            Foreground = foreground ?? (primary ? Brushes.White : blue)
        };
        if (!double.IsNaN(width)) b.Width = width;

        var border = new FrameworkElementFactory(typeof(Border));
        border.SetValue(Border.CornerRadiusProperty, new CornerRadius(8));
        border.SetValue(Border.BackgroundProperty, primary ? blue : Brushes.Transparent);
        var presenter = new FrameworkElementFactory(typeof(ContentPresenter));
        presenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
        presenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
        presenter.SetValue(ContentPresenter.RecognizesAccessKeyProperty, true);
        border.AppendChild(presenter);

        var template = new ControlTemplate(typeof(Button)) { VisualTree = border };
        template.Triggers.Add(new Trigger
        {
            Property = Button.IsMouseOverProperty, Value = true,
            Setters = { new Setter(Border.BackgroundProperty, primary ? blueHover : ghostHover) }
        });
        template.Triggers.Add(new Trigger
        {
            Property = Button.IsPressedProperty, Value = true,
            Setters = { new Setter(Border.BackgroundProperty, primary ? bluePressed : ghostPressed) }
        });
        template.Triggers.Add(new Trigger
        {
            Property = Button.IsEnabledProperty, Value = false,
            Setters = { new Setter(Border.BackgroundProperty, primary ? disabledBg : Brushes.Transparent) }
        });
        b.Template = template;
        b.Click += (_, _) => action();
        return b;
    }

    /// <summary>判断鼠标事件原始来源是否位于按钮内部（避免点击按钮时触发窗口拖动）</summary>
    private static bool IsInsideButton(object? source)
    {
        for (DependencyObject? d = source as DependencyObject; d != null; d = VisualTreeHelper.GetParent(d))
            if (d is Button) return true;
        return false;
    }
}
