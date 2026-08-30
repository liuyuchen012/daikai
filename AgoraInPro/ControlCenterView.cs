using System.Collections.ObjectModel;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
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
/// 「设备列表」内置呼叫面板：勾选设备后可批量发送呼叫（待下课/应急/传唤），也可单台呼叫。
/// 已连接平台持久化到 platforms.json（密码 DPAPI 加密），进入控制中心时自动重连。
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
    private string _platformUrl = DefaultPlatformUrl();
    private string _platformUsername = "";
    private string _platformPassword = "";

    private readonly StackPanel _navigation = new();
    private readonly ContentControl _content = new();
    private ClassHoursPanelControl? _hoursControl;
    private int _selectedNavIndex;

    // 设备/任务空态提示：页面每次切换重建，这里保存最新实例供刷新逻辑更新可见性
    private TextBlock? _deviceHint;
    private TextBlock? _taskHint;

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
            VerticalAlignment = VerticalAlignment.Center
        };
        // 先填充项再设置选中项：SelectedIndex 在空集合上设置无效，会错误地显示第一项「大屏模式」
        modeCombo.Items.Add("大屏模式");
        modeCombo.Items.Add("控制模式");
        modeCombo.SelectedIndex = 1;
        modeCombo.SelectionChanged += (_, _) =>
        {
            if (modeCombo.SelectedIndex == 0) CloseRequested?.Invoke();
        };
        left.Children.Add(modeCombo);
        title.Children.Add(left);

        // 窗口控制按钮：最小化 / 最大化（操作宿主主窗口）/ 关闭（回到大屏模式）
        var minBtn = RoundedButton("—", () =>
        {
            if (Window.GetWindow(this) is Window w) w.WindowState = WindowState.Minimized;
        }, primary: false, foreground: Brushes.White, width: 40, height: 30, padding: new Thickness(0));
        var maxBtn = RoundedButton("□", () =>
        {
            if (Window.GetWindow(this) is Window w)
                w.WindowState = w.WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
        }, primary: false, foreground: Brushes.White, width: 40, height: 30, padding: new Thickness(0));
        var close = RoundedButton("✕", () => CloseRequested?.Invoke(), primary: false, foreground: Brushes.White, width: 46, height: 30, padding: new Thickness(0));
        var right = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 8, 0) };
        right.Children.Add(minBtn);
        right.Children.Add(maxBtn);
        right.Children.Add(close);
        Grid.SetColumn(right, 1);
        title.Children.Add(right);
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

        var navBorder = new Border
        {
            Background = Brushes.White,
            BorderThickness = new Thickness(0, 0, 1, 0),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe0, 0xe0)),
            Child = _navigation
        };

        Grid.SetColumn(navBorder, 0);
        Grid.SetColumn(_content, 1);
        body.Children.Add(navBorder);
        body.Children.Add(_content);
        SetRow(body, 2);
        Children.Add(body);

        SetRow(_status, 3);
        Children.Add(_status);

        // 自动重连上一次连接过的集控平台（platforms.json，密码 DPAPI 加密）
        _ = AutoReconnectAsync();
    }

    /// <summary>默认平台地址：优先取本机 config.json 中已配置的服务器地址，取不到再用占位地址</summary>
    private static string DefaultPlatformUrl()
    {
        try
        {
            var cfgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
            if (File.Exists(cfgPath))
            {
                var cfg = JsonSerializer.Deserialize<AppConfig>(File.ReadAllText(cfgPath));
                if (cfg != null && !string.IsNullOrWhiteSpace(cfg.ServerIp))
                    return $"http://{cfg.ServerIp}:{cfg.ServerPort}";
            }
        }
        catch { /* 配置缺失时用占位地址 */ }
        return "http://192.168.1.100:5000";
    }

    private void BuildNavigation()
    {
        _navigation.Background = Brushes.White;
        _navigation.Orientation = Orientation.Vertical;

        var items = new[] { ("划课", "hours"), ("设备列表", "devices"), ("任务中心", "tasks"), ("集控平台列表", "platforms") };
        for (int i = 0; i < items.Length; i++)
        {
            var (header, tag) = items[i];
            var idx = i;
            var btn = new Button
            {
                Content = header,
                Style = (Style)Application.Current.FindResource("NavButtonStyle")
            };
            btn.Click += (_, _) =>
            {
                _selectedNavIndex = idx;
                UpdateNavSelection();
                _content.Content = tag switch
                {
                    "hours" => HoursPanel(),
                    "devices" => DevicePanel(),
                    "tasks" => TaskPanel(),
                    _ => PlatformPanel()
                };
            };
            _navigation.Children.Add(btn);
        }

        _selectedNavIndex = 0;
        UpdateNavSelection();
        _content.Content = HoursPanel();
    }

    private void UpdateNavSelection()
    {
        for (int i = 0; i < _navigation.Children.Count; i++)
        {
            if (_navigation.Children[i] is Button btn)
                SetNavButtonState(btn, i == _selectedNavIndex);
        }
    }

    private static void SetNavButtonState(Button btn, bool selected)
    {
        btn.Foreground = selected
            ? new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4))
            : Brushes.Gray;
        btn.FontWeight = selected ? FontWeights.SemiBold : FontWeights.Normal;
        btn.Background = selected
            ? new SolidColorBrush(Color.FromRgb(0xe8, 0xf0, 0xfe))
            : Brushes.Transparent;
    }

    /// <summary>「划课」页：直接嵌入课时划消界面（复用大屏同款 ClassHoursPanelControl），进入控制中心即可直接打卡/划课</summary>
    private UIElement HoursPanel()
    {
        _hoursControl ??= new ClassHoursPanelControl();
        return _hoursControl;
    }

    /// <summary>「设备列表」页：工具栏（刷新/全选/发送到所选）+ 设备表格（勾选列 + 呼叫操作列），即呼叫面板</summary>
    private UIElement DevicePanel()
    {
        var grid = new Grid { Margin = new Thickness(18) };
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition());

        var toolbar = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 10) };
        var refresh = RoundedButton("刷新设备", async () => await RefreshDevices(), primary: false, width: 96);
        var selectAll = RoundedButton("全选", ToggleSelectAll, primary: false, width: 70);
        selectAll.Margin = new Thickness(8, 0, 0, 0);
        var sendToSelected = RoundedButton("发送到所选", SendToSelected, primary: true, width: 110);
        sendToSelected.Margin = new Thickness(8, 0, 0, 0);
        toolbar.Children.Add(refresh);
        toolbar.Children.Add(selectAll);
        toolbar.Children.Add(sendToSelected);
        Grid.SetRow(toolbar, 0);
        grid.Children.Add(toolbar);

        var list = new DataGrid { ItemsSource = _devices, AutoGenerateColumns = false, IsReadOnly = false, BorderThickness = new Thickness(0) };
        list.Columns.Add(new DataGridCheckBoxColumn { Header = "选", Width = 46, Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.IsSelected)) { Mode = System.Windows.Data.BindingMode.TwoWay, UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged } });
        list.Columns.Add(new DataGridTextColumn { Header = "集控平台", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Platform)), Width = 130, IsReadOnly = true });
        list.Columns.Add(new DataGridTextColumn { Header = "设备名称", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Name)), Width = 190, IsReadOnly = true });
        list.Columns.Add(new DataGridTextColumn { Header = "状态", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Status)), Width = 70, IsReadOnly = true });
        list.Columns.Add(new DataGridTextColumn { Header = "版本", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Version)), Width = 90, IsReadOnly = true });
        list.Columns.Add(new DataGridTextColumn { Header = "UUID", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Uuid)), Width = new DataGridLength(1, DataGridLengthUnitType.Star), IsReadOnly = true });
        var opPanel = new FrameworkElementFactory(typeof(StackPanel));
        opPanel.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal);
        var callFactory = new FrameworkElementFactory(typeof(Button));
        callFactory.SetValue(Button.ContentProperty, "呼叫");
        callFactory.SetValue(Button.MarginProperty, new Thickness(2));
        callFactory.SetValue(Button.WidthProperty, 56d);
        callFactory.SetValue(Button.HeightProperty, 24d);
        callFactory.SetValue(Button.CursorProperty, Cursors.Hand);
        callFactory.AddHandler(Button.ClickEvent, new RoutedEventHandler((s, _) =>
        {
            if ((s as Button)?.DataContext is AggregatedDevice dev && dev.Service != null)
                ShowCallDialog(new[] { dev });
        }));
        opPanel.AppendChild(callFactory);
        var renameFactory = new FrameworkElementFactory(typeof(Button));
        renameFactory.SetValue(Button.ContentProperty, "改名");
        renameFactory.SetValue(Button.MarginProperty, new Thickness(2));
        renameFactory.SetValue(Button.WidthProperty, 56d);
        renameFactory.SetValue(Button.HeightProperty, 24d);
        renameFactory.SetValue(Button.CursorProperty, Cursors.Hand);
        renameFactory.AddHandler(Button.ClickEvent, new RoutedEventHandler(async (s, _) =>
        {
            if ((s as Button)?.DataContext is not AggregatedDevice dev || dev.Service == null) return;
            var dlg = new InputDialog("重命名设备", "请输入新的设备名称", dev.Name) { Owner = Window.GetWindow(this) };
            if (dlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(dlg.Value))
            {
                try { await dev.Service.RenameDeviceAsync(dev.Uuid, dlg.Value.Trim()); }
                catch (Exception ex) { MessageBox.Show(ex.Message, "重命名失败", MessageBoxButton.OK, MessageBoxImage.Warning); }
                await RefreshDevices();
            }
        }));
        opPanel.AppendChild(renameFactory);
        list.Columns.Add(new DataGridTemplateColumn { Header = "操作", Width = 124, CellTemplate = new DataTemplate { VisualTree = opPanel } });

        var cell = new Grid();
        cell.Children.Add(new Border
        {
            CornerRadius = new CornerRadius(10),
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe4, 0xe8)),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(1),
            Child = list
        });
        _deviceHint = new TextBlock
        {
            Foreground = Brushes.Gray,
            FontSize = 13,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(24),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Visibility = Visibility.Collapsed,
            IsHitTestVisible = false
        };
        cell.Children.Add(_deviceHint);
        Grid.SetRow(cell, 1);
        grid.Children.Add(cell);

        UpdateDeviceHint();
        _ = RefreshDevices();
        return grid;
    }

    /// <summary>「任务中心」页：显示已连接平台的签到任务（科目 · 班级 · 状态 · 签到进度 · 时间）</summary>
    private UIElement TaskPanel()
    {
        var grid = new Grid { Margin = new Thickness(18) };
        var list = new ListView { ItemsSource = _tasks, BorderThickness = new Thickness(0) };
        var taskText = new FrameworkElementFactory(typeof(TextBlock));
        taskText.SetValue(TextBlock.PaddingProperty, new Thickness(8, 6, 8, 6));
        taskText.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding(nameof(RemoteTask.Summary)));
        list.ItemTemplate = new DataTemplate { VisualTree = taskText };

        var cell = new Grid();
        cell.Children.Add(new Border
        {
            CornerRadius = new CornerRadius(10),
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe4, 0xe8)),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(1),
            Child = list
        });
        _taskHint = new TextBlock
        {
            Foreground = Brushes.Gray,
            FontSize = 13,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(24),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Visibility = Visibility.Collapsed,
            IsHitTestVisible = false
        };
        cell.Children.Add(_taskHint);
        grid.Children.Add(cell);

        UpdateTaskHint();
        _ = RefreshTasks();
        return grid;
    }

    private void UpdateDeviceHint()
    {
        if (_deviceHint == null) return;
        if (_platforms.Count == 0)
        {
            _deviceHint.Text = "尚未连接集控平台：请到左侧「集控平台列表」输入平台地址与账号并连接，连接后此处会显示所有设备，并可勾选后批量发送呼叫。";
            _deviceHint.Visibility = Visibility.Visible;
        }
        else if (_devices.Count == 0)
        {
            _deviceHint.Text = "暂无设备：请确认设备端已开机并在「远程 → 远程服务器设置」中指向同一台服务器。";
            _deviceHint.Visibility = Visibility.Visible;
        }
        else
        {
            _deviceHint.Visibility = Visibility.Collapsed;
        }
    }

    private void UpdateTaskHint()
    {
        if (_taskHint == null) return;
        if (_platforms.Count == 0)
        {
            _taskHint.Text = "尚未连接集控平台：请到左侧「集控平台列表」输入平台地址与账号并连接，连接后此处会显示服务器上的签到任务。";
            _taskHint.Visibility = Visibility.Visible;
        }
        else if (_tasks.Count == 0)
        {
            _taskHint.Text = "暂无签到任务：可在主界面「远程 → 创建签到」发起签到，创建后任务会显示在此处。";
            _taskHint.Visibility = Visibility.Visible;
        }
        else
        {
            _taskHint.Visibility = Visibility.Collapsed;
        }
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
            await ConnectPlatform(name.Text, url.Text, username.Text, password.Password, interactive: true);
        }, primary: true, width: 150);
        connect.Margin = new Thickness(0, 12, 0, 6);
        panel.Children.Add(connect);

        var platformList = new ListBox { ItemsSource = _platforms, DisplayMemberPath = nameof(ConnectedPlatform.Name), Height = 180 };

        var remove = RoundedButton("删除所选平台", () =>
        {
            if (platformList.SelectedItem is not ConnectedPlatform selected) return;
            _platforms.Remove(selected);
            PersistPlatforms();
            _status.Text = $"已移除平台，剩余 {_platforms.Count} 个";
        }, primary: false, width: 130);
        remove.HorizontalAlignment = HorizontalAlignment.Left;
        panel.Children.Add(remove);

        platformList.Margin = new Thickness(0, 12, 0, 0);
        panel.Children.Add(platformList);
        return panel;
    }

    private async Task ConnectPlatform(string name, string url, string username, string password, bool interactive)
    {
        if (_platforms.Any(p => p.Url == url))
        {
            _status.Text = $"平台 {url} 已在列表中";
            if (interactive) MessageBox.Show("该平台地址已在列表中，无需重复连接。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        try
        {
            var service = new RemoteControlService();
            service.SetBaseUrl(url);
            await service.LoginAsync(username, password);
            _platforms.Add(new ConnectedPlatform
            {
                Name = string.IsNullOrWhiteSpace(name) ? url : name,
                Url = url,
                Service = service,
                Username = username,
                ProtectedPassword = Protect(password)
            });
            PersistPlatforms();
            _status.Text = $"已连接 {_platforms.Count} 个集控平台";
            await RefreshAll();
        }
        catch (Exception ex)
        {
            _status.Text = $"连接 {url} 失败：{ex.Message}";
            if (interactive) MessageBox.Show(ex.Message, "连接失败", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    /// <summary>进入控制中心时自动重连上次保存的平台（静默，失败不打扰）</summary>
    private async Task AutoReconnectAsync()
    {
        var saved = PlatformStore.Load();
        foreach (var p in saved)
        {
            await ConnectPlatform(p.Name, p.Url, p.Username, Unprotect(p.ProtectedPassword), interactive: false);
        }
        if (_platforms.Count > 0)
            _status.Text = $"已自动连接 {_platforms.Count} 个集控平台";
    }

    private void PersistPlatforms() => PlatformStore.Save(_platforms);

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
                return (platform, devices, failed: false, error: "");
            }
            catch (Exception ex)
            {
                return (platform, devices: new List<DeviceItem>(), failed: true, error: ex.Message);
            }
        }));
        foreach (var (platform, devices, _, _) in results)
            foreach (var device in devices)
                _devices.Add(new AggregatedDevice
                {
                    Platform = platform.Name,
                    Name = device.Name,
                    Uuid = device.Uuid,
                    Online = device.Online,
                    LastSeen = device.LastSeen ?? "-",
                    Version = device.Version,
                    Service = platform.Service
                });
        _totalDevices.Text = _devices.Count.ToString();
        _onlineDevices.Text = _devices.Count(d => d.Online).ToString();
        var failed = results.Where(r => r.failed).Select(r => r.platform.Name).ToList();
        _status.Text = failed.Count > 0
            ? $"已连接 {_platforms.Count} 个平台，共 {_devices.Count} 台设备（刷新失败：{string.Join("、", failed)}）"
            : $"已连接 {_platforms.Count} 个平台，共 {_devices.Count} 台设备";
        UpdateDeviceHint();
    }

    private async Task RefreshTasks()
    {
        _tasks.Clear();
        // 并行拉取所有平台的任务，单平台失败不影响其他平台
        var snapshot = _platforms.ToList();
        var results = await Task.WhenAll(snapshot.Select(async platform =>
        {
            try { return (tasks: await platform.Service.GetTasksAsync(), failed: false, error: ""); }
            catch (Exception ex) { return (tasks: new List<RemoteTask>(), failed: true, error: ex.Message); }
        }));
        foreach (var (tasks, _, _) in results)
            foreach (var task in tasks) _tasks.Add(task);
        _totalTasks.Text = _tasks.Count.ToString();
        var failed = results.Where(r => r.failed).Select(r => r.error).ToList();
        if (failed.Count > 0) _status.Text = $"任务刷新失败：{string.Join("；", failed)}";
        UpdateTaskHint();
    }

    private void ToggleSelectAll()
    {
        var all = _devices.Count > 0 && _devices.All(d => d.IsSelected);
        foreach (var d in _devices) d.IsSelected = !all;
    }

    private void SendToSelected()
    {
        var targets = _devices.Where(d => d.IsSelected).ToList();
        if (targets.Count == 0)
        {
            MessageBox.Show("请先勾选要接收呼叫的设备", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        ShowCallDialog(targets);
    }

    private static void AddField(Panel panel, string label, Control control) { control.Height = 30; control.Margin = new Thickness(0, 0, 0, 8); panel.Children.Add(new TextBlock { Text = label, FontSize = 12, Foreground = Brushes.Gray }); panel.Children.Add(control); }
    private static TextBlock StatValue() => new() { Text = "0", FontSize = 18, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)) };
    private static UIElement Stat(string label, TextBlock value) { var p = new StackPanel { Margin = new Thickness(20, 0, 30, 0) }; p.Children.Add(new TextBlock { Text = label, FontSize = 11, Foreground = Brushes.Gray }); p.Children.Add(value); return p; }

    // ---- 平台凭据持久化（密码 DPAPI 按当前用户加密） ----

    private static string Protect(string plain) =>
        Convert.ToBase64String(ProtectedData.Protect(Encoding.UTF8.GetBytes(plain), null, DataProtectionScope.CurrentUser));

    private static string Unprotect(string? protectedBase64)
    {
        if (string.IsNullOrEmpty(protectedBase64)) return "";
        try { return Encoding.UTF8.GetString(ProtectedData.Unprotect(Convert.FromBase64String(protectedBase64), null, DataProtectionScope.CurrentUser)); }
        catch { return ""; }
    }

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
        presenter.SetValue(System.Windows.Documents.TextElement.ForegroundProperty, new TemplateBindingExtension(Button.ForegroundProperty));
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

    /// <summary>
    /// 发送呼叫对话框：三种模式
    /// prenotice 待下课时段通知（可设提前分钟数）/ emergency 上课应急通知 / summon 下课传唤（可填学生名单）
    /// 支持单台或多台设备同时接收
    /// </summary>
    private void ShowCallDialog(IReadOnlyList<AggregatedDevice> targets)
    {
        var win = new Window
        {
            Title = $"发送呼叫 - {targets.Count} 台设备",
            Width = 480,
            Height = 560,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = Window.GetWindow(this),
            Background = Brushes.White
        };

        var typeBox = new ComboBox { Width = 260, Height = 30, FontSize = 13 };
        typeBox.Items.Add(new ComboBoxItem { Content = "待下课时段通知（提醒学生即将下课）", Tag = "prenotice" });
        typeBox.Items.Add(new ComboBoxItem { Content = "上课应急通知（立即紧急播报）", Tag = "emergency" });
        typeBox.Items.Add(new ComboBoxItem { Content = "下课传唤（下课后叫学生）", Tag = "summon" });
        typeBox.SelectedIndex = 0;

        var titleBox = new TextBox { Width = 300, Height = 30, FontSize = 13 };
        var messageBox = new TextBox { Width = 420, Height = 90, FontSize = 13, TextWrapping = TextWrapping.Wrap, AcceptsReturn = true, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
        var minutesBox = new TextBox { Width = 70, Height = 30, FontSize = 13, Text = "5" };
        var studentsBox = new TextBox { Width = 420, Height = 70, FontSize = 13, TextWrapping = TextWrapping.Wrap, AcceptsReturn = true, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
        var minutesRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 4, 0, 8) };
        var studentsRow = new StackPanel { Margin = new Thickness(0, 4, 0, 8) };

        minutesRow.Children.Add(new TextBlock { Text = "提前", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 6, 0) });
        minutesRow.Children.Add(minutesBox);
        minutesRow.Children.Add(new TextBlock { Text = " 分钟提醒（0 = 到下课时间提醒）", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0, 0, 0) });
        studentsRow.Children.Add(new TextBlock { Text = "传唤学生名单（每行一个姓名，可留空 = 全体）", FontSize = 12, Margin = new Thickness(0, 0, 0, 4) });
        studentsRow.Children.Add(studentsBox);

        void RefreshExtraFields()
        {
            var type = (typeBox.SelectedItem as ComboBoxItem)?.Tag as string;
            minutesRow.Visibility = type == "prenotice" ? Visibility.Visible : Visibility.Collapsed;
            studentsRow.Visibility = type == "summon" ? Visibility.Visible : Visibility.Collapsed;
        }
        typeBox.SelectionChanged += (_, _) => RefreshExtraFields();
        RefreshExtraFields();

        var form = new StackPanel { Margin = new Thickness(20) };
        var targetNames = string.Join("、", targets.Select(t => t.Name));
        form.Children.Add(new TextBlock { Text = $"向「{targetNames}」发送呼叫", FontSize = 15, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 12), TextWrapping = TextWrapping.Wrap });
        form.Children.Add(new TextBlock { Text = "呼叫类型" });
        form.Children.Add(typeBox);
        form.Children.Add(minutesRow);
        form.Children.Add(new TextBlock { Text = "标题" });
        form.Children.Add(titleBox);
        form.Children.Add(new TextBlock { Text = "内容", Margin = new Thickness(0, 8, 0, 0) });
        form.Children.Add(messageBox);
        form.Children.Add(studentsRow);

        var send = new Button { Content = "发送呼叫", Width = 100, Height = 32, Margin = new Thickness(8) };
        var cancel = new Button { Content = "取消", Width = 80, Height = 32, Margin = new Thickness(8) };
        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 8, 0, 0) };
        btnRow.Children.Add(send);
        btnRow.Children.Add(cancel);
        form.Children.Add(btnRow);

        cancel.Click += (_, _) => win.Close();
        send.Click += async (_, _) =>
        {
            var type = (typeBox.SelectedItem as ComboBoxItem)?.Tag as string ?? "prenotice";
            var title = titleBox.Text.Trim();
            if (string.IsNullOrEmpty(title))
            {
                MessageBox.Show("请输入标题", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            var minutes = int.TryParse(minutesBox.Text, out var m) ? m : 0;
            send.IsEnabled = false;
            try
            {
                var ok = 0;
                var failures = new List<string>();
                foreach (var t in targets)
                {
                    if (t.Service == null)
                    {
                        failures.Add($"{t.Name}：平台未连接");
                        continue;
                    }
                    try
                    {
                        var id = await t.Service.SendCallAsync(t.Uuid, type, title, messageBox.Text.Trim(), minutes, type == "summon" ? studentsBox.Text.Trim() : null);
                        if (id > 0) ok++;
                        else failures.Add($"{t.Name}：发送失败");
                    }
                    catch (Exception ex)
                    {
                        failures.Add($"{t.Name}：{ex.Message}");
                    }
                }
                var msg = $"已向 {ok}/{targets.Count} 台设备发送呼叫";
                if (failures.Count > 0) msg += "\n\n" + string.Join("\n", failures);
                MessageBox.Show(msg, ok == targets.Count ? "发送成功" : "发送完成（部分失败）", MessageBoxButton.OK, ok == targets.Count ? MessageBoxImage.Information : MessageBoxImage.Warning);
                if (ok > 0) win.Close();
            }
            finally
            {
                send.IsEnabled = true;
            }
        };

        win.Content = form;
        win.ShowDialog();
    }
}

/// <summary>持久化的集控平台信息（密码仅存 DPAPI 密文，不落明文）</summary>
public sealed class SavedPlatform
{
    public string Name { get; set; } = "";
    public string Url { get; set; } = "";
    public string Username { get; set; } = "";
    public string ProtectedPassword { get; set; } = "";
}

/// <summary>platforms.json 读写：进入控制中心自动重连已保存的平台</summary>
public static class PlatformStore
{
    private static string FilePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "platforms.json");

    public static List<SavedPlatform> Load()
    {
        try
        {
            if (!File.Exists(FilePath)) return new List<SavedPlatform>();
            return JsonSerializer.Deserialize<List<SavedPlatform>>(File.ReadAllText(FilePath)) ?? new List<SavedPlatform>();
        }
        catch { return new List<SavedPlatform>(); }
    }

    public static void Save(IEnumerable<ConnectedPlatform> platforms)
    {
        try
        {
            var list = platforms.Select(p => new SavedPlatform
            {
                Name = p.Name,
                Url = p.Url,
                Username = p.Username,
                ProtectedPassword = p.ProtectedPassword
            }).ToList();
            File.WriteAllText(FilePath, JsonSerializer.Serialize(list, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { /* 持久化失败不影响本次会话 */ }
    }
}
