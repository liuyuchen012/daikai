using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using CheckIn.Client.Models;
using CheckIn.Client.Services;

namespace CheckIn.Client;

public sealed class ConnectedPlatform
{
    public string Name { get; init; } = "";
    public string Url { get; init; } = "";
    public RemoteControlService Service { get; init; } = null!;
    /// <summary>登录用户名（用于持久化后自动重连）</summary>
    public string Username { get; init; } = "";
    /// <summary>DPAPI 加密后的登录密码（Base64，不落明文）</summary>
    public string ProtectedPassword { get; init; } = "";
}

public sealed class AggregatedDevice : INotifyPropertyChanged
{
    public string Platform { get; init; } = "";
    public string Name { get; init; } = "";
    public string Uuid { get; init; } = "";
    public string Status => Online ? "在线" : "离线";
    public bool Online { get; init; }
    public string LastSeen { get; init; } = "-";
    /// <summary>客户端版本号（客户端上报，如 v3.2.4）</summary>
    public string Version { get; init; } = "";
    public RemoteControlService? Service { get; init; }

    private bool _isSelected;
    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected == value) return;
            _isSelected = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}

public class ControlModeWindow : Window
{
    private readonly ObservableCollection<ConnectedPlatform> _platforms = new();
    private readonly ObservableCollection<AggregatedDevice> _devices = new();
    private readonly TextBox _platformName = new() { Text = "集控平台" };
    private readonly TextBox _url = new() { Text = "http://192.168.1.100:5000" };
    private readonly TextBox _username = new();
    private readonly PasswordBox _password = new();
    private readonly TextBlock _status = new() { Foreground = Brushes.Gray };

    public ControlModeWindow()
    {
        Title = "AgoraIn · 控制模式";
        Width = 1100;
        Height = 700;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        WindowStyle = WindowStyle.None;
        AllowsTransparency = true;
        Background = Brushes.Transparent;

        var connect = new Button { Content = "连接平台", Width = 100, Height = 30, Margin = new Thickness(8, 0, 0, 0) };
        connect.Click += async (_, _) => await ConnectPlatform();
        var hours = new Button { Content = "课时划消", Width = 100, Height = 30, Margin = new Thickness(8, 0, 0, 0) };
        hours.Click += (_, _) => new ClassHoursWindow { Owner = this }.Show();
        var refresh = new Button { Content = "刷新设备", Width = 100, Height = 30, Margin = new Thickness(8, 0, 0, 0) };
        refresh.Click += async (_, _) => await RefreshDevices();
        var selectAll = new Button { Content = "全选", Width = 70, Height = 30, Margin = new Thickness(8, 0, 0, 0) };
        selectAll.Click += (_, _) => ToggleSelectAll();
        var batchCall = new Button { Content = "发送到所选", Width = 110, Height = 30, Margin = new Thickness(8, 0, 0, 0) };
        batchCall.Click += (_, _) => SendToSelected();

        var form = new WrapPanel { Margin = new Thickness(16), VerticalAlignment = VerticalAlignment.Center };
        AddField(form, "平台名", _platformName, 120);
        AddField(form, "地址", _url, 230);
        AddField(form, "用户名", _username, 120);
        AddField(form, "密码", _password, 120);
        form.Children.Add(connect);
        form.Children.Add(hours);
        form.Children.Add(refresh);
        form.Children.Add(selectAll);
        form.Children.Add(batchCall);

        var platformList = new ListBox { ItemsSource = _platforms, DisplayMemberPath = "Name", Width = 210, Margin = new Thickness(16, 0, 8, 16) };
        var deviceList = new DataGrid { ItemsSource = _devices, AutoGenerateColumns = false, IsReadOnly = true, Margin = new Thickness(8, 0, 16, 16) };
        deviceList.Columns.Add(new DataGridCheckBoxColumn { Header = "选", Width = 46, Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.IsSelected)) { Mode = System.Windows.Data.BindingMode.TwoWay, UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged } });
        deviceList.Columns.Add(new DataGridTextColumn { Header = "集控平台", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Platform)), Width = 130 });
        deviceList.Columns.Add(new DataGridTextColumn { Header = "设备名称", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Name)), Width = 180 });
        deviceList.Columns.Add(new DataGridTextColumn { Header = "状态", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Status)), Width = 70 });
        deviceList.Columns.Add(new DataGridTextColumn { Header = "UUID", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Uuid)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        var callFactory = new FrameworkElementFactory(typeof(Button));
        callFactory.SetValue(Button.ContentProperty, "呼叫");
        callFactory.SetValue(Button.MarginProperty, new Thickness(2));
        callFactory.SetValue(Button.WidthProperty, 64d);
        callFactory.SetValue(Button.HeightProperty, 24d);
        callFactory.AddHandler(Button.ClickEvent, new RoutedEventHandler((s, _) =>
        {
            if ((s as Button)?.DataContext is AggregatedDevice dev && dev.Service != null)
                ShowCallDialog(new[] { dev });
        }));
        deviceList.Columns.Add(new DataGridTemplateColumn { Header = "操作", Width = 80, CellTemplate = new DataTemplate { VisualTree = callFactory } });

        var body = new Grid();
        body.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(230) });
        body.ColumnDefinitions.Add(new ColumnDefinition());
        Grid.SetColumn(platformList, 0);
        Grid.SetColumn(deviceList, 1);
        body.Children.Add(platformList);
        body.Children.Add(deviceList);

        var titleBar = new Grid { Height = 44, Background = new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)) };
        titleBar.MouseLeftButtonDown += (_, e) => { if (e.ChangedButton == MouseButton.Left) DragMove(); };
        titleBar.ColumnDefinitions.Add(new ColumnDefinition());
        titleBar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var sceneTabs = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(12, 0, 0, 0) };
        sceneTabs.Children.Add(SceneButton("大屏", () => ReturnToDisplay()));
        sceneTabs.Children.Add(SceneButton("控制", () => { }));
        titleBar.Children.Add(sceneTabs);
        var close = new Button { Content = "✕", Width = 46, Height = 30, Foreground = Brushes.White, Background = Brushes.Transparent, BorderThickness = new Thickness(0), Cursor = Cursors.Hand };
        close.Click += (_, _) => Close();
        Grid.SetColumn(close, 1);
        titleBar.Children.Add(close);

        var root = new DockPanel();
        DockPanel.SetDock(titleBar, Dock.Top);
        root.Children.Add(titleBar);
        DockPanel.SetDock(form, Dock.Top);
        DockPanel.SetDock(_status, Dock.Bottom);
        root.Children.Add(form);
        root.Children.Add(_status);
        root.Children.Add(body);
        Content = root;
    }

    private void ReturnToDisplay()
    {
        if (Owner is MainWindow displayWindow)
        {
            Close();
            displayWindow.Show();
            displayWindow.Activate();
        }
    }

    private static Button SceneButton(string text, Action action)
    {
        var button = new Button { Content = text, Width = 58, Height = 30, Margin = new Thickness(2), Foreground = Brushes.White, Background = Brushes.Transparent, BorderThickness = new Thickness(0), Cursor = Cursors.Hand };
        button.Click += (_, _) => action();
        return button;
    }

    private async Task ConnectPlatform()
    {
        try
        {
            var service = new RemoteControlService();
            service.SetBaseUrl(_url.Text.Trim());
            await service.LoginAsync(_username.Text.Trim(), _password.Password);
            var platform = new ConnectedPlatform { Name = string.IsNullOrWhiteSpace(_platformName.Text) ? _url.Text.Trim() : _platformName.Text.Trim(), Url = _url.Text.Trim(), Service = service };
            _platforms.Add(platform);
            _status.Text = $"已连接 {_platforms.Count} 个平台";
            await RefreshDevices();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "连接失败", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async Task RefreshDevices()
    {
        _devices.Clear();
        foreach (var platform in _platforms)
        {
            try
            {
                foreach (var device in await platform.Service.GetDevicesAsync())
                    _devices.Add(new AggregatedDevice { Platform = platform.Name, Name = device.Name, Uuid = device.Uuid, Online = device.Online, LastSeen = device.LastSeen ?? "-", Service = platform.Service });
            }
            catch (Exception ex)
            {
                _status.Text = $"{platform.Name} 刷新失败：{ex.Message}";
            }
        }
        if (_platforms.Count > 0) _status.Text = $"已连接 {_platforms.Count} 个平台，共 {_devices.Count} 台设备";
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
            MessageBox.Show("请先勾选要接收信息的设备", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        ShowCallDialog(targets);
    }

    private static void AddField(Panel panel, string label, Control control, double width)
    {
        var group = new StackPanel { Margin = new Thickness(0, 0, 6, 0) };
        group.Children.Add(new TextBlock { Text = label, FontSize = 11, Foreground = Brushes.Gray });
        control.Width = width;
        control.Height = 30;
        group.Children.Add(control);
        panel.Children.Add(group);
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
            Owner = this,
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