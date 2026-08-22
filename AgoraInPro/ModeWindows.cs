using System.Collections.ObjectModel;
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
}

public sealed class AggregatedDevice
{
    public string Platform { get; init; } = "";
    public string Name { get; init; } = "";
    public string Uuid { get; init; } = "";
    public string Status => Online ? "在线" : "离线";
    public bool Online { get; init; }
    public string LastSeen { get; init; } = "-";
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

        var form = new WrapPanel { Margin = new Thickness(16), VerticalAlignment = VerticalAlignment.Center };
        AddField(form, "平台名", _platformName, 120);
        AddField(form, "地址", _url, 230);
        AddField(form, "用户名", _username, 120);
        AddField(form, "密码", _password, 120);
        form.Children.Add(connect);
        form.Children.Add(hours);
        form.Children.Add(refresh);

        var platformList = new ListBox { ItemsSource = _platforms, DisplayMemberPath = "Name", Width = 210, Margin = new Thickness(16, 0, 8, 16) };
        var deviceList = new DataGrid { ItemsSource = _devices, AutoGenerateColumns = false, IsReadOnly = true, Margin = new Thickness(8, 0, 16, 16) };
        deviceList.Columns.Add(new DataGridTextColumn { Header = "集控平台", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Platform)), Width = 150 });
        deviceList.Columns.Add(new DataGridTextColumn { Header = "设备名称", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Name)), Width = 220 });
        deviceList.Columns.Add(new DataGridTextColumn { Header = "状态", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Status)), Width = 90 });
        deviceList.Columns.Add(new DataGridTextColumn { Header = "UUID", Binding = new System.Windows.Data.Binding(nameof(AggregatedDevice.Uuid)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });

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
                    _devices.Add(new AggregatedDevice { Platform = platform.Name, Name = device.Name, Uuid = device.Uuid, Online = device.Online, LastSeen = device.LastSeen ?? "-" });
            }
            catch (Exception ex)
            {
                _status.Text = $"{platform.Name} 刷新失败：{ex.Message}";
            }
        }
        if (_platforms.Count > 0) _status.Text = $"已连接 {_platforms.Count} 个平台，共 {_devices.Count} 台设备";
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
}