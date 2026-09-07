using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using CallCenter.Models;
using CallCenter.Services;

namespace CallCenter;

/// <summary>呼出端主窗口（Avalonia，跨平台：Windows / macOS / Linux）</summary>
public partial class MainWindow : Window
{
    private readonly AppConfig _config;
    private readonly ApiClient _api = new();
    private readonly ObservableCollection<DeviceInfo> _devices = new();
    private readonly ObservableCollection<CallLog> _logs = new();
    private bool _connected;
    private readonly DispatcherTimer _refreshTimer;

    public MainWindow()
    {
        InitializeComponent();

        _config = AppConfig.Load();
        ServerBox.Text = _config.ServerUrl;
        SenderBox.Text = _config.Sender;

        DeviceList.ItemsSource = _devices;
        LogList.ItemsSource = _logs;

        // 每 10 秒自动刷新设备在线状态
        _refreshTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(10) };
        _refreshTimer.Tick += async (_, _) => await LoadDevicesAsync(true);
        _refreshTimer.Start();

        Opened += async (_, _) =>
        {
            await ConnectAsync();
            await LoadDevicesAsync(true);
        };
    }

    /// <summary>连接服务端（更新地址并探测连通性）</summary>
    private async Task ConnectAsync()
    {
        var url = ServerBox.Text?.Trim() ?? "";
        if (string.IsNullOrEmpty(url))
        {
            SetStatus("请填写服务端地址", false);
            return;
        }
        _api.BaseUrl = url;
        _config.ServerUrl = url;
        _config.Save();

        var ok = await _api.TestConnectionAsync();
        SetStatus(ok ? "已连接" : "连接失败，请检查服务端是否已启动", ok);
        _connected = ok;
    }

    private void SetStatus(string text, bool ok)
    {
        ConnStatusText.Text = text;
        ConnStatusText.Foreground = ok
            ? new SolidColorBrush(Color.FromRgb(0x22, 0xC5, 0x5E))
            : new SolidColorBrush(Color.FromRgb(0xEF, 0x44, 0x44));
    }

    /// <summary>加载设备列表</summary>
    private async Task LoadDevicesAsync(bool quiet)
    {
        if (!_connected) return;
        try
        {
            var list = await _api.GetDevicesAsync();
            _devices.Clear();
            foreach (var d in list) _devices.Add(d);
            DeviceCountText.Text = $"({_devices.Count})";
        }
        catch
        {
            if (!quiet) SetStatus("获取设备列表失败", false);
        }
    }

    private async void ConnectBtn_Click(object? sender, RoutedEventArgs e) => await ConnectAsync();

    private async void RefreshBtn_Click(object? sender, RoutedEventArgs e) => await LoadDevicesAsync(quiet: false);

    private void DeviceList_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        // 选中设备后自动切换到"已选设备"
        if (DeviceList.SelectedItem is DeviceInfo sel)
        {
            TargetOne.IsChecked = true;
            TargetHintText.Text = $"目标：{sel.DisplayName}";
        }
        else
        {
            TargetHintText.Text = "（未选中设备时按全部发送）";
        }
    }

    private async void SendBtn_Click(object? sender, RoutedEventArgs e)
    {
        if (!_connected)
        {
            ShowHint("尚未连接服务端，请先点击「连接」。");
            return;
        }

        var type = TypeUrgent.IsChecked == true ? "urgent" : TypeSpeech.IsChecked == true ? "speech" : "notice";
        var title = TitleBox.Text?.Trim() ?? "";
        var message = MessageText.Text?.Trim() ?? "";
        var senderName = SenderBox.Text?.Trim() ?? "";

        if (title.Length == 0 && message.Length == 0)
        {
            ShowHint("请填写呼叫标题或内容。");
            return;
        }

        // 目标：全部设备 或 当前选中的设备
        DeviceInfo? target = null;
        var targetText = "全部设备";
        if (TargetOne.IsChecked == true && DeviceList.SelectedItem is DeviceInfo sel)
        {
            target = sel;
            targetText = sel.DisplayName;
        }

        SendBtn.IsEnabled = false;
        try
        {
            var ok = await _api.SendCallAsync(target, type, title, message, senderName);
            _logs.Insert(0, new CallLog
            {
                Time = DateTime.Now,
                Type = type,
                Target = targetText,
                Title = title,
                Message = message,
                Sender = senderName,
                Success = ok
            });

            if (ok)
            {
                TitleBox.Text = "";
                MessageText.Text = "";
                SetStatus($"呼叫已发送 → {targetText}", true);
            }
            else
            {
                SetStatus("发送失败，请检查服务端", false);
            }
        }
        catch (Exception ex)
        {
            SetStatus($"发送失败：{ex.Message}", false);
        }
        finally
        {
            SendBtn.IsEnabled = true;
        }
    }

    /// <summary>简易提示对话框（Avalonia 无内置 MessageBox）</summary>
    private void ShowHint(string message)
    {
        var okBtn = new Button { Content = "确定", Width = 90, Height = 32, HorizontalAlignment = HorizontalAlignment.Right };
        var dlg = new Window
        {
            Title = "提示",
            Width = 360,
            Height = 160,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Content = new StackPanel
            {
                Margin = new Thickness(20),
                Spacing = 12,
                Children =
                {
                    new TextBlock { Text = message, TextWrapping = Avalonia.Media.TextWrapping.Wrap, FontSize = 14 },
                    okBtn
                }
            }
        };
        okBtn.Click += (_, _) => dlg.Close();
        dlg.ShowDialog(this);
    }

    protected override void OnClosing(WindowClosingEventArgs e)
    {
        _config.Sender = SenderBox.Text?.Trim() ?? "";
        _config.Save();
        base.OnClosing(e);
    }
}
