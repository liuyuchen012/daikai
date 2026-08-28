using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows.Input;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.ViewModels;

/// <summary>
/// 控制模式 ViewModel：设备列表（支持多选）+ 批量发送呼叫
/// 对应 AgoraInPro 的 ControlModeWindow
/// </summary>
public class ControlModeViewModel : INotifyPropertyChanged
{
    private readonly ApiService _api;

    private bool _isLoading;
    public bool IsLoading { get => _isLoading; set { _isLoading = value; OnPropertyChanged(); } }

    private string _statusText = "";
    public string StatusText { get => _statusText; set { _statusText = value; OnPropertyChanged(); } }

    public ObservableCollection<CallableDevice> Devices { get; } = new();

    public ICommand RefreshCommand { get; }
    public ICommand ToggleSelectAllCommand { get; }
    public ICommand SendToSelectedCommand { get; }

    /// <summary>请求打开呼叫对话框（携带已选设备）</summary>
    public event Action<List<CallableDevice>>? SendCallRequested;

    public ControlModeViewModel(ApiService api)
    {
        _api = api;

        RefreshCommand = new Command(async () =>
        {
            try { await LoadDevicesAsync(); }
            catch { /* 防止 async void 异常崩溃 */ }
        });
        ToggleSelectAllCommand = new Command(() =>
        {
            var all = Devices.Count > 0 && Devices.All(d => d.IsSelected);
            foreach (var d in Devices) d.IsSelected = !all;
        });
        SendToSelectedCommand = new Command(() =>
        {
            var targets = Devices.Where(d => d.IsSelected).ToList();
            if (targets.Count == 0)
            {
                _ = Shell.Current.DisplayAlertAsync("提示", "请先勾选要接收信息的设备", "确定");
                return;
            }
            SendCallRequested?.Invoke(targets);
        });
    }

    public async Task LoadDevicesAsync()
    {
        IsLoading = true;
        try
        {
            var result = await _api.GetAsync("/api/mobile/devices");
            if (ApiService.GetError(result) != null)
            {
                StatusText = "获取设备列表失败";
                return;
            }

            Devices.Clear();
            if (result.TryGetProperty("devices", out var devices) && devices.ValueKind == JsonValueKind.Array)
            {
                foreach (var d in devices.EnumerateArray())
                {
                    var dev = new CallableDevice
                    {
                        Name = ApiService.GetString(d, "name") ?? "未知设备",
                        Uuid = ApiService.GetString(d, "uuid") ?? "",
                        Online = d.TryGetProperty("online", out var on) && on.GetBoolean()
                    };
                    dev.PropertyChanged += (_, e) =>
                    {
                        if (e.PropertyName == nameof(CallableDevice.IsSelected)) UpdateStatus();
                    };
                    Devices.Add(dev);
                }
            }
            UpdateStatus();
        }
        catch (Exception ex)
        {
            StatusText = $"加载失败：{ex.Message}";
            System.Diagnostics.Debug.WriteLine($"ControlMode load error: {ex.Message}");
        }
        finally
        {
            IsLoading = false;
        }
    }

    private void UpdateStatus()
    {
        var selected = Devices.Count(d => d.IsSelected);
        StatusText = selected > 0
            ? $"共 {Devices.Count} 台设备，已选 {selected} 台"
            : $"共 {Devices.Count} 台设备";
    }

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    public event PropertyChangedEventHandler? PropertyChanged;
}

/// <summary>
/// 可呼叫设备（含勾选状态）
/// </summary>
public class CallableDevice : INotifyPropertyChanged
{
    public string Name { get; set; } = "";
    public string Uuid { get; set; } = "";
    public bool Online { get; set; }
    public string StatusText => Online ? "在线" : "离线";
    public Color StatusColor => Online ? Color.FromArgb("#34a853") : Color.FromArgb("#888888");

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
