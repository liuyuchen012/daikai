using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows.Input;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.ViewModels;

/// <summary>
/// 二维码签到生成 ViewModel（Admin/Teacher）
/// </summary>
public class QRCodeGenerateViewModel : INotifyPropertyChanged
{
    private readonly ApiService _api;

    private string _subject = "";
    public string Subject { get => _subject; set { _subject = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanGenerate)); } }

    private string _classroom = "";
    public string Classroom { get => _classroom; set { _classroom = value; OnPropertyChanged(); } }

    private string _signPassword = "";
    public string SignPassword { get => _signPassword; set { _signPassword = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanGenerate)); } }

    private string _studentListText = "";
    public string StudentListText
    {
        get => _studentListText;
        set { _studentListText = value; OnPropertyChanged(); }
    }

    private bool _isLoading;
    public bool IsLoading { get => _isLoading; set { _isLoading = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanGenerate)); } }

    private bool _generated;
    public bool Generated { get => _generated; set { _generated = value; OnPropertyChanged(); OnPropertyChanged(nameof(ShowForm)); } }

    public bool ShowForm => !Generated;
    public bool CanGenerate => !IsLoading && !string.IsNullOrWhiteSpace(Subject)
        && !string.IsNullOrWhiteSpace(SignPassword) && !string.IsNullOrWhiteSpace(SelectedDeviceUuid);

    // ---- 设备选择 ----
    private string _selectedDeviceUuid = "";
    public string SelectedDeviceUuid
    {
        get => _selectedDeviceUuid;
        set { _selectedDeviceUuid = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanGenerate)); }
    }

    private string _selectedDeviceName = "";
    public string SelectedDeviceName
    {
        get => _selectedDeviceName;
        set
        {
            _selectedDeviceName = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(SelectedDeviceHint));
            OnPropertyChanged(nameof(IsSelectedDeviceHintVisible));
        }
    }

    public string SelectedDeviceHint => string.IsNullOrEmpty(SelectedDeviceName) ? "" : $"已选设备: {SelectedDeviceName}";
    public bool IsSelectedDeviceHintVisible => !string.IsNullOrEmpty(SelectedDeviceName);

    public ObservableCollection<DevicePickerItem> DeviceList { get; } = [];

    private string _resultShortCode = "";
    public string ResultShortCode { get => _resultShortCode; set { _resultShortCode = value; OnPropertyChanged(); } }

    private string _resultSubject = "";
    public string ResultSubject { get => _resultSubject; set { _resultSubject = value; OnPropertyChanged(); } }

    private string _resultClassroom = "";
    public string ResultClassroom { get => _resultClassroom; set { _resultClassroom = value; OnPropertyChanged(); } }

    private int _resultStudentCount;
    public int ResultStudentCount { get => _resultStudentCount; set { _resultStudentCount = value; OnPropertyChanged(); } }

    private string _resultDeviceName = "";
    public string ResultDeviceName { get => _resultDeviceName; set { _resultDeviceName = value; OnPropertyChanged(); } }

    private string _baseUrl = "";
    public string SignInUrl => string.IsNullOrEmpty(_baseUrl) ? "" : $"{_baseUrl}/s/{ResultShortCode}";

    public ICommand GenerateCommand { get; }
    public ICommand ResetCommand { get; }
    public ICommand CopyUrlCommand { get; }
    public ICommand DeviceSelectedCommand { get; }

    public event PropertyChangedEventHandler? PropertyChanged;

    public QRCodeGenerateViewModel(ApiService api)
    {
        _api = api;
        _baseUrl = Preferences.Get("server_url", "http://localhost:5250");
        GenerateCommand = new Command(async () =>
        {
            try { await GenerateAsync(); }
            catch { /* 防止 async void 异常崩溃 */ }
        }, () => CanGenerate);
        ResetCommand = new Command(() =>
        {
            Generated = false;
            Subject = ""; Classroom = ""; SignPassword = ""; StudentListText = "";
            SelectedDeviceUuid = ""; SelectedDeviceName = "";
        });
        CopyUrlCommand = new Command(async () =>
        {
            try
            {
                if (!string.IsNullOrEmpty(SignInUrl))
                {
                    await Clipboard.SetTextAsync(SignInUrl);
                    await Shell.Current.DisplayAlertAsync("已复制", $"签到链接已复制到剪贴板\n{SignInUrl}", "确定");
                }
            }
            catch { /* 防止 async void 异常崩溃 */ }
        });
        DeviceSelectedCommand = new Command<DevicePickerItem>(item =>
        {
            if (item != null)
            {
                SelectedDeviceUuid = item.Uuid;
                SelectedDeviceName = item.DisplayText;
            }
        });
    }

    /// <summary>
    /// 加载可用设备列表
    /// </summary>
    public async Task LoadDevicesAsync()
    {
        try
        {
            var result = await _api.GetAsync("/api/mobile/devices");
            if (ApiService.GetError(result) != null) return;

            DeviceList.Clear();
            if (result.TryGetProperty("devices", out var devices) && devices.ValueKind == JsonValueKind.Array)
            {
                foreach (var d in devices.EnumerateArray())
                {
                    var uuid = ApiService.GetString(d, "uuid") ?? "";
                    var name = ApiService.GetString(d, "name") ?? "未知";
                    var online = d.TryGetProperty("online", out var on) && on.GetBoolean();
                    DeviceList.Add(new DevicePickerItem
                    {
                        Uuid = uuid,
                        Name = name,
                        DisplayText = $"{name}{(online ? " [在线]" : " [离线]")}"
                    });

                    // 如果还没选择设备，自动选第一个在线的
                    if (string.IsNullOrEmpty(SelectedDeviceUuid) && online)
                    {
                        SelectedDeviceUuid = uuid;
                        SelectedDeviceName = name;
                    }
                }
                // 没在线的就选第一个
                if (string.IsNullOrEmpty(SelectedDeviceUuid) && DeviceList.Count > 0)
                {
                    SelectedDeviceUuid = DeviceList[0].Uuid;
                    SelectedDeviceName = DeviceList[0].Name;
                }
            }
        }
        catch { }
    }

    private async Task GenerateAsync()
    {
        if (!CanGenerate) return;

        IsLoading = true;
        try
        {
            var students = StudentListText
                .Split('\n', StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();

            var result = await _api.PostAsync("/api/qrcode/generate", new
            {
                subject = Subject.Trim(),
                classroom = Classroom.Trim(),
                sign_password = SignPassword.Trim(),
                students = students,
                machine_uuid = SelectedDeviceUuid
            });

            var error = ApiService.GetError(result);
            if (error != null)
            {
                await Shell.Current.DisplayAlertAsync("创建失败", error, "确定");
                return;
            }

            ResultShortCode = ApiService.GetString(result, "short_code") ?? "";
            ResultSubject = ApiService.GetString(result, "subject") ?? "";
            ResultClassroom = ApiService.GetString(result, "classroom") ?? "";
            ResultDeviceName = ApiService.GetString(result, "machine_name") ?? SelectedDeviceName;
            if (result.TryGetProperty("student_count", out var sc) && sc.TryGetInt32(out var c))
                ResultStudentCount = c;
            else
                ResultStudentCount = students.Count;

            _baseUrl = Preferences.Get("server_url", "http://localhost:5250");
            OnPropertyChanged(nameof(SignInUrl));

            Generated = true;
        }
        catch (Exception ex)
        {
            await Shell.Current.DisplayAlertAsync("创建失败", $"网络错误: {ex.Message}", "确定");
        }
        finally
        {
            IsLoading = false;
        }
    }

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public class DevicePickerItem
{
    public string Uuid { get; set; } = "";
    public string Name { get; set; } = "";
    public string DisplayText { get; set; } = "";
}
