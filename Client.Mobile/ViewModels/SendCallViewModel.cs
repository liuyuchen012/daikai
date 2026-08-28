using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.ViewModels;

/// <summary>
/// 发送呼叫对话框 ViewModel
/// 三种模式：prenotice 待下课时段通知 / emergency 上课应急通知 / summon 下课传唤
/// 支持单台或多台设备同时接收（对应 AgoraInPro ModeWindows.ShowCallDialog）
/// </summary>
public class SendCallViewModel : INotifyPropertyChanged
{
    private readonly ApiService _api;
    private List<CallableDevice> _targets = new();

    public static readonly string[] CallTypeNames =
    {
        "待下课时段通知（提醒学生即将下课）",
        "上课应急通知（立即紧急播报）",
        "下课传唤（下课后叫学生）"
    };
    private static readonly string[] CallTypeValues = { "prenotice", "emergency", "summon" };

    private string _targetNames = "";
    public string TargetNames { get => _targetNames; set { _targetNames = value; OnPropertyChanged(); } }

    private int _targetCount;
    public int TargetCount { get => _targetCount; set { _targetCount = value; OnPropertyChanged(); } }

    private int _selectedTypeIndex;
    public int SelectedTypeIndex
    {
        get => _selectedTypeIndex;
        set
        {
            if (_selectedTypeIndex == value) return;
            _selectedTypeIndex = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsPrenotice));
            OnPropertyChanged(nameof(IsSummon));
        }
    }

    public string CallType => CallTypeValues[Math.Clamp(SelectedTypeIndex, 0, CallTypeValues.Length - 1)];
    public bool IsPrenotice => CallType == "prenotice";
    public bool IsSummon => CallType == "summon";

    private string _title = "";
    public string Title
    {
        get => _title;
        set { _title = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanSend)); }
    }

    private string _message = "";
    public string Message { get => _message; set { _message = value; OnPropertyChanged(); } }

    private string _minutesText = "5";
    public string MinutesText { get => _minutesText; set { _minutesText = value; OnPropertyChanged(); } }

    private string _studentsText = "";
    public string StudentsText { get => _studentsText; set { _studentsText = value; OnPropertyChanged(); } }

    private bool _isSending;
    public bool IsSending { get => _isSending; set { _isSending = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanSend)); } }

    public bool CanSend => !IsSending && !string.IsNullOrWhiteSpace(Title);

    public ICommand SendCommand { get; }
    public ICommand CloseCommand { get; }

    /// <summary>发送完成：参数为成功数、失败列表</summary>
    public event Action<int, List<string>>? SendCompleted;
    public event Action? CloseRequested;

    public SendCallViewModel(ApiService api)
    {
        _api = api;
        SendCommand = new Command(async () =>
        {
            try { await SendAsync(); }
            catch (Exception ex)
            {
                await Shell.Current.DisplayAlertAsync("发送失败", ex.Message, "确定");
            }
        });
        CloseCommand = new Command(() => CloseRequested?.Invoke());
    }

    /// <summary>设置目标设备（由页面在弹窗前调用）</summary>
    public void SetTargets(IReadOnlyList<CallableDevice> targets)
    {
        _targets = targets.ToList();
        TargetCount = _targets.Count;
        TargetNames = string.Join("、", _targets.Select(t => t.Name));
        Title = "";
        Message = "";
        MinutesText = "5";
        StudentsText = "";
        SelectedTypeIndex = 0;
    }

    public async Task SendAsync()
    {
        if (_targets.Count == 0) return;
        IsSending = true;
        try
        {
            var minutes = int.TryParse(MinutesText, out var m) ? m : 0;
            var ok = 0;
            var failures = new List<string>();
            foreach (var t in _targets)
            {
                try
                {
                    var result = await _api.PostAsync("/api/mobile/calls", new
                    {
                        machine_uuid = t.Uuid,
                        type = CallType,
                        title = Title.Trim(),
                        message = Message?.Trim() ?? "",
                        minutes_before = minutes,
                        student_names = CallType == "summon" ? StudentsText?.Trim() : null
                    });
                    var error = ApiService.GetError(result);
                    if (error != null) failures.Add($"{t.Name}：{error}");
                    else ok++;
                }
                catch (Exception ex)
                {
                    failures.Add($"{t.Name}：{ex.Message}");
                }
            }
            SendCompleted?.Invoke(ok, failures);
        }
        finally
        {
            IsSending = false;
        }
    }

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    public event PropertyChangedEventHandler? PropertyChanged;
}
