using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows.Input;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.ViewModels;

/// <summary>
/// 二维码签到生成 ViewModel（Admin/Operator）
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
    public bool CanGenerate => !IsLoading && !string.IsNullOrWhiteSpace(Subject) && !string.IsNullOrWhiteSpace(SignPassword);

    private string _resultShortCode = "";
    public string ResultShortCode { get => _resultShortCode; set { _resultShortCode = value; OnPropertyChanged(); } }

    private string _resultSubject = "";
    public string ResultSubject { get => _resultSubject; set { _resultSubject = value; OnPropertyChanged(); } }

    private string _resultClassroom = "";
    public string ResultClassroom { get => _resultClassroom; set { _resultClassroom = value; OnPropertyChanged(); } }

    private int _resultStudentCount;
    public int ResultStudentCount { get => _resultStudentCount; set { _resultStudentCount = value; OnPropertyChanged(); } }

    private string _baseUrl = "";
    public string SignInUrl => string.IsNullOrEmpty(_baseUrl) ? "" : $"{_baseUrl}/s/{ResultShortCode}";

    public ICommand GenerateCommand { get; }
    public ICommand ResetCommand { get; }
    public ICommand CopyUrlCommand { get; }

    public event PropertyChangedEventHandler? PropertyChanged;

    public QRCodeGenerateViewModel(ApiService api)
    {
        _api = api;
        _baseUrl = Preferences.Get("server_url", "http://localhost:5250");
        GenerateCommand = new Command(async () => await GenerateAsync(), () => CanGenerate);
        ResetCommand = new Command(() =>
        {
            Generated = false;
            Subject = ""; Classroom = ""; SignPassword = ""; StudentListText = "";
        });
        CopyUrlCommand = new Command(async () =>
        {
            if (!string.IsNullOrEmpty(SignInUrl))
            {
                await Clipboard.SetTextAsync(SignInUrl);
                await Shell.Current.DisplayAlertAsync("已复制", $"签到链接已复制到剪贴板\n{SignInUrl}", "确定");
            }
        });
    }

    private async Task GenerateAsync()
    {
        if (!CanGenerate) return;

        IsLoading = true;
        try
        {
            // 解析学生名单
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
                students = students
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
