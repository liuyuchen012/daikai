using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows.Input;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.ViewModels;

/// <summary>
/// 学生扫码签到 ViewModel
/// </summary>
public class StudentScanViewModel : INotifyPropertyChanged
{
    private readonly ApiService _api;

    private bool _isScanning = true;
    public bool IsScanning { get => _isScanning; set { _isScanning = value; OnPropertyChanged(); OnPropertyChanged(nameof(ShowForm)); } }

    private bool _isSubmitting;
    public bool IsSubmitting { get => _isSubmitting; set { _isSubmitting = value; OnPropertyChanged(); } }

    private string _scannedShortCode = "";
    public string ScannedShortCode { get => _scannedShortCode; set { _scannedShortCode = value; OnPropertyChanged(); } }

    private string _studentName = "";
    public string StudentName { get => _studentName; set { _studentName = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanSubmit)); } }

    private string _signPassword = "";
    public string SignPassword { get => _signPassword; set { _signPassword = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanSubmit)); } }

    private bool _checkInSuccess;
    public bool CheckInSuccess { get => _checkInSuccess; set { _checkInSuccess = value; OnPropertyChanged(); } }

    private string _resultMessage = "";
    public string ResultMessage { get => _resultMessage; set { _resultMessage = value; OnPropertyChanged(); } }

    private string _resultDetail = "";
    public string ResultDetail { get => _resultDetail; set { _resultDetail = value; OnPropertyChanged(); } }

    private bool _hasError;
    public bool HasError { get => _hasError; set { _hasError = value; OnPropertyChanged(); } }

    public bool ShowForm => !IsScanning;
    public bool CanSubmit => !IsSubmitting && !string.IsNullOrWhiteSpace(StudentName) &&
                              !string.IsNullOrWhiteSpace(SignPassword);

    public ICommand SubmitCommand { get; }
    public ICommand RescanCommand { get; }

    public event PropertyChangedEventHandler? PropertyChanged;

    public StudentScanViewModel(ApiService api)
    {
        _api = api;
        SubmitCommand = new Command(async () => await SubmitCheckInAsync(), () => CanSubmit);
        RescanCommand = new Command(() =>
        {
            IsScanning = true;
            CheckInSuccess = false;
            HasError = false;
            StudentName = "";
            SignPassword = "";
            ScannedShortCode = "";
        });
    }

    /// <summary>
    /// 处理扫码结果，提取短链码
    /// </summary>
    public void OnBarcodeDetected(string barcodeValue)
    {
        if (!IsScanning) return;

        // 从二维码内容中提取 shortCode
        // 支持的格式: http://xxx/s/abc123, /s/abc123, agorain://checkin/abc123, 或直接的 shortCode
        var shortCode = ExtractShortCode(barcodeValue);
        if (string.IsNullOrEmpty(shortCode)) return;

        ScannedShortCode = shortCode;
        IsScanning = false;
    }

    private async Task SubmitCheckInAsync()
    {
        if (!CanSubmit) return;

        IsSubmitting = true;
        HasError = false;
        try
        {
            var result = await _api.PostAsync("/api/qrcode/checkin", new
            {
                short_code = ScannedShortCode,
                student_name = StudentName.Trim(),
                sign_password = SignPassword.Trim()
            });

            var error = ApiService.GetError(result);
            if (error != null)
            {
                ResultMessage = error;
                HasError = true;
            }
            else
            {
                CheckInSuccess = true;
                ResultMessage = "签到成功！";
                var time = ApiService.GetString(result, "time") ?? "";
                var subject = ApiService.GetString(result, "subject") ?? "";
                var classroom = ApiService.GetString(result, "classroom") ?? "";
                ResultDetail = $"{subject}\n教室: {classroom}\n时间: {time}";
            }
        }
        catch (Exception ex)
        {
            ResultMessage = $"网络错误: {ex.Message}";
            HasError = true;
        }
        finally
        {
            IsSubmitting = false;
        }
    }

    private static string ExtractShortCode(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "";

        // 直接是短链码（6-8位字母数字）
        if (System.Text.RegularExpressions.Regex.IsMatch(value.Trim(), @"^[a-z0-9]{6,8}$"))
            return value.Trim();

        // 从 URL 中提取
        if (Uri.TryCreate(value, UriKind.Absolute, out var uri))
        {
            var segments = uri.AbsolutePath.Trim('/').Split('/');
            if (segments.Length >= 2 && segments[^2] == "s" && !string.IsNullOrEmpty(segments[^1]))
                return segments[^1];
        }

        // 从路径中提取: /s/abc123
        if (value.Contains("/s/"))
        {
            var idx = value.LastIndexOf("/s/");
            var code = value[(idx + 3)..].Split('?', '#', ' ', '\n')[0].Trim();
            if (!string.IsNullOrEmpty(code)) return code;
        }

        return "";
    }

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
