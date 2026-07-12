using System.Collections.ObjectModel;
using System.IO;
using System.Windows.Input;
using CheckIn.Client.Maui.Services;

namespace CheckIn.Client.Maui.ViewModels;

public class CreateSignInViewModel : BaseViewModel
{
    private readonly ServerService _serverService;
    private readonly MainViewModel _mainViewModel;

    private string _signPassword = "";
    public string SignPassword { get => _signPassword; set => SetProperty(ref _signPassword, value); }

    private string _confirmPassword = "";
    public string ConfirmPassword { get => _confirmPassword; set => SetProperty(ref _confirmPassword, value); }

    private string _classroom = "";
    public string Classroom { get => _classroom; set => SetProperty(ref _classroom, value); }

    private string _subject = "数学";
    public string Subject { get => _subject; set => SetProperty(ref _subject, value); }

    private string _studentListStatus = "未导入";
    public string StudentListStatus { get => _studentListStatus; set => SetProperty(ref _studentListStatus, value); }

    private bool _isCreating;
    public bool IsCreating { get => _isCreating; set => SetProperty(ref _isCreating, value); }

    private string _resultMessage = "";
    public string ResultMessage { get => _resultMessage; set => SetProperty(ref _resultMessage, value); }

    private bool _hasResult;
    public bool HasResult { get => _hasResult; set => SetProperty(ref _hasResult, value); }

    private string _shortCode = "";
    public string ShortCode { get => _shortCode; set => SetProperty(ref _shortCode, value); }

    private string _signUrl = "";
    public string SignUrl { get => _signUrl; set => SetProperty(ref _signUrl, value); }

    public ICommand ImportCsvCommand { get; }
    public ICommand CreateCommand { get; }
    public ICommand CopyLinkCommand { get; }
    public ICommand CloseCommand { get; }

    private List<string> _studentList = new();

    public CreateSignInViewModel(ServerService serverService, MainViewModel mainViewModel)
    {
        _serverService = serverService;
        _mainViewModel = mainViewModel;
        Title = "创建签到任务";

        ImportCsvCommand = new RelayCommand(async _ => await ImportCsvAsync());
        CreateCommand = new RelayCommand(async _ => await CreateAsync());
        CopyLinkCommand = new RelayCommand(async _ => await CopyLinkAsync());
        CloseCommand = new RelayCommand(async _ => await Shell.Current.Navigation.PopModalAsync());
    }

    private async Task ImportCsvAsync()
    {
        try
        {
            var result = await FilePicker.Default.PickAsync(new PickOptions
            {
                PickerTitle = "选择CSV学生名单",
                FileTypes = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
                {
                    { DevicePlatform.WinUI, new[] { ".csv" } },
                    { DevicePlatform.macOS, new[] { "public.comma-separated-values-text" } },
                })
            });

            if (result == null) return;

            using var stream = await result.OpenReadAsync();
            using var reader = new StreamReader(stream);
            _studentList = new List<string>();
            bool isFirst = true;
            while (!reader.EndOfStream)
            {
                var line = await reader.ReadLineAsync();
                if (isFirst) { isFirst = false; continue; }
                if (string.IsNullOrEmpty(line)) continue;
                var name = line.Split(',')[0].Trim();
                if (!string.IsNullOrEmpty(name) && name != "姓名")
                    _studentList.Add(name);
            }

            if (_studentList.Count == 0)
            {
                StudentListStatus = "未找到学生数据";
                return;
            }
            StudentListStatus = $"已导入 {_studentList.Count} 名学生";
        }
        catch (Exception ex)
        {
            StudentListStatus = $"导入失败: {ex.Message}";
        }
    }

    private async Task CreateAsync()
    {
        if (string.IsNullOrWhiteSpace(SignPassword))
        {
            await DialogHelper.AlertAsync("请输入签到密码");
            return;
        }
        if (SignPassword != ConfirmPassword)
        {
            await DialogHelper.AlertAsync("两次输入的密码不一致");
            return;
        }
        if (string.IsNullOrWhiteSpace(Classroom))
        {
            await DialogHelper.AlertAsync("请输入教室名称");
            return;
        }
        if (string.IsNullOrWhiteSpace(Subject))
        {
            await DialogHelper.AlertAsync("请输入科目名称");
            return;
        }

        IsCreating = true;
        ResultMessage = "正在创建签到任务，请稍候...";

        try
        {
            _serverService.Initialize(
                _mainViewModel.Config.ServerIp,
                _mainViewModel.Config.ServerPort,
                _mainViewModel.Config.ServerPassword);

            await _serverService.RegisterAsync("AgoraIn签到");

            var result = await _serverService.CreateSignInAsync(SignPassword, Classroom, Subject, _studentList);
            if (result == null)
            {
                await DialogHelper.AlertAsync("创建签到失败，请检查服务器连接");
                return;
            }

            var (shortCode, taskId) = result.Value;
            ShortCode = shortCode;
            SignUrl = $"http://{_mainViewModel.Config.ServerIp}:{_mainViewModel.Config.ServerPort}/s/{shortCode}";
            HasResult = true;
            ResultMessage = "创建成功";

            // Create local task tab
            _mainViewModel.AddTab($"{Classroom} {Subject} 签到", Subject, isSignIn: true, signInTaskId: taskId);
        }
        catch (Exception ex)
        {
            ResultMessage = $"创建失败: {ex.Message}";
            await DialogHelper.AlertAsync($"创建签到失败：{ex.Message}");
        }
        finally
        {
            IsCreating = false;
        }
    }

    private async Task CopyLinkAsync()
    {
        if (string.IsNullOrEmpty(SignUrl)) return;
        await Clipboard.Default.SetTextAsync(SignUrl);
        await DialogHelper.AlertAsync("签到链接已复制到剪贴板");
    }
}
