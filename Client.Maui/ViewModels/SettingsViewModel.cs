using System.IO;
using System.Text.Json;
using System.Windows.Input;
using CheckIn.Client.Models;

namespace CheckIn.Client.Maui.ViewModels;

public class SettingsViewModel : BaseViewModel
{
    private readonly MainViewModel _mainViewModel;

    private string _serverIp = "";
    public string ServerIp { get => _serverIp; set => SetProperty(ref _serverIp, value); }

    private string _serverPort = "5250";
    public string ServerPort { get => _serverPort; set => SetProperty(ref _serverPort, value); }

    private string _serverPassword = "";
    public string ServerPassword { get => _serverPassword; set => SetProperty(ref _serverPassword, value); }

    private string _newPassword = "";
    public string NewPassword { get => _newPassword; set => SetProperty(ref _newPassword, value); }

    private string _confirmPassword = "";
    public string ConfirmPassword { get => _confirmPassword; set => SetProperty(ref _confirmPassword, value); }

    private bool _onlineMode = true;
    public bool OnlineMode { get => _onlineMode; set => SetProperty(ref _onlineMode, value); }

    private string _statusMessage = "";
    public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

    public ICommand SaveServerSettingsCommand { get; }
    public ICommand SaveAdminSettingsCommand { get; }
    public ICommand GoBackCommand { get; }

    public SettingsViewModel(MainViewModel mainViewModel)
    {
        _mainViewModel = mainViewModel;
        Title = "设置";

        // Load current settings
        ServerIp = mainViewModel.Config.ServerIp;
        ServerPort = mainViewModel.Config.ServerPort.ToString();
        ServerPassword = mainViewModel.Config.ServerPassword;
        OnlineMode = mainViewModel.Config.OnlineMode;

        SaveServerSettingsCommand = new RelayCommand(async _ => await SaveServerSettingsAsync());
        SaveAdminSettingsCommand = new RelayCommand(async _ => await SaveAdminSettingsAsync());
        GoBackCommand = new RelayCommand(async _ => await Shell.Current.GoToAsync(".."));
    }

    private async Task SaveServerSettingsAsync()
    {
        if (!int.TryParse(ServerPort, out int port) || port <= 0 || port > 65535)
        {
            await DialogHelper.AlertAsync("请输入有效的端口号 (1-65535)");
            return;
        }

        _mainViewModel.Config.ServerIp = ServerIp;
        _mainViewModel.Config.ServerPort = port;
        _mainViewModel.Config.ServerPassword = ServerPassword;
        _mainViewModel.Config.OnlineMode = OnlineMode;
        _mainViewModel.SaveGlobalConfig();

        foreach (var tab in _mainViewModel.Tabs)
            tab.UpdateGlobalServerConfig(ServerIp, port, ServerPassword);

        StatusMessage = "服务器设置已保存";
        await DialogHelper.AlertAsync("服务器设置已保存");
    }

    private async Task SaveAdminSettingsAsync()
    {
        if (!string.IsNullOrEmpty(NewPassword))
        {
            if (NewPassword != ConfirmPassword)
            {
                await DialogHelper.AlertAsync("两次输入的密码不一致", "错误");
                return;
            }
            _mainViewModel.Config.AdminPasswordHash = HashPassword(NewPassword);
        }

        _mainViewModel.SaveGlobalConfig();
        StatusMessage = "管理员设置已保存";
        NewPassword = "";
        ConfirmPassword = "";
        await DialogHelper.AlertAsync("管理员设置已保存并生效");
    }

    private static string HashPassword(string pwd)
    {
        using var sha = System.Security.Cryptography.SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(pwd)));
    }
}
