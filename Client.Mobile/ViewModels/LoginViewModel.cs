using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.ViewModels;

/// <summary>
/// 登录页面 ViewModel
/// </summary>
public class LoginViewModel : INotifyPropertyChanged
{
    private readonly AuthService _auth;
    private readonly ApiService _api;

    private string _serverUrl = "http://";
    public string ServerUrl
    {
        get => _serverUrl;
        set { _serverUrl = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanLogin)); }
    }

    private string _username = "";
    public string Username
    {
        get => _username;
        set { _username = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanLogin)); }
    }

    private string _password = "";
    public string Password
    {
        get => _password;
        set { _password = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanLogin)); }
    }

    private string _errorMessage = "";
    public string ErrorMessage
    {
        get => _errorMessage;
        set { _errorMessage = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasError)); }
    }

    private bool _isLoading;
    public bool IsLoading
    {
        get => _isLoading;
        set { _isLoading = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanLogin)); }
    }

    public bool HasError => !string.IsNullOrEmpty(ErrorMessage);
    public bool CanLogin => !IsLoading && !string.IsNullOrWhiteSpace(ServerUrl) &&
                             !string.IsNullOrWhiteSpace(Username) &&
                             !string.IsNullOrWhiteSpace(Password);

    public ICommand LoginCommand { get; }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event Action? LoginSucceeded;

    public LoginViewModel(AuthService auth, ApiService api)
    {
        _auth = auth;
        _api = api;

        // 加载上次保存的服务器地址
        var savedUrl = Preferences.Get("server_url", "http://");
        _serverUrl = savedUrl;
        OnPropertyChanged(nameof(ServerUrl));

        LoginCommand = new Command(async () => await LoginAsync(), () => CanLogin);
    }

    private async Task LoginAsync()
    {
        if (!CanLogin) return;

        IsLoading = true;
        ErrorMessage = "";

        try
        {
            var url = ServerUrl.TrimEnd('/');
            var (success, error) = await _auth.LoginAsync(url, Username, Password);
            if (success)
            {
                // 保存服务器地址
                Preferences.Set("server_url", url);
                LoginSucceeded?.Invoke();
            }
            else
            {
                ErrorMessage = error ?? "登录失败";
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = $"连接失败: {ex.Message}";
        }
        finally
        {
            IsLoading = false;
        }
    }

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
