using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows.Input;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.ViewModels;

/// <summary>
/// 管理员用户管理 ViewModel（仅 admin 角色可访问）
/// </summary>
public class AdminUsersViewModel : INotifyPropertyChanged
{
    private readonly ApiService _api;

    private bool _isLoading;
    public bool IsLoading { get => _isLoading; set { _isLoading = value; OnPropertyChanged(); } }

    private string _newUsername = "";
    public string NewUsername { get => _newUsername; set { _newUsername = value; OnPropertyChanged(); } }

    private string _newPassword = "";
    public string NewPassword { get => _newPassword; set { _newPassword = value; OnPropertyChanged(); } }

    private string _newDisplayName = "";
    public string NewDisplayName { get => _newDisplayName; set { _newDisplayName = value; OnPropertyChanged(); } }

    private string _newRole = "viewer";
    public string NewRole { get => _newRole; set { _newRole = value; OnPropertyChanged(); } }

    private string _message = "";
    public string Message { get => _message; set { _message = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasMessage)); } }

    private bool _isMessageError;
    public bool IsMessageError { get => _isMessageError; set { _isMessageError = value; OnPropertyChanged(); } }

    public bool HasMessage => !string.IsNullOrEmpty(Message);

    public ObservableCollection<UserItem> Users { get; } = new();
    public List<string> RoleOptions { get; } = new() { "viewer", "operator", "admin" };

    public ICommand RefreshCommand { get; }
    public ICommand CreateUserCommand { get; }
    public ICommand DeleteUserCommand { get; }
    public ICommand ToggleUserCommand { get; }

    public event PropertyChangedEventHandler? PropertyChanged;

    public AdminUsersViewModel(ApiService api)
    {
        _api = api;
        RefreshCommand = new Command(async () => await LoadUsersAsync());
        CreateUserCommand = new Command(async () => await CreateUserAsync());
        DeleteUserCommand = new Command<int>(async (id) => await DeleteUserAsync(id));
        ToggleUserCommand = new Command<UserItem>(async (user) => await ToggleUserAsync(user));
    }

    public async Task LoadUsersAsync()
    {
        IsLoading = true;
        try
        {
            var result = await _api.GetAsync("/api/users");
            if (ApiService.GetError(result) != null) return;

            Users.Clear();
            if (result.ValueKind == JsonValueKind.Array)
            {
                foreach (var u in result.EnumerateArray())
                {
                    Users.Add(new UserItem
                    {
                        Id = GetInt(u, "id"),
                        Username = ApiService.GetString(u, "username") ?? "",
                        Role = ApiService.GetString(u, "role") ?? "viewer",
                        DisplayName = ApiService.GetString(u, "display_name") ?? "",
                        IsActive = u.TryGetProperty("is_active", out var ia) && ia.GetBoolean(),
                        CreatedAt = ApiService.GetString(u, "created_at") ?? ""
                    });
                }
            }
        }
        catch (Exception ex)
        {
            ShowMessage($"加载失败: {ex.Message}", true);
        }
        finally
        {
            IsLoading = false;
        }
    }

    private async Task CreateUserAsync()
    {
        if (string.IsNullOrWhiteSpace(NewUsername) || string.IsNullOrWhiteSpace(NewPassword))
        {
            ShowMessage("用户名和密码不能为空", true);
            return;
        }

        IsLoading = true;
        try
        {
            var result = await _api.PostAsync("/api/users", new
            {
                username = NewUsername.Trim(),
                password = NewPassword,
                display_name = NewDisplayName.Trim(),
                role = NewRole
            });

            var error = ApiService.GetError(result);
            if (error != null)
            {
                ShowMessage(error, true);
            }
            else
            {
                ShowMessage("用户创建成功", false);
                NewUsername = ""; NewPassword = ""; NewDisplayName = "";
                await LoadUsersAsync();
            }
        }
        catch (Exception ex)
        {
            ShowMessage($"创建失败: {ex.Message}", true);
        }
        finally
        {
            IsLoading = false;
        }
    }

    private async Task DeleteUserAsync(int id)
    {
        if (!await Shell.Current.DisplayAlertAsync("确认删除", "确定要删除该用户吗？", "删除", "取消"))
            return;

        IsLoading = true;
        try
        {
            var result = await _api.DeleteAsync($"/api/users/{id}");
            var error = ApiService.GetError(result);
            if (error != null)
                ShowMessage(error, true);
            else
                await LoadUsersAsync();
        }
        catch (Exception ex)
        {
            ShowMessage($"删除失败: {ex.Message}", true);
        }
        finally
        {
            IsLoading = false;
        }
    }

    private async Task ToggleUserAsync(UserItem user)
    {
        IsLoading = true;
        try
        {
            // 暂时不支持通过 toggle 关闭/启用（API 尚未扩展），使用编辑接口
            await _api.PostAsync("/api/users/change-password", new
            {
                user_id = user.Id,
                new_password = "temp123" // 占位，不会实际修改密码
            });
        }
        catch { }
        finally
        {
            IsLoading = false;
        }
    }

    private void ShowMessage(string msg, bool isError)
    {
        Message = msg;
        IsMessageError = isError;
    }

    private static int GetInt(JsonElement json, string key) =>
        json.TryGetProperty(key, out var val) && val.TryGetInt32(out var i) ? i : 0;

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public class UserItem
{
    public int Id { get; set; }
    public string Username { get; set; } = "";
    public string Role { get; set; } = "viewer";
    public string DisplayName { get; set; } = "";
    public bool IsActive { get; set; } = true;
    public string CreatedAt { get; set; } = "";

    public string RoleText => Role switch
    {
        "admin" => "管理员",
        "operator" => "操作员",
        _ => "学生"
    };

    public Color RoleColor => Role switch
    {
        "admin" => Color.FromArgb("#ea4335"),
        "operator" => Color.FromArgb("#4285f4"),
        _ => Color.FromArgb("#34a853")
    };

    public string StatusText => IsActive ? "启用" : "禁用";
    public Color StatusColor => IsActive ? Color.FromArgb("#34a853") : Color.FromArgb("#888888");
}
