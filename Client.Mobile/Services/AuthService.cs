using System.Text.Json;

namespace CheckIn.Client.Mobile.Services;

/// <summary>
/// 认证服务 - 管理用户登录状态、令牌持久化
/// </summary>
public class AuthService
{
    private readonly ApiService _api;
    private const string TokenKey = "auth_token";
    private const string UserKey = "auth_user";

    public string? CurrentToken => _api.Token;
    public string? CurrentUsername { get; private set; }
    public string? CurrentRole { get; private set; }
    public string? CurrentDisplayName { get; private set; }
    public bool IsLoggedIn => !string.IsNullOrEmpty(CurrentToken);

    /// <summary>标准化角色名（兼容旧角色）</summary>
    private string NormalizedRole => CurrentRole switch
    {
        "operator" => "teacher",
        "viewer" => "student",
        _ => CurrentRole ?? ""
    };

    public bool IsAdmin => NormalizedRole == "admin";
    public bool IsTeacher => NormalizedRole == "teacher";
    public bool IsStudent => NormalizedRole == "student";
    public bool IsParent => NormalizedRole == "parent";
    /// <summary>管理员或教师（有管理权限）</summary>
    public bool IsAdminOrTeacher => NormalizedRole == "admin" || NormalizedRole == "teacher";

    public AuthService(ApiService api)
    {
        _api = api;
    }

    /// <summary>
    /// 从本地存储加载保存的登录状态
    /// </summary>
    public void LoadSavedState()
    {
        var token = Preferences.Get(TokenKey, null);
        var userJson = Preferences.Get(UserKey, null);
        if (!string.IsNullOrEmpty(token))
        {
            _api.SetToken(token);
            if (!string.IsNullOrEmpty(userJson))
            {
                try
                {
                    var user = JsonDocument.Parse(userJson).RootElement;
                    CurrentUsername = user.TryGetProperty("username", out var un) ? un.GetString() : null;
                    CurrentRole = user.TryGetProperty("role", out var r) ? r.GetString() : null;
                    CurrentDisplayName = user.TryGetProperty("display_name", out var dn) ? dn.GetString() : null;
                }
                catch { }
            }
        }
    }

    /// <summary>
    /// 登录：发送用户名密码到服务器，获取 Token
    /// </summary>
    public async Task<(bool Success, string? Error)> LoginAsync(string baseUrl, string username, string password)
    {
        _api.BaseUrl = baseUrl;
        var result = await _api.PostAsync("/api/auth/login", new
        {
            username,
            password
        });

        var error = ApiService.GetError(result);
        if (error != null)
            return (false, error);

        var token = ApiService.GetString(result, "token");
        if (string.IsNullOrEmpty(token))
            return (false, "服务器未返回令牌");

        _api.SetToken(token);
        CurrentUsername = username;

        // 解析用户信息
        if (result.TryGetProperty("user", out var userNode))
        {
            CurrentRole = ApiService.GetString(userNode, "role");
            CurrentDisplayName = ApiService.GetString(userNode, "display_name");
        }

        // 保存到本地
        Preferences.Set(TokenKey, token);
        var userInfo = JsonSerializer.Serialize(new
        {
            username = CurrentUsername,
            role = CurrentRole,
            display_name = CurrentDisplayName
        });
        Preferences.Set(UserKey, userInfo);

        return (true, null);
    }

    /// <summary>
    /// 验证当前 Token 是否仍然有效
    /// </summary>
    public async Task<bool> VerifyTokenAsync()
    {
        if (!IsLoggedIn) return false;
        var result = await _api.PostAsync("/api/auth/verify");
        if (result.TryGetProperty("valid", out var valid))
            return valid.GetBoolean();
        return false;
    }

    /// <summary>
    /// 登出：清除 Token 和本地存储
    /// </summary>
    public void Logout()
    {
        _api.SetToken(null);
        CurrentUsername = null;
        CurrentRole = null;
        CurrentDisplayName = null;
        Preferences.Remove(TokenKey);
        Preferences.Remove(UserKey);
    }
}
