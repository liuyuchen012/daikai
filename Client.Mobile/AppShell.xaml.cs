using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile;

public partial class AppShell : Shell
{
    private readonly AuthService _auth;

    public AppShell(AuthService auth)
    {
        InitializeComponent();
        _auth = auth;

        // 注册子页面路由
        Routing.RegisterRoute("attendance", typeof(Pages.AttendanceDetailPage));
    }

    /// <summary>
    /// 根据用户角色切换显示的 TabBar
    /// </summary>
    public void SetRoleBasedTabs()
    {
        if (_auth.IsAdmin)
        {
            // 管理员：显示管理 TabBar
            AdminTabs.IsVisible = true;
            StudentTabs.IsVisible = false;
        }
        else
        {
            // 学生：显示学生 TabBar
            AdminTabs.IsVisible = false;
            StudentTabs.IsVisible = true;
        }
    }

    /// <summary>
    /// 登出：返回登录页
    /// </summary>
    public async Task LogoutAsync()
    {
        _auth.Logout();
        await GoToAsync("//login");
    }
}
