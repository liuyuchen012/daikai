using System.Linq;
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
    /// admin/teacher：显示管理 TabBar（教师隐藏用户管理）
    /// student/parent：显示学生 TabBar
    /// </summary>
    public void SetRoleBasedTabs()
    {
        if (_auth.IsAdminOrTeacher)
        {
            AdminTabs.IsVisible = true;
            StudentTabs.IsVisible = false;

            // 只有 admin 才能看到用户管理 tab
            var usersTab = AdminTabs.Items.FirstOrDefault(i => i.Route == "adminusers");
            if (usersTab != null)
            {
                usersTab.IsVisible = _auth.IsAdmin;
            }
        }
        else
        {
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
