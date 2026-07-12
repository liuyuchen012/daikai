using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile;

public partial class App : Application
{
    private readonly AuthService _auth;

    public App(AuthService auth)
    {
        InitializeComponent();
        _auth = auth;
    }

    protected override Window CreateWindow(IActivationState? activationState)
    {
        var shell = new AppShell(_auth);

        shell.Navigated += async (s, e) =>
        {
            // 导航到 main 路由时，根据角色切换 UI
            if (e.Current?.Location?.OriginalString?.Contains("main") == true)
            {
                shell.SetRoleBasedTabs();
            }
        };

        return new Window(shell);
    }
}
