using CheckIn.Client.Maui.Pages;

namespace CheckIn.Client.Maui;

public partial class AppShell : Shell
{
    public AppShell()
    {
        InitializeComponent();

        // Register routes for navigation
        Routing.RegisterRoute("TaskDetail", typeof(TaskDetailPage));
        Routing.RegisterRoute("CreateSignIn", typeof(CreateSignInPage));
        Routing.RegisterRoute("SignInResult", typeof(SignInResultPage));
        Routing.RegisterRoute("Settings", typeof(SettingsPage));
    }
}
