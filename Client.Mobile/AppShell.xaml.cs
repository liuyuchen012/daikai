namespace CheckIn.Client.Mobile;

public partial class AppShell : Shell
{
    public AppShell()
    {
        InitializeComponent();

        // Register routes for navigation
        Routing.RegisterRoute("taskdetail", typeof(Pages.TaskDetailPage));
        Routing.RegisterRoute("createsignin", typeof(Pages.CreateSignInPage));
        Routing.RegisterRoute("studentgrid", typeof(Pages.StudentGridPage));

        // Set responsive layout based on device
        UpdateResponsiveLayout();
    }

    public void UpdateResponsiveLayout()
    {
        var isTablet = Helpers.ResponsiveHelper.Instance.IsTablet;

        // Toggle visibility based on device type
        foreach (var item in Items)
        {
            if (item is FlyoutItem fi)
            {
                fi.IsVisible = isTablet;
            }
            if (item is TabBar tb)
            {
                tb.IsVisible = !isTablet;
            }
        }

        Shell.SetFlyoutBehavior(this, isTablet ? FlyoutBehavior.Locked : FlyoutBehavior.Disabled);
    }
}
