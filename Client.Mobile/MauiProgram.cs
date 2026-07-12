using CommunityToolkit.Maui;
using Microsoft.Extensions.Logging;

namespace CheckIn.Client.Mobile;

public static class MauiProgram
{
    public static MauiApp CreateMauiApp()
    {
        var builder = MauiApp.CreateBuilder();
        builder
            .UseMauiApp<App>()
            .UseMauiCommunityToolkit()
            .ConfigureFonts(fonts =>
            {
                fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
                fonts.AddFont("OpenSans-Semibold.ttf", "OpenSansSemibold");
            });

#if DEBUG
        builder.Logging.AddDebug();
#endif

        // Register services
        builder.Services.AddSingleton<CheckIn.Client.Mobile.Services.ServerService>();
        builder.Services.AddSingleton<CheckIn.Client.Mobile.ViewModels.MainViewModel>();

        // Register ViewModels as transient
        builder.Services.AddTransient<CheckIn.Client.Mobile.ViewModels.TaskTabViewModel>();

        // Register pages
        builder.Services.AddTransient<CheckIn.Client.Mobile.Pages.TaskListPage>();
        builder.Services.AddTransient<CheckIn.Client.Mobile.Pages.TaskDetailPage>();
        builder.Services.AddTransient<CheckIn.Client.Mobile.Pages.CreateSignInPage>();
        builder.Services.AddTransient<CheckIn.Client.Mobile.Pages.StudentGridPage>();
        builder.Services.AddTransient<CheckIn.Client.Mobile.Pages.SettingsPage>();

        return builder.Build();
    }
}
