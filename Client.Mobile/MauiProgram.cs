using CommunityToolkit.Maui;
using Microsoft.Extensions.Logging;
using ZXing.Net.Maui.Controls;

namespace CheckIn.Client.Mobile;

public static class MauiProgram
{
    public static MauiApp CreateMauiApp()
    {
        var builder = MauiApp.CreateBuilder();
        builder
            .UseMauiApp<App>()
            .UseMauiCommunityToolkit()
            .UseBarcodeReader()
            .ConfigureFonts(fonts =>
            {
                fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
                fonts.AddFont("OpenSans-Semibold.ttf", "OpenSansSemibold");
            });

#if DEBUG
        builder.Logging.AddDebug();
#endif

        // ===== 核心服务（Singleton） =====
        builder.Services.AddSingleton<Services.ApiService>();
        builder.Services.AddSingleton<Services.AuthService>();

        // ===== 保留旧的兼容服务 =====
        builder.Services.AddSingleton<Services.ServerService>();
        builder.Services.AddSingleton<ViewModels.MainViewModel>();

        // ===== ViewModels（Transient） =====
        builder.Services.AddTransient<ViewModels.LoginViewModel>();
        builder.Services.AddTransient<ViewModels.AdminDashboardViewModel>();
        builder.Services.AddTransient<ViewModels.AdminTasksViewModel>();
        builder.Services.AddTransient<ViewModels.AdminUsersViewModel>();
        builder.Services.AddTransient<ViewModels.QRCodeGenerateViewModel>();
        builder.Services.AddTransient<ViewModels.StudentScanViewModel>();
        builder.Services.AddTransient<ViewModels.TaskTabViewModel>();
        builder.Services.AddTransient<ViewModels.ControlModeViewModel>();
        builder.Services.AddTransient<ViewModels.SendCallViewModel>();

        // ===== Pages（Transient） =====
        // 新页面
        builder.Services.AddTransient<Pages.LoginPage>();
        builder.Services.AddTransient<Pages.AdminDashboardPage>();
        builder.Services.AddTransient<Pages.AdminTasksPage>();
        builder.Services.AddTransient<Pages.AdminUsersPage>();
        builder.Services.AddTransient<Pages.QRCodeGeneratePage>();
        builder.Services.AddTransient<Pages.StudentScanPage>();
        builder.Services.AddTransient<Pages.StudentHistoryPage>();
        builder.Services.AddTransient<Pages.AttendanceDetailPage>();
        builder.Services.AddTransient<Pages.ControlModePage>();
        builder.Services.AddTransient<Pages.SendCallPage>();

        // 保留旧页面兼容
        builder.Services.AddTransient<Pages.TaskListPage>();
        builder.Services.AddTransient<Pages.TaskDetailPage>();
        builder.Services.AddTransient<Pages.CreateSignInPage>();
        builder.Services.AddTransient<Pages.StudentGridPage>();
        builder.Services.AddTransient<Pages.SettingsPage>();

        return builder.Build();
    }
}
