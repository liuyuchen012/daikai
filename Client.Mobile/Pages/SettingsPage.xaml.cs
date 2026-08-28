using CheckIn.Client.Mobile.ViewModels;

namespace CheckIn.Client.Mobile.Pages;

public partial class SettingsPage : ContentPage
{
    private MainViewModel _mainVm = null!;

    public SettingsPage()
    {
        InitializeComponent();
    }

    protected override void OnAppearing()
    {
        base.OnAppearing();
        _mainVm = IPlatformApplication.Current!.Services.GetRequiredService<MainViewModel>();

        // Load current settings
        ServerIpEntry.Text = _mainVm.Config.ServerIp;
        ServerPortEntry.Text = _mainVm.Config.ServerPort.ToString();
        ServerPasswordEntry.Text = _mainVm.Config.ServerPassword;
    }

    private async void OnSaveServerClicked(object? sender, EventArgs e)
    {
        var ip = ServerIpEntry.Text?.Trim();
        var portText = ServerPortEntry.Text?.Trim();
        var password = ServerPasswordEntry.Text?.Trim();

        if (string.IsNullOrWhiteSpace(ip))
        {
            await DisplayAlertAsync("Error", "Please enter server IP address", "OK");
            return;
        }

        if (!int.TryParse(portText, out int port) || port <= 0 || port > 65535)
        {
            await DisplayAlertAsync("Error", "Please enter a valid port number (1-65535)", "OK");
            return;
        }

        _mainVm.UpdateServerConfig(ip, port, password ?? "");
        await DisplayAlertAsync("Success", "Server settings saved", "OK");
    }

    private async void OnSaveAdminClicked(object? sender, EventArgs e)
    {
        var newPassword = NewPasswordEntry.Text?.Trim();
        var confirmPassword = ConfirmAdminPasswordEntry.Text?.Trim();

        if (!string.IsNullOrEmpty(newPassword))
        {
            if (newPassword != confirmPassword)
            {
                await DisplayAlertAsync("Error", "Passwords do not match", "OK");
                return;
            }

            // Verify current admin password if one is set
            if (!string.IsNullOrEmpty(_mainVm.Config.AdminPasswordHash))
            {
                string currentPwd = await DisplayPromptAsync(
                    "Admin Verification",
                    "Enter current admin password:",
                    "OK", "Cancel",
                    keyboard: Keyboard.Default);

                if (string.IsNullOrEmpty(currentPwd))
                    return;

                if (MainViewModel.HashPassword(currentPwd) != _mainVm.Config.AdminPasswordHash)
                {
                    await DisplayAlertAsync("Error", "Incorrect admin password", "OK");
                    return;
                }
            }

            _mainVm.Config.AdminPasswordHash = MainViewModel.HashPassword(newPassword);
            _mainVm.SaveGlobalConfig();
            await DisplayAlertAsync("Success", "Admin password updated", "OK");
        }
        else
        {
            await DisplayAlertAsync("Info", "No changes made (password left empty)", "OK");
        }
    }

    private async void OnCreateSignInClicked(object? sender, EventArgs e)
    {
        await Shell.Current.GoToAsync("createsignin");
    }
}
