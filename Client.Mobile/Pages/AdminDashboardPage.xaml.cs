using CheckIn.Client.Mobile.ViewModels;

namespace CheckIn.Client.Mobile.Pages;

public partial class AdminDashboardPage : ContentPage
{
    private readonly AdminDashboardViewModel _viewModel;

    public AdminDashboardPage(AdminDashboardViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        BindingContext = viewModel;

        _viewModel.NavigateRequested += async (page) =>
        {
            await Shell.Current.GoToAsync($"//main/{page}");
        };

        _viewModel.DeviceTapped += async (uuid, name) =>
        {
            await Shell.Current.GoToAsync($"attendance?machineUuid={Uri.EscapeDataString(uuid)}&deviceName={Uri.EscapeDataString(name)}");
        };
    }

    protected override async void OnAppearing()
    {
        base.OnAppearing();
        await _viewModel.LoadDashboardAsync();
    }
}
