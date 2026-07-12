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
            // 切换到对应的 Tab
            await Shell.Current.GoToAsync($"//main/{page}");
        };
    }

    protected override async void OnAppearing()
    {
        base.OnAppearing();
        await _viewModel.LoadDashboardAsync();
    }
}
