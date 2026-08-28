using CheckIn.Client.Mobile.ViewModels;

namespace CheckIn.Client.Mobile.Pages;

public partial class AdminTasksPage : ContentPage
{
    private readonly AdminTasksViewModel _viewModel;

    public AdminTasksPage(AdminTasksViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        BindingContext = viewModel;

        _viewModel.ViewAttendanceRequested += async (shortCode, subject) =>
        {
            await Shell.Current.GoToAsync($"attendance?shortCode={shortCode}&subject={subject}");
        };
    }

    protected override async void OnAppearing()
    {
        base.OnAppearing();
        await _viewModel.LoadTasksAsync();
    }
}
