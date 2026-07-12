using CheckIn.Client.Mobile.ViewModels;

namespace CheckIn.Client.Mobile.Pages;

public partial class LoginPage : ContentPage
{
    private readonly LoginViewModel _viewModel;

    public LoginPage(LoginViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        BindingContext = viewModel;
    }

    protected override void OnAppearing()
    {
        base.OnAppearing();
        _viewModel.LoginSucceeded += OnLoginSucceeded;
        _viewModel.PropertyChanged += OnViewModelPropertyChanged;
    }

    protected override void OnDisappearing()
    {
        base.OnDisappearing();
        _viewModel.LoginSucceeded -= OnLoginSucceeded;
        _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
    }

    private void OnViewModelPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(LoginViewModel.IsLoading) ||
            e.PropertyName == nameof(LoginViewModel.CanLogin))
        {
            MainThread.BeginInvokeOnMainThread(() =>
            {
                btnLogin.IsEnabled = _viewModel.CanLogin;
            });
        }
    }

    private async void OnLoginClicked(object? sender, EventArgs e)
    {
        if (!_viewModel.CanLogin) return;

        btnLogin.IsEnabled = false;

        try
        {
            await _viewModel.LoginAsync();
        }
        catch (Exception ex)
        {
            _viewModel.ErrorMessage = $"发生错误: {ex.Message}";
        }
        finally
        {
            btnLogin.IsEnabled = true;
        }
    }

    private void OnLoginSucceeded()
    {
        MainThread.BeginInvokeOnMainThread(async () =>
        {
            try
            {
                await Shell.Current.GoToAsync("//main");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Navigation error: {ex.Message}");
            }
        });
    }
}
