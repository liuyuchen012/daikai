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
        // 监听登录成功事件
        _viewModel.LoginSucceeded += OnLoginSucceeded;
    }

    protected override void OnDisappearing()
    {
        base.OnDisappearing();
        _viewModel.LoginSucceeded -= OnLoginSucceeded;
    }

    private void OnLoginSucceeded()
    {
        MainThread.BeginInvokeOnMainThread(async () =>
        {
            await Shell.Current.GoToAsync("//main");
        });
    }
}
