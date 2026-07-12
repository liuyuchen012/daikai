using CheckIn.Client.Maui.Services;
using CheckIn.Client.Maui.ViewModels;

namespace CheckIn.Client.Maui.Pages;

public partial class CreateSignInPage : ContentPage
{
    public CreateSignInPage(ServerService serverService, MainViewModel mainViewModel)
    {
        InitializeComponent();
        BindingContext = new CreateSignInViewModel(serverService, mainViewModel);

        // Register the InverseBoolConverter if not already in global resources
        Resources.Add("InverseBoolConverter", new Converters.InverseBoolConverter());
    }
}
