using CheckIn.Client.Mobile.ViewModels;
using Microsoft.Extensions.DependencyInjection;

namespace CheckIn.Client.Mobile.Pages;

public partial class ControlModePage : ContentPage
{
    private readonly ControlModeViewModel _viewModel;
    private readonly IServiceProvider _services;

    public ControlModePage(ControlModeViewModel viewModel, IServiceProvider services)
    {
        InitializeComponent();
        _viewModel = viewModel;
        _services = services;
        BindingContext = viewModel;

        _viewModel.SendCallRequested += async (targets) =>
        {
            try
            {
                var page = _services.GetRequiredService<SendCallPage>();
                page.Initialize(targets);
                await Navigation.PushModalAsync(page);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Open send call page error: {ex.Message}");
            }
        };
    }

    protected override async void OnAppearing()
    {
        base.OnAppearing();
        await _viewModel.LoadDevicesAsync();
    }
}
