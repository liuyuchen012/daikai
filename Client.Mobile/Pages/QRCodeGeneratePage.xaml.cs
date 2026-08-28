using CheckIn.Client.Mobile.ViewModels;

namespace CheckIn.Client.Mobile.Pages;

public partial class QRCodeGeneratePage : ContentPage
{
    private readonly QRCodeGenerateViewModel _viewModel;

    public QRCodeGeneratePage(QRCodeGenerateViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        BindingContext = viewModel;
    }

    protected override async void OnAppearing()
    {
        base.OnAppearing();
        await _viewModel.LoadDevicesAsync();
    }

    private void OnDeviceSelected(object? sender, EventArgs e)
    {
        if (sender is Picker picker &&
            picker.SelectedIndex >= 0 &&
            picker.SelectedItem is DevicePickerItem item)
        {
            _viewModel.SelectedDeviceUuid = item.Uuid;
            _viewModel.SelectedDeviceName = item.Name;
        }
    }
}
