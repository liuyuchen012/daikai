using CheckIn.Client.Mobile.ViewModels;

namespace CheckIn.Client.Mobile.Pages;

public partial class StudentScanPage : ContentPage
{
    private readonly StudentScanViewModel _viewModel;

    public StudentScanPage(StudentScanViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        BindingContext = viewModel;

        barcodeReader.BarcodesDetected += OnBarcodesDetected;
    }

    private void OnBarcodesDetected(object? sender, ZXing.Net.Maui.BarcodeDetectionEventArgs e)
    {
        if (e.Results?.Length > 0)
        {
            var firstResult = e.Results[0];
            MainThread.BeginInvokeOnMainThread(() =>
            {
                _viewModel.OnBarcodeDetected(firstResult.Value);
            });
        }
    }
}
