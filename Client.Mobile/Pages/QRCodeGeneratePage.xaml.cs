using CheckIn.Client.Mobile.ViewModels;

namespace CheckIn.Client.Mobile.Pages;

public partial class QRCodeGeneratePage : ContentPage
{
    public QRCodeGeneratePage(QRCodeGenerateViewModel viewModel)
    {
        InitializeComponent();
        BindingContext = viewModel;
    }
}
