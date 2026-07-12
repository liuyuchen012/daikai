namespace CheckIn.Client.Maui.Pages;

public partial class SignInResultPage : ContentPage
{
    private string _signUrl = "";

    public SignInResultPage(string shortCode, string serverIp, int serverPort)
    {
        InitializeComponent();
        _signUrl = $"http://{serverIp}:{serverPort}/s/{shortCode}";
        SignUrlLabel.Text = _signUrl;
        QrPlaceholder.IsVisible = true;
    }

    private async void OnCopyLinkClicked(object sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(_signUrl)) return;
        await Clipboard.Default.SetTextAsync(_signUrl);
        await DisplayAlert("提示", "签到链接已复制到剪贴板", "确定");
    }

    private async void OnCloseClicked(object sender, EventArgs e)
    {
        await Shell.Current.Navigation.PopModalAsync();
    }
}
