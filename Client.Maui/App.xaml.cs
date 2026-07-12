namespace CheckIn.Client.Maui;

public partial class App : Application
{
    public App()
    {
        InitializeComponent();
    }

    protected override Window CreateWindow(IActivationState? activationState)
    {
        return new Window(new AppShell())
        {
            Title = $"AgoraIn {CheckIn.Client.Models.AppConfig.Version}  作者: 刘宇晨",
            Width = 1200,
            Height = 800,
            MinimumWidth = 900,
            MinimumHeight = 600
        };
    }
}
