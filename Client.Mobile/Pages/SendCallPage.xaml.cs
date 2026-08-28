using CheckIn.Client.Mobile.ViewModels;

namespace CheckIn.Client.Mobile.Pages;

public partial class SendCallPage : ContentPage
{
    private readonly SendCallViewModel _viewModel;

    public SendCallPage(SendCallViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        BindingContext = viewModel;

        _viewModel.SendCompleted += OnSendCompleted;
        _viewModel.CloseRequested += () =>
        {
            MainThread.BeginInvokeOnMainThread(async () => await Navigation.PopModalAsync());
        };
    }

    /// <summary>弹窗前设置目标设备</summary>
    public void Initialize(IReadOnlyList<CallableDevice> targets) => _viewModel.SetTargets(targets);

    private async void OnSendCompleted(int ok, List<string> failures)
    {
        var total = _viewModel.TargetCount;
        var msg = $"已向 {ok}/{total} 台设备发送呼叫";
        if (failures.Count > 0) msg += "\n\n" + string.Join("\n", failures);
        await DisplayAlertAsync(ok == total ? "发送成功" : "发送完成（部分失败）", msg, "确定");
        if (ok > 0) await Navigation.PopModalAsync();
    }
}
