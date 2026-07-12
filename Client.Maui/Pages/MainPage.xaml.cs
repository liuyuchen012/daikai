using CheckIn.Client.Models;
using CheckIn.Client.Maui.ViewModels;

namespace CheckIn.Client.Maui.Pages;

public partial class MainPage : ContentPage
{
    private readonly MainViewModel _viewModel;

    public MainPage(MainViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        BindingContext = _viewModel;
    }

    // ---- Menu Handlers ----
    private async void OnMenuFileClicked(object sender, EventArgs e)
    {
        var actions = new Dictionary<string, Action>
        {
            ["导出打卡数据"] = () => _viewModel.ActiveTab?.ExportCommand.Execute(null),
            ["导入打卡数据"] = () => _viewModel.ActiveTab?.ImportCommand.Execute(null),
            ["清空打卡记录"] = () => _viewModel.ActiveTab?.ClearAllCommand.Execute(null),
            ["新建任务"] = () => ShowNewTaskDialog(),
            ["重命名任务"] = () => ShowRenameDialog(),
        };

        var choice = await DisplayActionSheetAsync("文件", "取消", null, actions.Keys.ToArray());
        if (choice != null && actions.TryGetValue(choice, out var action))
            action();
    }

    private async void OnMenuRemoteClicked(object sender, EventArgs e)
    {
        var actions = new Dictionary<string, Action>
        {
            ["远程服务器设置"] = () => _viewModel.ShowRemoteSettingsCommand.Execute(null),
            ["创建签到"] = () => _viewModel.CreateSignInCommand.Execute(null),
            ["检查服务器状态"] = () => _viewModel.ActiveTab?.CheckServerStatusCommand.Execute(null),
            ["从服务器加载数据"] = () => _viewModel.ActiveTab?.LoadFromServerCommand.Execute(null),
            ["同步数据到服务器"] = () => _viewModel.ActiveTab?.SyncToServerCommand.Execute(null),
        };

        var choice = await DisplayActionSheetAsync("远程", "取消", null, actions.Keys.ToArray());
        if (choice != null && actions.TryGetValue(choice, out var action))
            action();
    }

    private async void OnMenuSettingsClicked(object sender, EventArgs e)
    {
        var actions = new Dictionary<string, Action>
        {
            ["任务设置"] = () => _viewModel.ActiveTab?.ShowAdminSettingsCommand.Execute(null),
            ["管理员设置"] = () => _viewModel.ShowAdminSettingsCommand.Execute(null),
        };

        var choice = await DisplayActionSheetAsync("设置", "取消", null, actions.Keys.ToArray());
        if (choice != null && actions.TryGetValue(choice, out var action))
            action();
    }

    private async void OnMenuHelpClicked(object sender, EventArgs e)
    {
        var actions = new Dictionary<string, Action>
        {
            ["GitHub"] = () => _viewModel.OpenGithubCommand.Execute(null),
            ["检查版本列表"] = () => _viewModel.CheckVersionCommand.Execute(null),
            ["关于"] = () => _viewModel.ShowAboutCommand.Execute(null),
        };

        var choice = await DisplayActionSheetAsync("帮助", "取消", null, actions.Keys.ToArray());
        if (choice != null && actions.TryGetValue(choice, out var action))
            action();
    }

    // ---- Toolbar buttons ----
    private void OnAddTabClicked(object sender, EventArgs e) => ShowNewTaskDialog();

    private async void OnCreateSignInClicked(object sender, EventArgs e)
        => _viewModel.CreateSignInCommand.Execute(null);

    // ---- Task Tree ----
    private void OnTaskTreeSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (e.CurrentSelection.FirstOrDefault() is TaskTreeNode node)
            _viewModel.OnTaskTreeSelected(node);
    }

    // ---- Student Tap ----
    private void OnStudentTapped(object? sender, TappedEventArgs e)
    {
        if (sender is BindableObject b && b.BindingContext is StudentModel student)
        {
            if (student.IsCheckedIn)
                _viewModel.ActiveTab?.CancelCheckInCommand.Execute(student);
            else
                _viewModel.ActiveTab?.CheckInCommand.Execute(student);
        }
    }

    // ---- Dialogs ----
    private async void ShowNewTaskDialog()
    {
        string name = await DisplayPromptAsync("新建打卡任务", "任务名称:", "创建", "取消", $"任务{_viewModel.Tabs.Count + 1}");
        if (string.IsNullOrWhiteSpace(name)) return;

        string km = await DisplayPromptAsync("新建打卡任务", "课程名称:", "创建", "取消", "数学");
        if (string.IsNullOrWhiteSpace(km)) return;

        _viewModel.AddTab(name.Trim(), km.Trim());
    }

    private async void ShowRenameDialog()
    {
        if (_viewModel.ActiveTab == null) return;
        string newName = await DisplayPromptAsync("重命名任务", "任务名称:", "确定", "取消",
            _viewModel.ActiveTab.Config.Name);
        if (string.IsNullOrWhiteSpace(newName)) return;
        _viewModel.RenameTab(_viewModel.ActiveTab, newName);
    }
}
