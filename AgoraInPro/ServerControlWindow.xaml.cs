using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CheckIn.Client.Models;
using CheckIn.Client.ViewModels;

namespace CheckIn.Client;

public partial class ServerControlWindow : Window
{
    private readonly ServerControlViewModel _vm;

    public ServerControlWindow()
    {
        InitializeComponent();
        _vm = new ServerControlViewModel();
        DataContext = _vm;
    }

    private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.LeftButton == MouseButtonState.Pressed && e.ClickCount == 1)
            DragMove();
    }

    private void Minimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;

    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    // ============ 菜单 ============

    private void MenuConnect_Click(object sender, RoutedEventArgs e)
        => ShowMenu(sender as Button,
            new Mn("登录", "Login"),
            new Mn("重新登录", "Relogin"),
            new Mn("退出登录", "Logout"));

    private void MenuData_Click(object sender, RoutedEventArgs e)
        => ShowMenu(sender as Button,
            new Mn("刷新仪表盘", "Dashboard"),
            new Mn("刷新设备", "Devices"),
            new Mn("刷新任务", "Tasks"),
            new Mn("刷新考勤", "Attendance"),
            new Mn("刷新历史", "History"),
            new Mn("刷新用户", "Users"));

    private void MenuAction_Click(object sender, RoutedEventArgs e)
        => ShowMenu(sender as Button,
            new Mn("新建用户…", "CreateUser"),
            new Mn("复制服务器地址", "CopyUrl"));

    private void MenuHelp_Click(object sender, RoutedEventArgs e)
        => ShowMenu(sender as Button, new Mn("使用说明", "Help"));

    public class Mn
    {
        public string Header { get; }
        public string Tag { get; }
        public Mn(string header, string tag) { Header = header; Tag = tag; }
    }

    private void ShowMenu(Button? anchor, params Mn[] items)
    {
        if (anchor == null) return;
        var menu = new ContextMenu { Style = (Style)FindResource("ModernContextMenu") };
        foreach (var item in items)
        {
            var mi = new MenuItem { Header = item.Header, Tag = item.Tag };
            mi.Click += (s, e) => { menu.IsOpen = false; HandleMenu(item.Tag); };
            menu.Items.Add(mi);
        }
        menu.PlacementTarget = anchor;
        menu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
        menu.IsOpen = true;
    }

    private void HandleMenu(string tag)
    {
        switch (tag)
        {
            case "Login": DoLogin(); break;
            case "Relogin": _vm.Logout(); break;
            case "Logout": _vm.Logout(); break;
            case "Dashboard": _ = _vm.LoadDashboardAsync(); break;
            case "Devices": _ = _vm.LoadDevicesAsync(); break;
            case "Tasks": _ = _vm.LoadTasksAsync(); break;
            case "Attendance": _ = _vm.LoadAttendanceAsync(); break;
            case "History": _ = _vm.LoadHistoryAsync(); break;
            case "Users": _ = _vm.LoadUsersAsync(); break;
            case "CreateUser": CreateUser_Click(this, new RoutedEventArgs()); break;
            case "CopyUrl":
                Clipboard.SetText(_vm.ServerUrl);
                _vm.StatusMessage = "服务器地址已复制";
                break;
            case "Help": ShowHelp(); break;
        }
    }

    private void ShowHelp()
    {
        MessageBox.Show(
            "【远程打卡服务器控制】\n" +
            "1. 输入服务器地址（如 http://192.168.1.100:5000）、用户名和密码登录。\n" +
            "2. 仪表盘：查看设备、在线数、今日签到、活跃任务等概览。\n" +
            "3. 设备管理：查看/重命名/删除打卡设备。\n" +
            "4. 任务管理：查看/重命名/关闭/删除签到任务。\n" +
            "5. 考勤与签到历史：查看学生打卡情况。\n" +
            "6. 用户管理（管理员）：新建/启用/禁用/删除用户。",
            "使用说明", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    // ============ 登录 ============

    private void Login_Click(object sender, RoutedEventArgs e) => DoLogin();

    private void DoLogin()
    {
        _vm.Password = PasswordBox?.Password ?? "";
        _ = _vm.LoginAsync().ContinueWith(t =>
        {
            if (!t.Result) return;
            Dispatcher.Invoke(async () =>
            {
                await _vm.LoadDashboardAsync();
                _ = _vm.LoadDevicesAsync();
                _ = _vm.LoadTasksAsync();
                _ = _vm.LoadAttendanceAsync();
                _ = _vm.LoadHistoryAsync();
                if (_vm.IsAdmin) _ = _vm.LoadUsersAsync();
            });
        });
    }

    // ============ 刷新 ============

    private void RefreshDevices_Click(object sender, RoutedEventArgs e) => _ = _vm.LoadDevicesAsync();
    private void RefreshTasks_Click(object sender, RoutedEventArgs e) => _ = _vm.LoadTasksAsync();
    private void RefreshAttendance_Click(object sender, RoutedEventArgs e) => _ = _vm.LoadAttendanceAsync();
    private void RefreshHistory_Click(object sender, RoutedEventArgs e) => _ = _vm.LoadHistoryAsync();
    private void RefreshUsers_Click(object sender, RoutedEventArgs e) => _ = _vm.LoadUsersAsync();

    // ============ 设备操作 ============

    private void RenameDevice_Click(object sender, RoutedEventArgs e)
    {
        if (GetDataContext<DeviceItem>(sender) is not { } device) return;
        var input = new InputDialog("重命名设备", "设备名称：", device.Name) { Owner = this };
        if (input.ShowDialog() == true && !string.IsNullOrWhiteSpace(input.Value))
            _ = _vm.RenameDeviceAsync(device, input.Value.Trim());
    }

    private void DeleteDevice_Click(object sender, RoutedEventArgs e)
    {
        if (GetDataContext<DeviceItem>(sender) is not { } device) return;
        if (MessageBox.Show($"确定删除设备「{device.Name}」？将同时删除其绑定的任务与签到数据。",
                "删除设备", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            _ = _vm.DeleteDeviceAsync(device);
    }

    // ============ 任务操作 ============

    private void RenameTask_Click(object sender, RoutedEventArgs e)
    {
        if (GetDataContext<RemoteTask>(sender) is not { } task) return;
        var input = new InputDialog("重命名任务", "任务名称：", task.Subject) { Owner = this };
        if (input.ShowDialog() == true && !string.IsNullOrWhiteSpace(input.Value))
            _ = _vm.RenameTaskAsync(task, input.Value.Trim());
    }

    private void CloseTask_Click(object sender, RoutedEventArgs e)
    {
        if (GetDataContext<RemoteTask>(sender) is not { } task) return;
        if (MessageBox.Show($"确定关闭任务「{task.Subject}」？", "关闭任务",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            _ = _vm.CloseTaskAsync(task);
    }

    private void DeleteTask_Click(object sender, RoutedEventArgs e)
    {
        if (GetDataContext<RemoteTask>(sender) is not { } task) return;
        if (MessageBox.Show($"确定删除任务「{task.Subject}」及其签到数据？", "删除任务",
                MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            _ = _vm.DeleteTaskAsync(task);
    }

    // ============ 用户操作 ============

    private void CreateUser_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new CreateUserDialog { Owner = this };
        if (dialog.ShowDialog() == true)
            _ = _vm.CreateUserAsync(dialog.Username, dialog.Password, dialog.Role, dialog.DisplayName);
    }

    private void ToggleUserActive_Click(object sender, RoutedEventArgs e)
    {
        if (GetDataContext<RemoteUserItem>(sender) is not { } user) return;
        var action = user.IsActive ? "禁用" : "启用";
        if (MessageBox.Show($"确定{action}用户「{user.Username}」？", "用户管理",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            _ = _vm.ToggleUserActiveAsync(user);
    }

    private void DeleteUser_Click(object sender, RoutedEventArgs e)
    {
        if (GetDataContext<RemoteUserItem>(sender) is not { } user) return;
        if (MessageBox.Show($"确定删除用户「{user.Username}」？", "删除用户",
                MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            _ = _vm.DeleteUserAsync(user);
    }

    private static T? GetDataContext<T>(object sender) where T : class
        => sender is FrameworkElement fe ? fe.DataContext as T : null;
}
