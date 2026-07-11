using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using CheckIn.Client.ViewModels;

namespace CheckIn.Client;

/// <summary>
/// 应用程序主窗口，包含任务树、标签栏、打卡面板和排名展示
/// 负责所有 UI 交互逻辑，包括菜单、对话框、标签高亮和窗口控制
/// </summary>
public partial class MainWindow : Window
{
    /// <summary>DWM 窗口圆角属性常量</summary>
    private const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;
    /// <summary>DWM 圆角模式值</summary>
    private const int DWMWCP_ROUND = 2;

    /// <summary>导入 DWM API 用于设置窗口圆角效果</summary>
    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

    private readonly MainViewModel _vm;

    /// <summary>
    /// 构造函数：初始化组件、创建 ViewModel 并绑定数据上下文
    /// 订阅 ActiveTab 变化以自动更新标签栏高亮
    /// </summary>
    public MainWindow()
    {
        InitializeComponent();
        _vm = new MainViewModel();
        DataContext = _vm;
        SourceInitialized += OnSourceInitialized;

        // 监听 ActiveTab 变化以更新标签高亮
        _vm.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(MainViewModel.ActiveTab))
                UpdateTabHighlighting();
        };
    }

    /// <summary>
    /// 窗口资源初始化完成后调用 DWM API 设置圆角效果
    /// </summary>
    private void OnSourceInitialized(object? sender, EventArgs e)
    {
        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd != IntPtr.Zero)
        {
            int preference = DWMWCP_ROUND;
            DwmSetWindowAttribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, ref preference, sizeof(int));
        }
    }

    // ---- 任务树选择 ----
    private TaskTreeNode? _selectedTreeNode;

    /// <summary>
    /// 任务树选中项变更处理：通知 ViewModel 打开对应的标签页
    /// </summary>
    private void TaskTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (e.NewValue is TaskTreeNode node)
        {
            _selectedTreeNode = node;
            _vm.OnTaskTreeSelected(node);
        }
    }

    /// <summary>
    /// 右键菜单弹出前：通过命中测试获取鼠标下的节点并选中
    /// </summary>
    private void TaskTree_ContextMenuOpening(object sender, ContextMenuEventArgs e)
    {
        var treeView = (TreeView)sender;
        var mousePos = Mouse.GetPosition(treeView);
        var hitResult = VisualTreeHelper.HitTest(treeView, mousePos);
        if (hitResult?.VisualHit == null) return;

        var treeViewItem = FindVisualParent<TreeViewItem>(hitResult.VisualHit);
        if (treeViewItem == null) return;

        treeViewItem.IsSelected = true;
        _selectedTreeNode = treeViewItem.DataContext as TaskTreeNode;
    }

    /// <summary>
    /// 向上遍历视觉树查找指定类型的父元素
    /// </summary>
    private static T? FindVisualParent<T>(DependencyObject child) where T : DependencyObject
    {
        var parent = VisualTreeHelper.GetParent(child);
        if (parent == null) return null;
        if (parent is T t) return t;
        return FindVisualParent<T>(parent);
    }

    // ---- 任务树右键菜单 ----
    /// <summary>
    /// 任务树右键菜单点击处理：根据 Tag 区分 打开/属性/重命名/删除 操作
    /// </summary>
    private void TaskTreeMenu_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem mi && mi.Tag is string action)
        {
            // 获取选中的节点
            if (_selectedTreeNode == null || _selectedTreeNode.IsFolder) return;
            if (_selectedTreeNode.Tab == null) return;

            switch (action)
            {
                case "Open":
                    _vm.SwitchToTab(_selectedTreeNode.Tab);
                    UpdateTabHighlighting();
                    break;
                case "Properties":
                    ShowTaskPropertiesDialog(_selectedTreeNode.Tab);
                    break;
                case "Rename":
                    ShowRenameTabDialog(_selectedTreeNode.Tab);
                    break;
                case "Delete":
                    DeleteTask(_selectedTreeNode.Tab);
                    break;
            }
        }
    }

    /// <summary>
    /// 显示任务属性对话框（名称、课程、行数、列数）
    /// </summary>
    private void ShowTaskPropertiesDialog(TaskTabViewModel tab)
    {
        var win = new Window
        {
            Title = $"任务属性 - {tab.TabDisplayName}",
            Width = 420, Height = 380,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false
        };

        var panel = new StackPanel { Margin = new Thickness(20) };
        string name = tab.Config.Name;
        string km = tab.Config.Km;
        int rows = tab.Config.ButtonRows;
        int cols = tab.Config.ButtonCols;

        // 名称
        var row1 = new WrapPanel { Margin = new Thickness(0, 0, 0, 10) };
        row1.Children.Add(new TextBlock { Text = "任务名称:", Width = 90, VerticalAlignment = VerticalAlignment.Center });
        var tb1 = new TextBox { Text = tab.Config.Name, Width = 240 };
        tb1.TextChanged += (_, _) => name = tb1.Text;
        row1.Children.Add(tb1);
        panel.Children.Add(row1);

        // 课程
        var row2 = new WrapPanel { Margin = new Thickness(0, 0, 0, 10) };
        row2.Children.Add(new TextBlock { Text = "课程名称:", Width = 90, VerticalAlignment = VerticalAlignment.Center });
        var tb2 = new TextBox { Text = tab.Config.Km, Width = 240 };
        tb2.TextChanged += (_, _) => km = tb2.Text;
        row2.Children.Add(tb2);
        panel.Children.Add(row2);

        // 行数
        var row3 = new WrapPanel { Margin = new Thickness(0, 0, 0, 10) };
        row3.Children.Add(new TextBlock { Text = "按钮行数:", Width = 90, VerticalAlignment = VerticalAlignment.Center });
        var tb3 = new TextBox { Text = tab.Config.ButtonRows.ToString(), Width = 240 };
        tb3.TextChanged += (_, _) => { if (int.TryParse(tb3.Text, out var n)) rows = n; };
        row3.Children.Add(tb3);
        panel.Children.Add(row3);

        // 列数
        var row4 = new WrapPanel { Margin = new Thickness(0, 0, 0, 10) };
        row4.Children.Add(new TextBlock { Text = "按钮列数:", Width = 90, VerticalAlignment = VerticalAlignment.Center });
        var tb4 = new TextBox { Text = tab.Config.ButtonCols.ToString(), Width = 240 };
        tb4.TextChanged += (_, _) => { if (int.TryParse(tb4.Text, out var n)) cols = n; };
        row4.Children.Add(tb4);
        panel.Children.Add(row4);

        // 提示
        panel.Children.Add(new TextBlock { Text = "提示：修改行数/列数后需重新导入学生名单",
            Foreground = new SolidColorBrush(Colors.Gray), FontSize = 11, Margin = new Thickness(0, 0, 0, 15) });

        // 按钮
        var btnPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 10, 0, 0) };
        var okBtn = new Button { Content = "保存", Width = 80, Height = 32, Margin = new Thickness(5),
            Background = new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)), Foreground = Brushes.White,
            BorderThickness = new Thickness(0), Cursor = Cursors.Hand };
        okBtn.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(name)) { ModernDialog.Alert("名称不能为空"); return; }
            tab.Config.Name = name.Trim();
            tab.Config.Km = km.Trim();
            tab.Config.ButtonRows = Math.Max(1, rows);
            tab.Config.ButtonCols = Math.Max(1, cols);
            tab.SaveConfig();
            tab.NotifyPropertyChanged(nameof(TaskTabViewModel.TabDisplayName));
            tab.NotifyPropertyChanged(nameof(TaskTabViewModel.Config));
            _vm.RefreshTaskTree();
            ModernDialog.Alert("属性已保存");
            win.Close();
        };
        var cancelBtn = new Button { Content = "取消", Width = 80, Height = 32, Margin = new Thickness(5) };
        cancelBtn.Click += (_, _) => win.Close();
        btnPanel.Children.Add(okBtn); btnPanel.Children.Add(cancelBtn);
        panel.Children.Add(btnPanel);

        win.Content = panel;
        win.ShowDialog();
    }

    /// <summary>
    /// 删除任务：需要两次确认，防止误删除
    /// </summary>
    private void DeleteTask(TaskTabViewModel tab)
    {
        var result = ModernDialog.Confirm(
            $"确定删除任务「{tab.TabDisplayName}」？\n\n" +
            "⚠️ 此操作将删除该任务的所有数据，且无法恢复！\n\n" +
            "建议先导出备份后再删除。",
            "删除任务");

        if (!result) return;

        // 二次确认
        var confirm = ModernDialog.Confirm("再次确认：真的要删除吗？此操作不可撤销！", "最后确认");
        if (!confirm) return;

        try
        {
            _vm.DeleteTask(tab);
            UpdateTabHighlighting();
            ModernDialog.Alert($"任务「{tab.TabDisplayName}」已删除");
        }
        catch (Exception ex)
        {
            ModernDialog.Alert($"删除失败：{ex.Message}");
        }
    }

    // ---- 标签栏操作 ----
    /// <summary>
    /// 标签栏鼠标点击：单击切换标签页，双击打开重命名对话框
    /// </summary>
    private void Tab_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is Border { Tag: TaskTabViewModel tab })
        {
            if (e.ClickCount == 2)
            {
                e.Handled = true;
                ShowRenameTabDialog(tab);
            }
            else
            {
                _vm.SwitchToTab(tab);
                UpdateTabHighlighting();
            }
        }
    }

    /// <summary>
    /// 标签关闭按钮点击：关闭当前标签页（数据保留在磁盘）
    /// </summary>
    private void TabClose_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn)
        {
            var tab = btn.DataContext as TaskTabViewModel;
            _vm.CloseTabCommand.Execute(tab);
            UpdateTabHighlighting();
        }
    }

    /// <summary>
    /// 显示重命名任务对话框
    /// </summary>
    private void ShowRenameTabDialog(TaskTabViewModel tab)
    {
        var win = new Window
        {
            Title = "重命名任务", Width = 360, Height = 170,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false
        };

        var panel = new StackPanel { Margin = new Thickness(20) };
        string newName = tab.Config.Name;

        var row = new WrapPanel { Margin = new Thickness(0, 0, 0, 12) };
        row.Children.Add(new TextBlock { Text = "任务名称:", Width = 70, VerticalAlignment = VerticalAlignment.Center });
        var tb = new TextBox { Text = tab.Config.Name, Width = 220 };
        tb.SelectAll();
        tb.TextChanged += (_, _) => newName = tb.Text;
        row.Children.Add(tb);
        panel.Children.Add(row);

        var btnPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center };
        var okBtn = new Button { Content = "确定", Width = 80, Height = 32, Margin = new Thickness(5),
            Background = new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)), Foreground = Brushes.White,
            BorderThickness = new Thickness(0), Cursor = Cursors.Hand };
        okBtn.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(newName)) { ModernDialog.Alert("名称不能为空"); return; }
            _vm.RenameTab(tab, newName);
            UpdateTabHighlighting();
            win.Close();
        };
        var cancelBtn = new Button { Content = "取消", Width = 80, Height = 32, Margin = new Thickness(5) };
        cancelBtn.Click += (_, _) => win.Close();
        btnPanel.Children.Add(okBtn); btnPanel.Children.Add(cancelBtn);
        panel.Children.Add(btnPanel);

        win.Content = panel;
        win.ShowDialog();
    }

    /// <summary>
    /// 新建标签按钮点击事件
    /// </summary>
    private void AddTab_Click(object sender, RoutedEventArgs e)
    {
        ShowNewTaskDialog();
    }

    /// <summary>
    /// 显示新建打卡任务对话框
    /// </summary>
    private void ShowNewTaskDialog()
    {
        var win = new Window
        {
            Title = "新建打卡任务", Width = 380, Height = 220,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false
        };

        var panel = new StackPanel { Margin = new Thickness(20) };
        string taskName = ""; string taskKm = "";

        var row1 = new WrapPanel { Margin = new Thickness(0, 0, 0, 10) };
        row1.Children.Add(new TextBlock { Text = "任务名称:", Width = 80, VerticalAlignment = VerticalAlignment.Center });
        var tb1 = new TextBox { Text = $"任务{_vm.Tabs.Count + 1}", Width = 230 };
        tb1.TextChanged += (_, _) => taskName = tb1.Text;
        taskName = tb1.Text;
        row1.Children.Add(tb1);
        panel.Children.Add(row1);

        var row2 = new WrapPanel { Margin = new Thickness(0, 0, 0, 10) };
        row2.Children.Add(new TextBlock { Text = "课程名称:", Width = 80, VerticalAlignment = VerticalAlignment.Center });
        var tb2 = new TextBox { Text = "数学", Width = 230 };
        tb2.TextChanged += (_, _) => taskKm = tb2.Text;
        taskKm = tb2.Text;
        row2.Children.Add(tb2);
        panel.Children.Add(row2);

        var btnPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 10, 0, 0) };
        var okBtn = new Button { Content = "创建", Width = 80, Height = 32, Margin = new Thickness(5),
            Background = new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)), Foreground = Brushes.White,
            BorderThickness = new Thickness(0), Cursor = Cursors.Hand };
        okBtn.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(taskName)) { ModernDialog.Alert("请输入任务名称"); return; }
            _vm.AddTab(taskName.Trim(), taskKm.Trim());
            UpdateTabHighlighting();
            win.Close();
        };
        var cancelBtn = new Button { Content = "取消", Width = 80, Height = 32, Margin = new Thickness(5) };
        cancelBtn.Click += (_, _) => win.Close();
        btnPanel.Children.Add(okBtn); btnPanel.Children.Add(cancelBtn);
        panel.Children.Add(btnPanel);

        win.Content = panel;
        win.ShowDialog();
    }

    // ---- 标签高亮 ----
    /// <summary>
    /// 通过视觉树遍历更新标签栏高亮状态（活跃标签白色背景）
    /// </summary>
    private void UpdateTabHighlighting()
    {
        // 通过视觉树遍历标签 Border 元素来更新高亮
        // 由于 ItemsControl 虚拟化，在布局更新后调用
        Dispatcher.BeginInvoke(new Action(() =>
        {
            try
            {
                // 简单方式：直接查找所有标签 Border
                FindTabBorders(this);
            }
            catch { }
        }), System.Windows.Threading.DispatcherPriority.Loaded);
    }

    /// <summary>
    /// 递归查找所有标签 Border 元素并更新背景色
    /// </summary>
    private void FindTabBorders(DependencyObject parent)
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is Border bdr && bdr.Tag is TaskTabViewModel tab)
            {
                bdr.Background = tab == _vm.ActiveTab
                    ? new SolidColorBrush(Colors.White)
                    : Brushes.Transparent;
                // 活跃标签文字颜色变深
                var textBlocks = FindVisualChildren<TextBlock>(bdr);
                foreach (var tb in textBlocks)
                {
                    tb.Foreground = tab == _vm.ActiveTab
                        ? new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))
                        : new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55));
                    tb.FontWeight = tab == _vm.ActiveTab ? FontWeights.SemiBold : FontWeights.Normal;
                }
            }
            FindTabBorders(child);
        }
    }

    /// <summary>
    /// 递归查找指定类型的所有可视子元素
    /// </summary>
    private static IEnumerable<T> FindVisualChildren<T>(DependencyObject parent) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T t) yield return t;
            foreach (var c in FindVisualChildren<T>(child)) yield return c;
        }
    }

    // ---- 窗口控制（无边框窗口的拖拽和最小化/最大化/关闭） ----
    /// <summary>标题栏拖拽移动窗口</summary>
    private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e) { if (e.ChangedButton == MouseButton.Left) DragMove(); }
    /// <summary>最小化窗口</summary>
    private void Minimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
    /// <summary>切换最大化/还原窗口</summary>
    private void Maximize_Click(object sender, RoutedEventArgs e) { WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized; }
    /// <summary>关闭窗口</summary>
    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    // ---- 现代化圆角弹出菜单 ----
    /// <summary>菜单项悬停背景色</summary>
    private static readonly SolidColorBrush _hBg = new(Color.FromRgb(0xe8, 0xf0, 0xfe));
    /// <summary>菜单项文字颜色</summary>
    private static readonly SolidColorBrush _fg = new(Color.FromRgb(0x33, 0x33, 0x33));

    /// <summary>菜单项数据模型</summary>
    private class Mn
    {
        public string? Header { get; set; }
        public Action? Act { get; set; }
        public bool IsSep { get; set; }
    }
    /// <summary>创建分隔线菜单项</summary>
    private static Mn Sep() => new() { IsSep = true };
    /// <summary>创建普通菜单项</summary>
    private static Mn I(string h, Action a) => new() { Header = h, Act = a };

    /// <summary>文件菜单：导出/导入/清空/新建/重命名/退出</summary>
    private void MenuFile_Click(object sender, RoutedEventArgs e) => ShowMenu(sender,
        I("导出打卡数据", () => _vm.ActiveTab?.ExportCommand.Execute(null)),
        I("导入打卡数据", () => _vm.ActiveTab?.ImportCommand.Execute(null)),
        Sep(),
        I("清空打卡记录", () => _vm.ActiveTab?.ClearAllCommand.Execute(null)),
        Sep(),
        I("新建任务", () => ShowNewTaskDialog()),
        I("重命名任务", () => { if (_vm.ActiveTab != null) ShowRenameTabDialog(_vm.ActiveTab); }),
        I("退出", () => _vm.ExitCommand.Execute(null)));

    /// <summary>远程菜单：服务器设置/状态检查/加载数据/同步数据</summary>
    private void MenuRemote_Click(object sender, RoutedEventArgs e) => ShowMenu(sender,
        I("远程服务器设置", () => _vm.ShowRemoteSettingsCommand.Execute(null)),
        I("检查服务器状态", () => _vm.ActiveTab?.CheckServerStatusCommand.Execute(null)),
        I("从服务器加载数据", () => _vm.ActiveTab?.LoadFromServerCommand.Execute(null)),
        I("同步数据到服务器", () => _vm.ActiveTab?.SyncToServerCommand.Execute(null)));

    /// <summary>设置菜单：任务设置/管理员设置</summary>
    private void MenuSettings_Click(object sender, RoutedEventArgs e) => ShowMenu(sender,
        I("任务设置", () => _vm.ActiveTab?.ShowAdminSettingsCommand.Execute(null)),
        Sep(),
        I("管理员设置", () => _vm.ShowAdminSettingsCommand.Execute(null)));

    /// <summary>帮助菜单：Github/版本检查/关于</summary>
    private void MenuHelp_Click(object sender, RoutedEventArgs e) => ShowMenu(sender,
        I("Github", () => _vm.OpenGithubCommand.Execute(null)),
        I("检查版本列表", () => _vm.CheckVersionCommand.Execute(null)),
        Sep(),
        I("关于", () => _vm.ShowAboutCommand.Execute(null)));

    /// <summary>
    /// 显示现代化圆角弹出菜单（使用 Popup 实现，非系统原生菜单）
    /// </summary>
    private static void ShowMenu(object sender, params Mn[] items)
    {
        if (sender is not Button btn) return;

        var popup = new Popup
        {
            PlacementTarget = btn, Placement = PlacementMode.Bottom,
            StaysOpen = false, AllowsTransparency = true,
            PopupAnimation = PopupAnimation.Slide
        };

        var bdr = new Border
        {
            CornerRadius = new CornerRadius(10),
            Background = new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe0, 0xe0)),
            BorderThickness = new Thickness(1), Padding = new Thickness(4),
            Effect = new System.Windows.Media.Effects.DropShadowEffect
            { BlurRadius = 14, ShadowDepth = 2, Color = Color.FromArgb(0x40, 0, 0, 0), Opacity = 0.3 }
        };

        var stack = new StackPanel();
        foreach (var m in items)
        {
            if (m.IsSep)
            {
                stack.Children.Add(new Separator
                { Background = new SolidColorBrush(Color.FromRgb(0xe8, 0xe8, 0xe8)), Margin = new Thickness(8, 3, 8, 3), Height = 1 });
                continue;
            }

            var ib = new Border
            {
                CornerRadius = new CornerRadius(6), Background = Brushes.Transparent,
                Padding = new Thickness(14, 7, 40, 7), Cursor = Cursors.Hand,
                Child = new TextBlock { Text = m.Header, Foreground = _fg, FontSize = 13 }
            };
            ib.MouseEnter += (_, _) => ib.Background = _hBg;
            ib.MouseLeave += (_, _) => ib.Background = Brushes.Transparent;
            ib.MouseLeftButtonUp += (_, _) => { popup.IsOpen = false; m.Act?.Invoke(); };
            stack.Children.Add(ib);
        }

        bdr.Child = stack;
        popup.Child = bdr;
        popup.IsOpen = true;
    }
}
