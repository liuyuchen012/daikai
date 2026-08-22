using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace CheckIn.Client;

/// <summary>
/// 课时划消管理窗口（壳）：标题栏 + 嵌入的 ClassHoursPanelControl + 底部状态栏。
/// 业务逻辑（学生管理 / 课时划消 / 排课）均在 ClassHoursPanelControl 中。
/// </summary>
public partial class ClassHoursWindow : Window
{
    public ClassHoursWindow()
    {
        InitializeComponent();
        // 窗口绑定与面板共享同一个 ClassHoursViewModel 实例
        DataContext = HoursPanel.HoursViewModel;
    }

    // ============ 标题栏 ============

    private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.LeftButton == MouseButtonState.Pressed && e.ClickCount == 1)
            DragMove();
    }

    private void Minimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;

    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    // ============ 菜单 ============

    private void MenuFile_Click(object sender, RoutedEventArgs e)
        => ShowMenu(sender as Button,
            new Mn("退出", "Exit"));

    private void MenuStudent_Click(object sender, RoutedEventArgs e)
        => ShowMenu(sender as Button,
            new Mn("添加学生", "AddStudent"),
            new Mn("删除选中学生", "DeleteSelected"));

    private void MenuSchedule_Click(object sender, RoutedEventArgs e)
        => ShowMenu(sender as Button,
            new Mn("复制当前日期排课…", "CopySchedule"),
            new Mn("设为不排课日/恢复", "ToggleOffDay"),
            new Mn("清空当前日期排课", "ClearSchedule"));

    private void MenuHelp_Click(object sender, RoutedEventArgs e)
        => ShowMenu(sender as Button,
            new Mn("使用说明", "Help"));

    // ---- 简易菜单弹出（仿 MainWindow 风格）----

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
            mi.Click += (s, e) =>
            {
                menu.IsOpen = false;
                HandleMenu(item.Tag);
            };
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
            case "Exit": Close(); break;
            case "AddStudent": HoursPanel.FocusAddStudent(); break;
            case "DeleteSelected": HoursPanel.DeleteSelectedStudent(); break;
            case "CopySchedule": HoursPanel.CopySchedule(); break;
            case "ToggleOffDay": HoursPanel.ToggleOffDay(); break;
            case "ClearSchedule": HoursPanel.ClearSchedule(); break;
            case "Help": HoursPanel.ShowHelp(); break;
        }
    }
}
