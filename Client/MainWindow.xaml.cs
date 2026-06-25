using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using CheckIn.Client.ViewModels;

namespace CheckIn.Client;

public partial class MainWindow : Window
{
    private const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;
    private const int DWMWCP_ROUND = 2;

    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

    private readonly MainViewModel _vm;

    public MainWindow()
    {
        InitializeComponent();
        _vm = new MainViewModel();
        DataContext = _vm;
        SourceInitialized += OnSourceInitialized;
    }

    private void OnSourceInitialized(object? sender, EventArgs e)
    {
        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd != IntPtr.Zero)
        {
            int preference = DWMWCP_ROUND;
            DwmSetWindowAttribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, ref preference, sizeof(int));
        }
    }

    private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e) { if (e.ChangedButton == MouseButton.Left) DragMove(); }
    private void Minimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
    private void Maximize_Click(object sender, RoutedEventArgs e) { WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized; }
    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    // ---- 现代化圆角弹出菜单 ----
    private static readonly SolidColorBrush _hBg = new(Color.FromRgb(0xe8, 0xf0, 0xfe));
    private static readonly SolidColorBrush _fg = new(Color.FromRgb(0x33, 0x33, 0x33));

    // 菜单项标记
    private class Mn
    {
        public string? Header { get; set; }
        public Action? Act { get; set; }
        public bool IsSep { get; set; }
    }
    private static Mn Sep() => new() { IsSep = true };
    private static Mn I(string h, Action a) => new() { Header = h, Act = a };

    private void MenuFile_Click(object sender, RoutedEventArgs e) => ShowMenu(sender,
        I("导出打卡数据", () => _vm.ExportCommand.Execute(null)),
        I("导入打卡数据", () => _vm.ImportCommand.Execute(null)),
        Sep(),
        I("清空打卡记录", () => _vm.ClearAllCommand.Execute(null)),
        Sep(),
        I("退出", () => _vm.ExitCommand.Execute(null)));

    private void MenuRemote_Click(object sender, RoutedEventArgs e) => ShowMenu(sender,
        I("远程服务器设置", () => _vm.ShowRemoteSettingsCommand.Execute(null)),
        I("检查服务器状态", () => _vm.CheckServerStatusCommand.Execute(null)),
        I("从服务器加载数据", () => _vm.LoadFromServerCommand.Execute(null)),
        I("同步数据到服务器", () => _vm.SyncToServerCommand.Execute(null)));

    private void MenuSettings_Click(object sender, RoutedEventArgs e) => ShowMenu(sender,
        I("管理员设置", () => _vm.ShowAdminSettingsCommand.Execute(null)));

    private void MenuHelp_Click(object sender, RoutedEventArgs e) => ShowMenu(sender,
        I("Github", () => _vm.OpenGithubCommand.Execute(null)),
        I("检查版本列表", () => _vm.CheckVersionCommand.Execute(null)),
        Sep(),
        I("关于", () => _vm.ShowAboutCommand.Execute(null)));

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
