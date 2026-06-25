using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace CheckIn.Client;

public static class ModernDialog
{
    public static bool Confirm(string message, string title = "确认")
    {
        var result = false;
        var win = CreateWindow(title, 400, 200);

        var grid = new Grid { Margin = new Thickness(24) };
        grid.RowDefinitions.Add(new RowDefinition());
        grid.RowDefinitions.Add(new RowDefinition());
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        grid.Children.Add(new TextBlock
        {
            Text = title, FontSize = 16, FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)),
            Margin = new Thickness(0, 0, 0, 12), VerticalAlignment = VerticalAlignment.Top
        });

        grid.Children.Add(new TextBlock
        {
            Text = message, FontSize = 14, TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)),
            Margin = new Thickness(0, 0, 0, 16)
        });
        Grid.SetRow(grid.Children[^1], 1);

        var cancelBtn = NewBtn("取消", new SolidColorBrush(Color.FromRgb(0xf0, 0xf0, 0xf0)),
            new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)));
        var okBtn = NewBtn("确定", new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)),
            new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)));

        cancelBtn.Click += (_, _) => { result = false; win.Close(); };
        okBtn.Click += (_, _) => { result = true; win.Close(); };

        var btnPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        btnPanel.Children.Add(cancelBtn);
        btnPanel.Children.Add(okBtn);
        Grid.SetRow(btnPanel, 2);
        grid.Children.Add(btnPanel);

        win.Content = new Border
        {
            CornerRadius = new CornerRadius(12),
            Background = new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe0, 0xe0)),
            BorderThickness = new Thickness(1),
            Child = grid
        };
        win.ShowDialog();
        return result;
    }

    public static void Alert(string message, string title = "提示")
    {
        var win = CreateWindow(title, 380, 170);

        var grid = new Grid { Margin = new Thickness(24) };
        grid.RowDefinitions.Add(new RowDefinition());
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        grid.Children.Add(new TextBlock
        {
            Text = title, FontSize = 16, FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)),
            Margin = new Thickness(0, 0, 0, 12)
        });

        grid.Children.Add(new TextBlock
        {
            Text = message, FontSize = 14, TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)),
            Margin = new Thickness(0, 0, 0, 16)
        });
        Grid.SetRow(grid.Children[^1], 1);

        var okBtn = NewBtn("确定", new SolidColorBrush(Color.FromRgb(0x42, 0x85, 0xf4)),
            new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)));
        okBtn.Click += (_, _) => win.Close();

        var btnPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Bottom
        };
        btnPanel.Children.Add(okBtn);

        Grid.SetRow(btnPanel, 2);
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.Children.Add(btnPanel);

        win.Content = new Border
        {
            CornerRadius = new CornerRadius(12),
            Background = new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe0, 0xe0)),
            BorderThickness = new Thickness(1),
            Child = grid
        };
        win.ShowDialog();
    }

    private static Window CreateWindow(string title, int w, int h)
    {
        var win = new Window
        {
            Title = title, Width = w, Height = h,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false,
            WindowStyle = WindowStyle.None, AllowsTransparency = true,
            Background = Brushes.Transparent, Topmost = true
        };
        try
        {
            var hwnd = new WindowInteropHelper(win).EnsureHandle();
            int p = 2;
            NativeMethods.DwmSetWindowAttribute(hwnd, 33, ref p, sizeof(int));
        }
        catch { }
        return win;
    }

    private static Button NewBtn(string text, Brush bg, Brush fg)
    {
        var btn = new Button
        {
            Content = text, Width = 80, Height = 32, Cursor = Cursors.Hand,
            FontSize = 13, Foreground = fg, Background = bg,
            BorderThickness = new Thickness(0), Margin = new Thickness(5, 0, 0, 0)
        };
        var template = new ControlTemplate(typeof(Button));
        var border = new FrameworkElementFactory(typeof(Border));
        border.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
        border.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
        var presenter = new FrameworkElementFactory(typeof(ContentPresenter));
        presenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
        presenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
        border.AppendChild(presenter);
        template.VisualTree = border;
        var trigger = new Trigger { Property = Button.IsMouseOverProperty, Value = true };
        trigger.Setters.Add(new Setter(Border.OpacityProperty, 0.85));
        template.Triggers.Add(trigger);
        btn.Template = template;
        return btn;
    }
}

internal static class NativeMethods
{
    [System.Runtime.InteropServices.DllImport("dwmapi.dll", PreserveSig = true)]
    internal static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
