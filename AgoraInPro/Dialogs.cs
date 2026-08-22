using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace CheckIn.Client;

public class InputDialog : Window
{
    private readonly TextBox _inputBox;

    public string Value => _inputBox.Text;

    public InputDialog(string title, string label, string value = "")
    {
        Title = title;
        Width = 380;
        Height = 190;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        ShowInTaskbar = false;
        WindowStyle = WindowStyle.None;
        AllowsTransparency = true;
        Background = Brushes.Transparent;

        _inputBox = new TextBox { Text = value, Height = 32, FontSize = 14 };
        var okButton = NewButton("确定", Color.FromRgb(0x42, 0x85, 0xf4), Color.FromRgb(0xff, 0xff, 0xff));
        var cancelButton = NewButton("取消", Color.FromRgb(0xf0, 0xf0, 0xf0), Color.FromRgb(0x55, 0x55, 0x55));
        okButton.Click += (_, _) => { DialogResult = true; Close(); };
        cancelButton.Click += (_, _) => { DialogResult = false; Close(); };

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 16, 0, 0)
        };
        buttons.Children.Add(cancelButton);
        buttons.Children.Add(okButton);

        var panel = new StackPanel { Margin = new Thickness(24) };
        panel.Children.Add(new TextBlock
        {
            Text = label,
            FontSize = 13,
            Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)),
            Margin = new Thickness(0, 0, 0, 8)
        });
        panel.Children.Add(_inputBox);
        panel.Children.Add(buttons);

        Content = CreateSurface(panel);
        Loaded += (_, _) => { _inputBox.Focus(); _inputBox.SelectAll(); };
    }

    private static Border CreateSurface(UIElement child) => new()
    {
        CornerRadius = new CornerRadius(12),
        Background = Brushes.White,
        BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe0, 0xe0)),
        BorderThickness = new Thickness(1),
        Child = child
    };

    private static Button NewButton(string text, Color background, Color foreground) => new()
    {
        Content = text,
        Width = 80,
        Height = 32,
        FontSize = 13,
        Cursor = Cursors.Hand,
        Foreground = new SolidColorBrush(foreground),
        Background = new SolidColorBrush(background),
        BorderThickness = new Thickness(0),
        Margin = new Thickness(5, 0, 0, 0)
    };
}

public class CreateUserDialog : Window
{
    private readonly TextBox _usernameBox = new();
    private readonly PasswordBox _passwordBox = new();
    private readonly ComboBox _roleBox = new();
    private readonly TextBox _displayNameBox = new();

    public string Username => _usernameBox.Text.Trim();
    public string Password => _passwordBox.Password;
    public string Role => (_roleBox.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "user";
    public string DisplayName => string.IsNullOrWhiteSpace(_displayNameBox.Text) ? Username : _displayNameBox.Text.Trim();

    public CreateUserDialog()
    {
        Title = "新建用户";
        Width = 420;
        Height = 330;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        ShowInTaskbar = false;
        WindowStyle = WindowStyle.None;
        AllowsTransparency = true;
        Background = Brushes.Transparent;

        _roleBox.Items.Add(new ComboBoxItem { Content = "普通用户", Tag = "user", IsSelected = true });
        _roleBox.Items.Add(new ComboBoxItem { Content = "管理员", Tag = "admin" });
        _roleBox.Height = 32;
        _roleBox.FontSize = 14;

        var okButton = InputDialogButton("创建", Color.FromRgb(0x42, 0x85, 0xf4), Color.FromRgb(0xff, 0xff, 0xff));
        var cancelButton = InputDialogButton("取消", Color.FromRgb(0xf0, 0xf0, 0xf0), Color.FromRgb(0x55, 0x55, 0x55));
        okButton.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(Username) || string.IsNullOrEmpty(Password))
            {
                MessageBox.Show("请输入用户名和密码。", "新建用户", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            DialogResult = true;
            Close();
        };
        cancelButton.Click += (_, _) => { DialogResult = false; Close(); };

        var panel = new StackPanel { Margin = new Thickness(24) };
        panel.Children.Add(Field("用户名：", _usernameBox));
        panel.Children.Add(Field("密码：", _passwordBox));
        panel.Children.Add(Field("角色：", _roleBox));
        panel.Children.Add(Field("显示名：", _displayNameBox));

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 16, 0, 0)
        };
        buttons.Children.Add(cancelButton);
        buttons.Children.Add(okButton);
        panel.Children.Add(buttons);

        Content = new Border
        {
            CornerRadius = new CornerRadius(12),
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe0, 0xe0)),
            BorderThickness = new Thickness(1),
            Child = panel
        };
    }

    private static UIElement Field(string label, Control input)
    {
        input.Height = 30;
        input.FontSize = 14;
        var row = new Grid { Margin = new Thickness(0, 4, 0, 4) };
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(70) });
        row.ColumnDefinitions.Add(new ColumnDefinition());
        var labelBlock = new TextBlock
        {
            Text = label,
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55))
        };
        Grid.SetColumn(input, 1);
        row.Children.Add(labelBlock);
        row.Children.Add(input);
        return row;
    }

    private static Button InputDialogButton(string text, Color background, Color foreground) => new()
    {
        Content = text,
        Width = 80,
        Height = 32,
        FontSize = 13,
        Cursor = Cursors.Hand,
        Foreground = new SolidColorBrush(foreground),
        Background = new SolidColorBrush(background),
        BorderThickness = new Thickness(0),
        Margin = new Thickness(5, 0, 0, 0)
    };
}