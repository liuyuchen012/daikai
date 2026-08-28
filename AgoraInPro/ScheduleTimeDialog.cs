using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace CheckIn.Client;

/// <summary>
/// 添加排课对话框：必须填写上课时间与下课时间（均无默认值），校验通过后才能提交。
/// </summary>
public class ScheduleTimeDialog : Window
{
    private readonly TextBox _startBox;
    private readonly TextBox _endBox;

    /// <summary>上课时间（HH:mm）</summary>
    public string StartTime => _startBox.Text.Trim();

    /// <summary>下课时间（HH:mm）</summary>
    public string EndTime => _endBox.Text.Trim();

    public ScheduleTimeDialog(string studentName)
    {
        Title = "添加排课";
        Width = 400;
        Height = 280;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        ResizeMode = ResizeMode.NoResize;
        ShowInTaskbar = false;
        WindowStyle = WindowStyle.None;
        AllowsTransparency = true;
        Background = Brushes.Transparent;

        var grid = new Grid { Margin = new Thickness(24) };
        grid.RowDefinitions.Add(new RowDefinition());
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        // 标题
        grid.Children.Add(new TextBlock
        {
            Text = "添加排课",
            FontSize = 16,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)),
            Margin = new Thickness(0, 0, 0, 4)
        });

        // 学生提示
        grid.Children.Add(new TextBlock
        {
            Text = $"为「{studentName}」设置上课与下课时间（必填，无默认值）：",
            FontSize = 12,
            Foreground = new SolidColorBrush(Color.FromRgb(0x88, 0x88, 0x88)),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8)
        });
        Grid.SetRow(grid.Children[^1], 1);

        // 上课时间
        _startBox = NewBox("如 08:00");
        _startBox.TextChanged += (_, _) => _startBox.ToolTip = "上课时间，格式 HH:mm";
        grid.Children.Add(Row("上课时间：", _startBox));
        Grid.SetRow(grid.Children[^1], 2);

        // 下课时间
        _endBox = NewBox("如 10:00");
        grid.Children.Add(Row("下课时间：", _endBox));
        Grid.SetRow(grid.Children[^1], 3);

        // 说明
        grid.Children.Add(new TextBlock
        {
            Text = "系统将按课程实际时长（下课-上课）× 每小时消耗课时计算本次消耗。",
            FontSize = 11,
            Foreground = new SolidColorBrush(Color.FromRgb(0x99, 0x99, 0x99)),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 8, 0, 0)
        });
        Grid.SetRow(grid.Children[^1], 4);

        // 按钮
        var cancelBtn = NewBtn("取消", Color.FromRgb(0xf0, 0xf0, 0xf0), Color.FromRgb(0x55, 0x55, 0x55));
        var okBtn = NewBtn("确认添加", Color.FromRgb(0x42, 0x85, 0xf4), Color.FromRgb(0xff, 0xff, 0xff));
        cancelBtn.Click += (_, _) => { DialogResult = false; Close(); };
        okBtn.Click += (_, _) =>
        {
            var start = StartTime;
            var end = EndTime;
            if (string.IsNullOrWhiteSpace(start) || string.IsNullOrWhiteSpace(end))
            {
                ModernDialog.Alert("请填写上课时间和下课时间，二者均不能为空。", "时间未填写");
                return;
            }
            if (!DateTime.TryParse(start, out var s) || !DateTime.TryParse(end, out var e))
            {
                ModernDialog.Alert("时间格式无效，请使用 HH:mm（如 08:30）。", "时间格式错误");
                return;
            }
            if (e == s)
            {
                ModernDialog.Alert("下课时间不能等于上课时间。", "时间范围错误");
                return;
            }
            DialogResult = true;
            Close();
        };

        var btnPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 16, 0, 0)
        };
        btnPanel.Children.Add(cancelBtn);
        btnPanel.Children.Add(okBtn);
        grid.Children.Add(btnPanel);
        Grid.SetRow(btnPanel, 5);

        Content = new Border
        {
            CornerRadius = new CornerRadius(12),
            Background = new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe0, 0xe0)),
            BorderThickness = new Thickness(1),
            Child = grid
        };
    }

    private static UIElement Row(string label, UIElement input)
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 10, 0, 0) };
        panel.Children.Add(new TextBlock
        {
            Text = label,
            FontSize = 13,
            Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)),
            VerticalAlignment = VerticalAlignment.Center,
            Width = 80
        });
        panel.Children.Add(input);
        return panel;
    }

    private static TextBox NewBox(string watermark)
    {
        var box = new TextBox
        {
            Width = 150,
            Height = 30,
            FontSize = 13,
            VerticalContentAlignment = VerticalAlignment.Center,
            ToolTip = watermark,
            Tag = watermark
        };
        box.GotFocus += (_, _) =>
        {
            if (box.Text == watermark)
                box.Text = "";
        };
        box.LostFocus += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(box.Text))
                box.Text = watermark;
        };
        box.Text = watermark;
        return box;
    }

    private static Button NewBtn(string text, Color bg, Color fg)
    {
        var btn = new Button
        {
            Content = text,
            Width = 90,
            Height = 32,
            Cursor = Cursors.Hand,
            FontSize = 13,
            Foreground = new SolidColorBrush(fg),
            Background = new SolidColorBrush(bg),
            BorderThickness = new Thickness(0),
            Margin = new Thickness(5, 0, 0, 0)
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
