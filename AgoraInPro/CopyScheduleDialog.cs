using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace CheckIn.Client;

/// <summary>
/// 复制排课对话框：选择源日期、目标日期（单日或范围）、是否跳过不排课日
/// </summary>
public class CopyScheduleDialog : Window
{
    private readonly DateTime _defaultFrom;
    private readonly DatePicker _fromPicker = new() { Width = 140 };
    private readonly RadioButton _singleRadio = new() { Content = "单日", IsChecked = true };
    private readonly RadioButton _rangeRadio = new() { Content = "日期范围" };
    private readonly DatePicker _toSinglePicker = new() { Width = 140 };
    private readonly DatePicker _rangeStartPicker = new() { Width = 120 };
    private readonly DatePicker _rangeEndPicker = new() { Width = 120 };
    private readonly CheckBox _skipOffBox = new() { Content = "跳过不排课日", IsChecked = true };
    private StackPanel? _rangePanel;

    public DateTime FromDate => _fromPicker.SelectedDate ?? _defaultFrom;
    public bool SkipOffDays => _skipOffBox.IsChecked == true;
    public bool IsRange => _rangeRadio.IsChecked == true;

    /// <summary>单次复制允许的最大天数（L4：防止超长日期范围卡死 UI）</summary>
    private const int MaxCopyDays = 366;

    /// <summary>目标日期集合（L4：生成时也受最大天数约束，双重保护）</summary>
    public List<DateTime> ToDates
    {
        get
        {
            if (IsRange)
            {
                var start = _rangeStartPicker.SelectedDate ?? _defaultFrom;
                var end = _rangeEndPicker.SelectedDate ?? _defaultFrom;
                if (end < start) (start, end) = (end, start);
                var list = new List<DateTime>();
                var count = 0;
                for (var d = start; d <= end && count < MaxCopyDays; d = d.AddDays(1))
                {
                    list.Add(d);
                    count++;
                }
                return list;
            }
            var single = _toSinglePicker.SelectedDate ?? _defaultFrom;
            return new List<DateTime> { single };
        }
    }

    public CopyScheduleDialog(DateTime defaultFrom)
    {
        _defaultFrom = defaultFrom;
        Title = "复制排课";
        Width = 460;
        Height = 400;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        ResizeMode = ResizeMode.NoResize;
        ShowInTaskbar = false;
        WindowStyle = WindowStyle.None;
        AllowsTransparency = true;
        Background = Brushes.Transparent;
        // L2：不再全局置顶，避免遮挡其他窗口

        var grid = new Grid { Margin = new Thickness(24) };
        grid.RowDefinitions.Add(new RowDefinition());
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        // 标题
        grid.Children.Add(new TextBlock
        {
            Text = "复制排课", FontSize = 16, FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)),
            Margin = new Thickness(0, 0, 0, 12)
        });

        // 源日期
        _fromPicker.SelectedDate = _defaultFrom;
        grid.Children.Add(Row("源日期：", _fromPicker));
        Grid.SetRow(grid.Children[^1], 1);

        // 目标方式
        var modePanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 10, 0, 0) };
        _singleRadio.Checked += (_, _) => UpdateMode();
        _rangeRadio.Checked += (_, _) => UpdateMode();
        modePanel.Children.Add(_singleRadio);
        modePanel.Children.Add(_rangeRadio);
        grid.Children.Add(modePanel);
        Grid.SetRow(modePanel, 2);

        // 单日目标
        _toSinglePicker.SelectedDate = _defaultFrom;
        var singlePanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 10, 0, 0) };
        singlePanel.Children.Add(new TextBlock { Text = "目标日期：", FontSize = 13, Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)), VerticalAlignment = VerticalAlignment.Center });
        singlePanel.Children.Add(_toSinglePicker);
        grid.Children.Add(singlePanel);
        Grid.SetRow(singlePanel, 3);

        // 范围目标
        var rangePanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 10, 0, 0), Visibility = Visibility.Collapsed };
        _rangeStartPicker.SelectedDate = _defaultFrom;
        _rangeEndPicker.SelectedDate = _defaultFrom;
        rangePanel.Children.Add(new TextBlock { Text = "从：", FontSize = 13, Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)), VerticalAlignment = VerticalAlignment.Center });
        rangePanel.Children.Add(_rangeStartPicker);
        rangePanel.Children.Add(new TextBlock { Text = "  到：", FontSize = 13, Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)), VerticalAlignment = VerticalAlignment.Center });
        rangePanel.Children.Add(_rangeEndPicker);
        grid.Children.Add(rangePanel);
        Grid.SetRow(rangePanel, 4);
        _rangePanel = rangePanel;

        // 跳过不排课日
        grid.Children.Add(_skipOffBox);
        Grid.SetRow(_skipOffBox, 5);

        // 按钮
        var cancelBtn = NewBtn("取消", Color.FromRgb(0xf0, 0xf0, 0xf0), Color.FromRgb(0x55, 0x55, 0x55));
        var okBtn = NewBtn("复制", Color.FromRgb(0x42, 0x85, 0xf4), Color.FromRgb(0xff, 0xff, 0xff));
        cancelBtn.Click += (_, _) => { DialogResult = false; Close(); };
        okBtn.Click += (_, _) =>
        {
            // L4：超长日期范围在复制前拦截，避免生成海量日期导致 UI 卡死
            var start = _rangeStartPicker.SelectedDate ?? _defaultFrom;
            var end = _rangeEndPicker.SelectedDate ?? _defaultFrom;
            if (end < start) (start, end) = (end, start);
            if (IsRange && (end - start).Days + 1 > MaxCopyDays)
            {
                ModernDialog.Alert(
                    $"日期范围最多支持 {MaxCopyDays} 天（当前选择 {(end - start).Days + 1} 天）。\n请缩小范围后重试。",
                    "范围过大");
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
        Grid.SetRow(btnPanel, 6);

        Content = new Border
        {
            CornerRadius = new CornerRadius(12),
            Background = new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xe0, 0xe0, 0xe0)),
            BorderThickness = new Thickness(1),
            Child = grid
        };
    }

    private void UpdateMode()
    {
        if (_rangePanel != null)
            _rangePanel.Visibility = IsRange ? Visibility.Visible : Visibility.Collapsed;
    }

    private static UIElement Row(string label, UIElement input)
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 10, 0, 0) };
        panel.Children.Add(new TextBlock
        {
            Text = label, FontSize = 13,
            Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)),
            VerticalAlignment = VerticalAlignment.Center, Width = 70
        });
        panel.Children.Add(input);
        return panel;
    }

    private static Button NewBtn(string text, Color bg, Color fg)
    {
        var btn = new Button
        {
            Content = text, Width = 80, Height = 32, Cursor = Cursors.Hand,
            FontSize = 13, Foreground = new SolidColorBrush(fg),
            Background = new SolidColorBrush(bg),
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
