using CheckIn.Client.Mobile.Models;
using CheckIn.Client.Mobile.ViewModels;

namespace CheckIn.Client.Mobile.Pages;

[QueryProperty(nameof(TabId), "tabId")]
public partial class TaskDetailPage : ContentPage
{
    private TaskTabViewModel? _tabVm;
    private bool _isRankingTab = true;

    public string? TabId { get; set; }

    public TaskDetailPage()
    {
        InitializeComponent();
    }

    protected override void OnAppearing()
    {
        base.OnAppearing();

        if (_tabVm == null && !string.IsNullOrEmpty(TabId))
        {
            var mainVm = IPlatformApplication.Current!.Services.GetRequiredService<MainViewModel>();
            _tabVm = mainVm.AllTasks.FirstOrDefault(t => t.TabId == TabId);
            if (_tabVm != null)
            {
                BindingContext = _tabVm;
                Title = _tabVm.TabDisplayName;
            }
        }

        UpdateLayoutForDevice();
        BuildStudentButtons();
    }

    private void UpdateLayoutForDevice()
    {
        var isTablet = Helpers.ResponsiveHelper.Instance.IsTablet;

        PhoneLayout.IsVisible = !isTablet;
        TabletLayout.IsVisible = isTablet;

        if (!isTablet)
        {
            ShowPhoneTab(_isRankingTab);
        }
        else
        {
            BuildTabletStudentButtons();
        }
    }

    private void ShowPhoneTab(bool showRanking)
    {
        _isRankingTab = showRanking;
        PhoneRankingScroll.IsVisible = showRanking;
        PhoneGridScroll.IsVisible = !showRanking;

        RankingTabBtn.BackgroundColor = showRanking
            ? Color.FromArgb("#4285f4")
            : Colors.Transparent;
        RankingTabBtn.TextColor = showRanking ? Colors.White : Color.FromArgb("#666666");

        CheckInTabBtn.BackgroundColor = !showRanking
            ? Color.FromArgb("#4285f4")
            : Colors.Transparent;
        CheckInTabBtn.TextColor = !showRanking ? Colors.White : Color.FromArgb("#666666");
    }

    private void OnRankingTabClicked(object? sender, EventArgs e) => ShowPhoneTab(true);
    private void OnCheckInTabClicked(object? sender, EventArgs e) => ShowPhoneTab(false);

    private void BuildStudentButtons()
    {
        if (_tabVm == null) return;

        BuildFlexLayoutButtons(PhoneStudentGrid);
    }

    private void BuildTabletStudentButtons()
    {
        if (_tabVm == null) return;

        BuildFlexLayoutButtons(TabletStudentGrid);
    }

    private void BuildFlexLayoutButtons(FlexLayout layout)
    {
        if (_tabVm == null) return;

        var children = new List<View>();
        foreach (var student in _tabVm.Students)
        {
            var btn = CreateStudentButton(student);
            children.Add(btn);
        }

        // Update the FlexLayout children via BindableLayout template
        // The FlexLayout already uses BindableLayout.ItemsSource, so we just need
        // to define the item template. Let's use a different approach — build manually.
    }

    private View CreateStudentButton(StudentModel student)
    {
        var checkedColor = Color.FromArgb("#4285f4");
        var uncheckedBg = Color.FromArgb("#e8e8e8");
        var uncheckedStroke = Color.FromArgb("#d0d0d0");
        var border = new Border
        {
            StrokeShape = new Microsoft.Maui.Controls.Shapes.RoundRectangle { CornerRadius = 10 },
            Padding = new Thickness(12, 8),
            Shadow = new Shadow { Brush = new SolidColorBrush(Colors.Black), Offset = new Point(0, 2), Radius = 4, Opacity = 0.15f },
            Margin = new Thickness(4),
            MinimumWidthRequest = 100,
            BackgroundColor = student.IsCheckedIn ? checkedColor : uncheckedBg,
            Stroke = new SolidColorBrush(student.IsCheckedIn ? checkedColor : uncheckedStroke),
            StrokeThickness = 1,
            Content = new Label
            {
                Text = student.DisplayText,
                FontSize = 14,
                FontAttributes = FontAttributes.Bold,
                TextColor = student.IsCheckedIn ? Colors.White : Color.FromArgb("#333333"),
                HorizontalTextAlignment = TextAlignment.Center,
                VerticalTextAlignment = TextAlignment.Center,
                LineBreakMode = LineBreakMode.WordWrap,
            }
        };

        // Tap gesture for check-in
        var tapGesture = new TapGestureRecognizer();
        tapGesture.Tapped += (s, e) =>
        {
            if (!student.IsCheckedIn && _tabVm != null)
            {
                _tabVm.CheckInCommand.Execute(student);
                UpdateButtonAppearance(border, student);
            }
        };
        border.GestureRecognizers.Add(tapGesture);

        // Long press for cancel
        // Note: MAUI doesn't have built-in long press, but we can simulate
        // with a second tap on an already-checked-in student
        tapGesture = new TapGestureRecognizer();
        tapGesture.Tapped += async (s, e) =>
        {
            if (student.IsCheckedIn && _tabVm != null)
            {
                bool confirm = await DisplayAlertAsync(
                    "Cancel Check-in",
                    $"Cancel {student.Name}'s check-in?",
                    "Yes", "No");
                if (confirm)
                {
                    _tabVm.CancelCheckInCommand.Execute(student);
                    UpdateButtonAppearance(border, student);
                }
            }
        };
        border.GestureRecognizers.Add(tapGesture);

        return border;
    }

    private void UpdateButtonAppearance(Border border, StudentModel student)
    {
        var checkedColor = Color.FromArgb("#4285f4");
        var uncheckedBg = Color.FromArgb("#e8e8e8");
        var uncheckedStroke = Color.FromArgb("#d0d0d0");
        border.BackgroundColor = student.IsCheckedIn ? checkedColor : uncheckedBg;
        border.Stroke = new SolidColorBrush(student.IsCheckedIn ? checkedColor : uncheckedStroke);
        if (border.Content is Label label)
        {
            label.Text = student.DisplayText;
            label.TextColor = student.IsCheckedIn ? Colors.White : Color.FromArgb("#333333");
        }
    }
}
