using CheckIn.Client.Mobile.Models;
using CheckIn.Client.Mobile.ViewModels;

namespace CheckIn.Client.Mobile.Pages;

[QueryProperty(nameof(TabId), "tabId")]
public partial class StudentGridPage : ContentPage
{
    private TaskTabViewModel? _tabVm;
    public string? TabId { get; set; }

    public StudentGridPage()
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

        BuildStudentGrid();
    }

    private void BuildStudentGrid()
    {
        if (_tabVm == null) return;

        var children = new List<View>();
        foreach (var student in _tabVm.Students)
        {
            children.Add(CreateStudentButton(student));
        }

        // Replace FlexLayout children
        StudentGrid.Children.Clear();
        foreach (var child in children)
            StudentGrid.Children.Add(child);
    }

    private View CreateStudentButton(StudentModel student)
    {
        var frame = new Frame
        {
            CornerRadius = 10,
            Padding = new Thickness(12, 10),
            HasShadow = true,
            Margin = new Thickness(4),
            MinimumWidthRequest = 100,
            BackgroundColor = student.IsCheckedIn
                ? Color.FromArgb("#4285f4")
                : Color.FromArgb("#e8e8e8"),
            BorderColor = student.IsCheckedIn
                ? Color.FromArgb("#4285f4")
                : Color.FromArgb("#d0d0d0"),
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

        // Tap: check in
        var tapGesture = new TapGestureRecognizer();
        tapGesture.Tapped += (s, e) =>
        {
            if (!student.IsCheckedIn && _tabVm != null)
            {
                _tabVm.CheckInCommand.Execute(student);
                UpdateButtonAppearance(frame, student);
            }
        };
        frame.GestureRecognizers.Add(tapGesture);

        // Long press: cancel
        tapGesture = new TapGestureRecognizer();
        tapGesture.Tapped += async (s, e) =>
        {
            if (student.IsCheckedIn && _tabVm != null)
            {
                bool confirm = await DisplayAlert(
                    "Cancel Check-in",
                    $"Cancel {student.Name}'s check-in?",
                    "Yes", "No");
                if (confirm)
                {
                    _tabVm.CancelCheckInCommand.Execute(student);
                    UpdateButtonAppearance(frame, student);
                }
            }
        };
        frame.GestureRecognizers.Add(tapGesture);

        return frame;
    }

    private void UpdateButtonAppearance(Frame frame, StudentModel student)
    {
        frame.BackgroundColor = student.IsCheckedIn
            ? Color.FromArgb("#4285f4")
            : Color.FromArgb("#e8e8e8");
        frame.BorderColor = student.IsCheckedIn
            ? Color.FromArgb("#4285f4")
            : Color.FromArgb("#d0d0d0");
        if (frame.Content is Label label)
        {
            label.Text = student.DisplayText;
            label.TextColor = student.IsCheckedIn ? Colors.White : Color.FromArgb("#333333");
        }
    }
}
