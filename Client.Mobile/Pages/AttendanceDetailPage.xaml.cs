using System.Text.Json;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.Pages;

[QueryProperty(nameof(ShortCode), "shortCode")]
[QueryProperty(nameof(Subject), "subject")]
public partial class AttendanceDetailPage : ContentPage
{
    private readonly ApiService _api;
    private string _shortCode = "";
    private string _subject = "";

    public string ShortCode
    {
        get => _shortCode;
        set { _shortCode = value; OnPropertyChanged(); }
    }

    public string Subject
    {
        get => _subject;
        set { _subject = value; OnPropertyChanged(); }
    }

    public AttendanceDetailPage(ApiService api)
    {
        InitializeComponent();
        _api = api;
    }

    protected override async void OnAppearing()
    {
        base.OnAppearing();
        lblSubject.Text = _subject;
        lblShortCode.Text = $"签到码: {_shortCode}";
        await LoadAttendanceAsync();
    }

    private async Task LoadAttendanceAsync()
    {
        try
        {
            var result = await _api.GetAsync($"/api/mobile/attendance?task_id=signin_{_shortCode}");
            if (ApiService.GetError(result) != null) return;

            if (result.TryGetProperty("tasks", out var tasks) && tasks.ValueKind == JsonValueKind.Array)
            {
                foreach (var task in tasks.EnumerateArray())
                {
                    var total = GetInt(task, "total_students");
                    var punched = GetInt(task, "punched_count");
                    var unpunched = GetInt(task, "unpunched_count");

                    lblTotal.Text = total.ToString();
                    lblCheckedIn.Text = punched.ToString();
                    lblNotCheckedIn.Text = unpunched.ToString();

                    if (task.TryGetProperty("students", out var students) && students.ValueKind == JsonValueKind.Array)
                    {
                        studentListStack.Children.Clear();
                        foreach (var s in students.EnumerateArray())
                        {
                            var name = ApiService.GetString(s, "name") ?? "";
                            var checkedIn = s.TryGetProperty("checked_in", out var ci) && ci.GetBoolean();
                            var time = ApiService.GetString(s, "first_time") ?? "";

                            var frame = new Border
                            {
                                BackgroundColor = checkedIn ? Color.FromArgb("#e6f4ea") : Colors.White,
                                StrokeShape = new Microsoft.Maui.Controls.Shapes.RoundRectangle { CornerRadius = 8 },
                                Padding = new Thickness(12),
                                Stroke = new SolidColorBrush(checkedIn ? Color.FromArgb("#34a853") : Color.FromArgb("#e8e8e8")),
                                StrokeThickness = 0.5
                            };

                            var grid = new Grid
                            {
                                ColumnDefinitions =
                                {
                                    new ColumnDefinition(GridLength.Star),
                                    new ColumnDefinition(GridLength.Auto)
                                }
                            };

                            grid.Add(new Label
                            {
                                Text = name,
                                FontSize = 14,
                                FontAttributes = FontAttributes.Bold,
                                TextColor = Color.FromArgb("#333333"),
                                VerticalOptions = LayoutOptions.Center
                            }, 0);

                            var statusText = checkedIn ? "✅ 已签到" : "❌ 未签到";
                            grid.Add(new Label
                            {
                                Text = checkedIn ? $"{statusText}  {time}" : statusText,
                                FontSize = 11,
                                TextColor = checkedIn ? Color.FromArgb("#34a853") : Color.FromArgb("#ea4335"),
                                VerticalOptions = LayoutOptions.Center
                            }, 1);

                            frame.Content = grid;
                            studentListStack.Children.Add(frame);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Attendance load error: {ex.Message}");
        }
    }

    private static int GetInt(JsonElement json, string key) =>
        json.TryGetProperty(key, out var val) && val.TryGetInt32(out var i) ? i : 0;
}
