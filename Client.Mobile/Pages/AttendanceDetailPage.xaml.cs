using System.Text.Json;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.Pages;

[QueryProperty(nameof(ShortCode), "shortCode")]
[QueryProperty(nameof(Subject), "subject")]
[QueryProperty(nameof(MachineUuid), "machineUuid")]
[QueryProperty(nameof(DeviceName), "deviceName")]
public partial class AttendanceDetailPage : ContentPage
{
    private readonly ApiService _api;
    private string _shortCode = "";
    private string _subject = "";
    private string _machineUuid = "";
    private string _deviceName = "";

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

    public string MachineUuid
    {
        get => _machineUuid;
        set { _machineUuid = value; OnPropertyChanged(); }
    }

    public string DeviceName
    {
        get => _deviceName;
        set { _deviceName = value; OnPropertyChanged(); }
    }

    public AttendanceDetailPage(ApiService api)
    {
        InitializeComponent();
        _api = api;
    }

    protected override async void OnAppearing()
    {
        base.OnAppearing();
        if (!string.IsNullOrEmpty(_machineUuid))
        {
            // 设备模式：每个任务独立显示统计
            Title = string.IsNullOrEmpty(_deviceName) ? "设备考勤" : _deviceName;
            lblSubject.Text = _deviceName;
            lblShortCode.Text = $"设备: {_machineUuid}";
            summaryStats.IsVisible = false;
            await LoadDeviceAttendanceAsync();
        }
        else
        {
            // 单任务模式：显示汇总统计
            Title = "考勤详情";
            lblSubject.Text = _subject;
            lblShortCode.Text = $"签到码: {_shortCode}";
            summaryStats.IsVisible = true;
            await LoadTaskAttendanceAsync();
        }
    }

    private async Task LoadTaskAttendanceAsync()
    {
        try
        {
            var result = await _api.GetAsync($"/api/mobile/attendance?task_id=signin_{_shortCode}");
            if (ApiService.GetError(result) != null) return;

            if (result.TryGetProperty("tasks", out var tasks) && tasks.ValueKind == JsonValueKind.Array)
            {
                foreach (var task in tasks.EnumerateArray())
                {
                    RenderSingleTask(task);
                    break; // 单任务模式只取第一个
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Attendance load error: {ex.Message}");
        }
    }

    private async Task LoadDeviceAttendanceAsync()
    {
        try
        {
            var result = await _api.GetAsync($"/api/mobile/attendance?machine_uuid={Uri.EscapeDataString(_machineUuid)}");
            if (ApiService.GetError(result) != null) return;

            if (result.TryGetProperty("tasks", out var tasks) && tasks.ValueKind == JsonValueKind.Array)
            {
                studentListStack.Children.Clear();
                foreach (var task in tasks.EnumerateArray())
                {
                    var taskId = ApiService.GetString(task, "task_id") ?? "";
                    var total = GetInt(task, "total_students");
                    var punched = GetInt(task, "punched_count");
                    var unpunched = GetInt(task, "unpunched_count");

                    // 任务标题
                    var taskSubject = taskId.StartsWith("signin_") ? taskId[7..] : taskId;

                    // 任务卡片容器
                    var taskCard = new Border
                    {
                        BackgroundColor = Colors.White,
                        StrokeShape = new Microsoft.Maui.Controls.Shapes.RoundRectangle { CornerRadius = 12 },
                        Padding = new Thickness(14),
                        Stroke = new SolidColorBrush(Color.FromArgb("#e8e8e8")),
                        StrokeThickness = 1,
                        Margin = new Thickness(0, 6, 0, 0)
                    };

                    var taskStack = new VerticalStackLayout { Spacing = 10 };

                    // 任务标题行
                    taskStack.Children.Add(new Label
                    {
                        Text = taskSubject,
                        FontSize = 15,
                        FontAttributes = FontAttributes.Bold,
                        TextColor = Color.FromArgb("#333333")
                    });

                    // 任务独立统计（三列）
                    var statsGrid = new Grid
                    {
                        ColumnDefinitions =
                        {
                            new ColumnDefinition(GridLength.Star),
                            new ColumnDefinition(GridLength.Star),
                            new ColumnDefinition(GridLength.Star)
                        },
                        ColumnSpacing = 6
                    };

                    statsGrid.Add(BuildMiniStatCard("总人数", total.ToString(), "#333333", "#f0f0f0"), 0);
                    statsGrid.Add(BuildMiniStatCard("已签到", punched.ToString(), "#34a853", "#e6f4ea"), 1);
                    statsGrid.Add(BuildMiniStatCard("未签到", unpunched.ToString(), "#ea4335", "#fce8e6"), 2);

                    taskStack.Children.Add(statsGrid);

                    // 分隔线
                    taskStack.Children.Add(new BoxView
                    {
                        HeightRequest = 1,
                        Color = Color.FromArgb("#eeeeee"),
                        Margin = new Thickness(0, 2)
                    });

                    // 学生列表
                    if (task.TryGetProperty("students", out var students) && students.ValueKind == JsonValueKind.Array)
                    {
                        var hasAny = false;
                        foreach (var s in students.EnumerateArray())
                        {
                            hasAny = true;
                            taskStack.Children.Add(BuildStudentRow(s));
                        }
                        if (!hasAny)
                        {
                            taskStack.Children.Add(new Label
                            {
                                Text = "暂无学生",
                                FontSize = 12,
                                TextColor = Color.FromArgb("#aaaaaa"),
                                HorizontalOptions = LayoutOptions.Center
                            });
                        }
                    }

                    taskCard.Content = taskStack;
                    studentListStack.Children.Add(taskCard);
                }
            }
            else
            {
                studentListStack.Children.Clear();
                studentListStack.Children.Add(new Label
                {
                    Text = "该设备暂无考勤记录",
                    FontSize = 14,
                    TextColor = Color.FromArgb("#888888"),
                    HorizontalOptions = LayoutOptions.Center,
                    Margin = new Thickness(0, 40)
                });
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Device attendance load error: {ex.Message}");
        }
    }

    private static Border BuildMiniStatCard(string label, string value, string valueColor, string bgColor)
    {
        var border = new Border
        {
            BackgroundColor = Color.FromArgb(bgColor),
            StrokeShape = new Microsoft.Maui.Controls.Shapes.RoundRectangle { CornerRadius = 6 },
            Padding = new Thickness(6, 8),
            Stroke = new SolidColorBrush(Colors.Transparent),
            StrokeThickness = 0
        };

        var stack = new VerticalStackLayout { Spacing = 2, HorizontalOptions = LayoutOptions.Center };
        stack.Children.Add(new Label
        {
            Text = label,
            FontSize = 10,
            TextColor = Color.FromArgb("#888888"),
            HorizontalOptions = LayoutOptions.Center
        });
        stack.Children.Add(new Label
        {
            Text = value,
            FontSize = 16,
            FontAttributes = FontAttributes.Bold,
            TextColor = Color.FromArgb(valueColor),
            HorizontalOptions = LayoutOptions.Center
        });

        border.Content = stack;
        return border;
    }

    private void RenderSingleTask(JsonElement task)
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
                studentListStack.Children.Add(BuildStudentRow(s));
            }
        }
    }

    private static View BuildStudentRow(JsonElement s)
    {
        var name = ApiService.GetString(s, "name") ?? "";
        var checkedIn = s.TryGetProperty("checked_in", out var ci) && ci.GetBoolean();
        var time = ApiService.GetString(s, "first_time") ?? "";

        var border = new Border
        {
            BackgroundColor = checkedIn ? Color.FromArgb("#e6f4ea") : Colors.White,
            StrokeShape = new Microsoft.Maui.Controls.Shapes.RoundRectangle { CornerRadius = 6 },
            Padding = new Thickness(10, 8),
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
            FontSize = 13,
            FontAttributes = FontAttributes.Bold,
            TextColor = Color.FromArgb("#333333"),
            VerticalOptions = LayoutOptions.Center
        }, 0);

        var statusText = checkedIn ? "已签到" : "未签到";
        grid.Add(new Label
        {
            Text = checkedIn ? $"{statusText}  {time}" : statusText,
            FontSize = 11,
            TextColor = checkedIn ? Color.FromArgb("#34a853") : Color.FromArgb("#ea4335"),
            VerticalOptions = LayoutOptions.Center
        }, 1);

        border.Content = grid;
        return border;
    }

    private static int GetInt(JsonElement json, string key) =>
        json.TryGetProperty(key, out var val) && val.TryGetInt32(out var i) ? i : 0;
}
