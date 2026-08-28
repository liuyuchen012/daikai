using System.Text.Json;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.Pages;

public partial class StudentHistoryPage : ContentPage
{
    private readonly ApiService _api;
    private readonly AuthService _auth;

    public StudentHistoryPage(ApiService api, AuthService auth)
    {
        InitializeComponent();
        _api = api;
        _auth = auth;
    }

    protected override async void OnAppearing()
    {
        base.OnAppearing();
        await LoadUserInfo();
        await LoadHistoryAsync();
    }

    private Task LoadUserInfo()
    {
        lblUsername.Text = _auth.CurrentDisplayName ?? _auth.CurrentUsername ?? "用户";
        lblRole.Text = _auth.IsAdmin ? "管理员" :
                       _auth.IsTeacher ? "普通教师" :
                       _auth.IsParent ? "家长" : "学生";

        // 管理员/教师：更新统计标签
        if (_auth.IsAdminOrTeacher)
        {
            Title = "签到任务";
            lblStatsLabel.Text = "总签到次数: ";
        }
        return Task.CompletedTask;
    }

    private async Task LoadHistoryAsync()
    {
        historyStack.Children.Clear();
        try
        {
            var result = await _api.GetAsync("/api/mobile/students/history");
            if (ApiService.GetError(result) != null) return;

            if (_auth.IsAdminOrTeacher)
                await LoadAdminTeacherHistory(result);
            else
                await LoadStudentHistory(result);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"History load error: {ex.Message}");
        }
    }

    /// <summary>
    /// 管理员/教师视图：按任务展示签到汇总
    /// </summary>
    private Task LoadAdminTeacherHistory(JsonElement result)
    {
        if (result.TryGetProperty("total_checkins", out var total))
            lblTotal.Text = total.GetInt32().ToString();

        if (result.TryGetProperty("total_tasks", out var taskCount))
            lblStatsExtra.Text = $"  |  任务数: {taskCount.GetInt32()}";

        if (!result.TryGetProperty("history", out var history) || history.ValueKind != JsonValueKind.Array)
            return Task.CompletedTask;

        var items = history.EnumerateArray().ToList();
        if (!items.Any())
        {
            historyStack.Children.Add(new Label
            {
                Text = "暂无签到任务",
                TextColor = Color.FromArgb("#888888"),
                HorizontalOptions = LayoutOptions.Center,
                Margin = new Thickness(0, 20)
            });
            return Task.CompletedTask;
        }

        foreach (var item in items)
        {
            var subject = ApiService.GetString(item, "subject") ?? "未知";
            var classroom = ApiService.GetString(item, "classroom") ?? "";
            var status = ApiService.GetString(item, "status") ?? "active";
            var signedCount = 0;
            var studentCount = 0;
            if (item.TryGetProperty("signed_count", out var sc)) signedCount = sc.GetInt32();
            if (item.TryGetProperty("student_count", out var stc)) studentCount = stc.GetInt32();

            var frame = new Border
            {
                BackgroundColor = Colors.White,
                StrokeShape = new Microsoft.Maui.Controls.Shapes.RoundRectangle { CornerRadius = 10 },
                Padding = new Thickness(14),
                Stroke = new SolidColorBrush(Color.FromArgb("#e8e8e8")),
                StrokeThickness = 0.5
            };

            var grid = new Grid
            {
                ColumnDefinitions =
                {
                    new ColumnDefinition(GridLength.Star),
                    new ColumnDefinition(GridLength.Auto)
                },
                RowDefinitions =
                {
                    new RowDefinition(GridLength.Auto),
                    new RowDefinition(GridLength.Auto),
                    new RowDefinition(GridLength.Auto)
                },
                RowSpacing = 4
            };

            // 第一行：科目 + 状态
            grid.Add(new Label
            {
                Text = subject,
                FontSize = 15,
                FontAttributes = FontAttributes.Bold,
                TextColor = Color.FromArgb("#333333")
            }, 0, 0);

            var statusBadge = new Border
            {
                BackgroundColor = status == "active" ? Color.FromArgb("#e8f0fe") : Color.FromArgb("#f0f0f0"),
                StrokeShape = new Microsoft.Maui.Controls.Shapes.RoundRectangle { CornerRadius = 4 },
                Padding = new Thickness(6, 2),
                Stroke = new SolidColorBrush(Colors.Transparent),
                StrokeThickness = 0,
                Content = new Label
                {
                    Text = status == "active" ? "进行中" : "已结束",
                    FontSize = 10,
                    TextColor = status == "active" ? Color.FromArgb("#4285f4") : Color.FromArgb("#888888")
                }
            };
            grid.Add(statusBadge, 1, 0);

            // 第二行：进度
            grid.Add(new Label
            {
                Text = $"签到进度: {signedCount}/{studentCount}",
                FontSize = 12,
                TextColor = Color.FromArgb("#4285f4")
            }, 0, 1);
            Grid.SetColumnSpan((BindableObject)grid.Children.Last(), 2);

            // 第三行：教室信息 + 签到人员名单
            var recordsText = "";
            if (item.TryGetProperty("records", out var records) && records.ValueKind == JsonValueKind.Array)
            {
                var recordList = records.EnumerateArray()
                    .Select(r => ApiService.GetString(r, "name") ?? "")
                    .Where(n => !string.IsNullOrEmpty(n))
                    .ToList();
                if (recordList.Count > 0)
                    recordsText = $"已签到: {string.Join("、", recordList.Take(10))}" +
                                  (recordList.Count > 10 ? $" 等{recordList.Count}人" : "");
                else
                    recordsText = "暂无签到";
            }

            var infoLabel = new Label
            {
                Text = string.IsNullOrEmpty(classroom) ? recordsText : $"{classroom}  |  {recordsText}",
                FontSize = 11,
                TextColor = Color.FromArgb("#888888"),
                LineBreakMode = LineBreakMode.TailTruncation,
                MaxLines = 2
            };
            grid.Add(infoLabel, 0, 2);
            Grid.SetColumnSpan((BindableObject)infoLabel, 2);

            frame.Content = grid;
            historyStack.Children.Add(frame);
        }

        return Task.CompletedTask;
    }

    /// <summary>
    /// 学生/家长视图：按时间展示签到记录
    /// </summary>
    private Task LoadStudentHistory(JsonElement result)
    {
        if (result.TryGetProperty("total_checkins", out var total))
            lblTotal.Text = total.GetInt32().ToString();

        if (!result.TryGetProperty("history", out var history) || history.ValueKind != JsonValueKind.Array)
            return Task.CompletedTask;

        var items = history.EnumerateArray().ToList();
        if (!items.Any())
        {
            historyStack.Children.Add(new Label
            {
                Text = "暂无签到记录",
                TextColor = Color.FromArgb("#888888"),
                HorizontalOptions = LayoutOptions.Center,
                Margin = new Thickness(0, 20)
            });
            return Task.CompletedTask;
        }

        foreach (var item in items)
        {
            var subject = ApiService.GetString(item, "subject") ?? "未知";
            var classroom = ApiService.GetString(item, "classroom") ?? "";
            var time = ApiService.GetString(item, "time") ?? "";
            var status = ApiService.GetString(item, "status") ?? "active";

            var frame = new Border
            {
                BackgroundColor = Colors.White,
                StrokeShape = new Microsoft.Maui.Controls.Shapes.RoundRectangle { CornerRadius = 10 },
                Padding = new Thickness(14),
                Stroke = new SolidColorBrush(Color.FromArgb("#e8e8e8")),
                StrokeThickness = 0.5
            };

            var grid = new Grid
            {
                ColumnDefinitions =
                {
                    new ColumnDefinition(GridLength.Star),
                    new ColumnDefinition(GridLength.Auto)
                },
                RowDefinitions =
                {
                    new RowDefinition(GridLength.Auto),
                    new RowDefinition(GridLength.Auto)
                },
                RowSpacing = 4
            };

            grid.Add(new Label
            {
                Text = subject,
                FontSize = 15,
                FontAttributes = FontAttributes.Bold,
                TextColor = Color.FromArgb("#333333")
            }, 0, 0);

            var statusBadge = new Border
            {
                BackgroundColor = status == "active" ? Color.FromArgb("#e8f0fe") : Color.FromArgb("#f0f0f0"),
                StrokeShape = new Microsoft.Maui.Controls.Shapes.RoundRectangle { CornerRadius = 4 },
                Padding = new Thickness(6, 2),
                Stroke = new SolidColorBrush(Colors.Transparent),
                StrokeThickness = 0,
                Content = new Label
                {
                    Text = status == "active" ? "进行中" : "已结束",
                    FontSize = 10,
                    TextColor = status == "active" ? Color.FromArgb("#4285f4") : Color.FromArgb("#888888")
                }
            };
            grid.Add(statusBadge, 1, 0);

            var classroomLabel = new Label
            {
                Text = string.IsNullOrEmpty(classroom) ? time : $"{classroom}  |  {time}",
                FontSize = 12,
                TextColor = Color.FromArgb("#888888")
            };
            grid.Add(classroomLabel, 0, 1);
            Grid.SetColumnSpan((BindableObject)classroomLabel, 2);

            frame.Content = grid;
            historyStack.Children.Add(frame);
        }

        return Task.CompletedTask;
    }

    private async void OnLogoutClicked(object? sender, EventArgs e)
    {
        var confirmed = await DisplayAlertAsync("退出登录", "确定要退出登录吗？", "退出", "取消");
        if (!confirmed) return;

        if (Shell.Current is AppShell shell)
        {
            await shell.LogoutAsync();
        }
    }
}
