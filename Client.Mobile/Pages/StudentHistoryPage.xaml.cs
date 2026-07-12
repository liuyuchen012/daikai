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
        lblRole.Text = _auth.IsAdmin ? "管理员" : "学生";
        return Task.CompletedTask;
    }

    private async Task LoadHistoryAsync()
    {
        historyStack.Children.Clear();
        try
        {
            var result = await _api.GetAsync("/api/mobile/students/history");
            if (ApiService.GetError(result) != null) return;

            if (result.TryGetProperty("total_checkins", out var total))
                lblTotal.Text = total.GetInt32().ToString();

            if (result.TryGetProperty("history", out var history) && history.ValueKind == JsonValueKind.Array)
            {
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
                    return;
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
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"History load error: {ex.Message}");
        }
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
