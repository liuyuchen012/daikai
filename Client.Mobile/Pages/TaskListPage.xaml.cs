using CheckIn.Client.Mobile.ViewModels;

namespace CheckIn.Client.Mobile.Pages;

public partial class TaskListPage : ContentPage
{
    private MainViewModel _vm = null!;

    public TaskListPage()
    {
        InitializeComponent();
    }

    protected override void OnAppearing()
    {
        base.OnAppearing();

        if (_vm == null)
        {
            _vm = IPlatformApplication.Current!.Services.GetRequiredService<MainViewModel>();
            BindingContext = _vm;
        }

        _vm.RefreshAllTasks();

        // Update responsive layout
        UpdateLayoutForDevice();
    }

    private void UpdateLayoutForDevice()
    {
        var isTablet = Helpers.ResponsiveHelper.Instance.IsTablet;
        if (isTablet)
        {
            // Tablet: use 2-column grid layout
            var gridLayout = new GridItemsLayout(2, ItemsLayoutOrientation.Vertical);
            PhoneTaskList.ItemsLayout = gridLayout;
        }
        else
        {
            // Phone: vertical list
            PhoneTaskList.ItemsLayout = new LinearItemsLayout(ItemsLayoutOrientation.Vertical);
        }
    }

    private async void OnTaskSelected(object? sender, SelectionChangedEventArgs e)
    {
        if (e.CurrentSelection.FirstOrDefault() is not TaskTabViewModel tab) return;

        // Clear selection
        if (sender is CollectionView cv)
            cv.SelectedItem = null;

        _vm.SwitchToTab(tab);

        // Navigate to task detail
        await Shell.Current.GoToAsync($"taskdetail?tabId={tab.TabId}");
    }

    private async void OnAddTaskClicked(object? sender, EventArgs e)
    {
        string name = await DisplayPromptAsync("New Task", "Task name:", "Create", "Cancel", "Task Name");
        if (string.IsNullOrWhiteSpace(name)) return;

        string km = await DisplayPromptAsync("Subject", "Subject name:", "OK", "Cancel", initialValue: "Math");
        _vm.AddTab(name.Trim(), km?.Trim());
    }

    private async void OnRefreshing(object? sender, EventArgs e)
    {
        _vm.RefreshAllTasks();
        await Task.Delay(500);
        TaskRefreshView.IsRefreshing = false;
    }
}
