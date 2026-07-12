using CheckIn.Client.Models;
using CheckIn.Client.Maui.ViewModels;

namespace CheckIn.Client.Maui.Pages;

public partial class TaskDetailPage : ContentPage
{
    private readonly TaskTabViewModel _viewModel;

    public TaskDetailPage(TaskTabViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        BindingContext = _viewModel;
    }

    private void OnStudentTapped(object? sender, TappedEventArgs e)
    {
        if (sender is BindableObject b && b.BindingContext is StudentModel student)
        {
            if (student.IsCheckedIn)
                _viewModel.CancelCheckInCommand.Execute(student);
            else
                _viewModel.CheckInCommand.Execute(student);
        }
    }
}
