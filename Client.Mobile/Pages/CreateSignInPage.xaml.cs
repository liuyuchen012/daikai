using CheckIn.Client.Mobile.ViewModels;

namespace CheckIn.Client.Mobile.Pages;

public partial class CreateSignInPage : ContentPage
{
    private MainViewModel _mainVm = null!;

    public CreateSignInPage()
    {
        InitializeComponent();
    }

    protected override void OnAppearing()
    {
        base.OnAppearing();
        _mainVm = IPlatformApplication.Current!.Services.GetRequiredService<MainViewModel>();
    }

    private async void OnCreateClicked(object? sender, EventArgs e)
    {
        // Validate inputs
        var password = PasswordEntry.Text?.Trim();
        var confirmPassword = ConfirmPasswordEntry.Text?.Trim();
        var classroom = ClassroomEntry.Text?.Trim();
        var subject = SubjectEntry.Text?.Trim();
        var studentsText = StudentsEditor.Text?.Trim();

        if (string.IsNullOrWhiteSpace(password))
        {
            await DisplayAlertAsync("Error", "Please enter a sign-in password", "OK");
            return;
        }

        if (password != confirmPassword)
        {
            await DisplayAlertAsync("Error", "Passwords do not match", "OK");
            return;
        }

        if (string.IsNullOrWhiteSpace(classroom))
        {
            await DisplayAlertAsync("Error", "Please enter a classroom name", "OK");
            return;
        }

        if (string.IsNullOrWhiteSpace(subject))
        {
            await DisplayAlertAsync("Error", "Please enter a subject name", "OK");
            return;
        }

        // Parse students
        var students = new List<string>();
        if (!string.IsNullOrWhiteSpace(studentsText))
        {
            students = studentsText
                .Split('\n', StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();
        }

        if (students.Count == 0)
        {
            bool proceed = await DisplayAlertAsync("Warning",
                "No students entered. Continue anyway? (Students will not be verified)",
                "Continue", "Cancel");
            if (!proceed) return;
        }

        // Show loading
        SubmitButton.IsEnabled = false;
        LoadingIndicator.IsVisible = true;
        LoadingIndicator.IsRunning = true;
        StatusLabel.IsVisible = true;
        StatusLabel.Text = "Creating sign-in task...";

        try
        {
            var result = await _mainVm.CreateRemoteSignInAsync(password, classroom, subject, students);
            if (result == null)
            {
                await DisplayAlertAsync("Error", "Failed to create sign-in. Check server connection.", "OK");
                return;
            }

            var (shortCode, taskId) = result.Value;

            // Create local task tab
            var taskName = $"{classroom} {subject} Sign-In";
            _mainVm.AddTab(taskName, subject, isSignIn: true, signInTaskId: taskId);

            // Show success with link
            var signUrl = $"http://{_mainVm.Config.ServerIp}:{_mainVm.Config.ServerPort}/s/{shortCode}";
            await DisplayAlertAsync("Success",
                $"Sign-in task created!\n\nLink: {signUrl}\n\nStudents can use this link to sign in.",
                "OK");

            await Shell.Current.GoToAsync("..");
        }
        catch (Exception ex)
        {
            await DisplayAlertAsync("Error", $"Failed: {ex.Message}", "OK");
        }
        finally
        {
            SubmitButton.IsEnabled = true;
            LoadingIndicator.IsVisible = false;
            LoadingIndicator.IsRunning = false;
            StatusLabel.IsVisible = false;
        }
    }
}
