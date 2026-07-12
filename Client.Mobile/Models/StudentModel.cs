using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace CheckIn.Client.Mobile.Models;

/// <summary>
/// Student data model with INotifyPropertyChanged for UI binding.
/// Tracks name, check-in count, first time, and history.
/// </summary>
public class StudentModel : INotifyPropertyChanged
{
    private string _name = "";
    private int _count;
    private string? _firstTime;
    private bool _isCheckedIn;
    private List<string> _history = new();

    public string Name
    {
        get => _name;
        set { _name = value; OnPropertyChanged(); }
    }

    public int Count
    {
        get => _count;
        set { _count = value; OnPropertyChanged(); OnPropertyChanged(nameof(DisplayText)); }
    }

    public string? FirstTime
    {
        get => _firstTime;
        set { _firstTime = value; IsCheckedIn = value != null; OnPropertyChanged(); OnPropertyChanged(nameof(DisplayText)); }
    }

    public bool IsCheckedIn
    {
        get => _isCheckedIn;
        set { _isCheckedIn = value; OnPropertyChanged(); OnPropertyChanged(nameof(ButtonBackground)); OnPropertyChanged(nameof(ButtonTextColor)); OnPropertyChanged(nameof(DisplayText)); }
    }

    public List<string> History
    {
        get => _history;
        set { _history = value; OnPropertyChanged(); }
    }

    /// <summary>
    /// Display text for the student button: name + check count.
    /// </summary>
    public string DisplayText => Count > 0 ? $"{Name}\n({Count})" : Name;

    /// <summary>
    /// Button background color based on check-in status.
    /// </summary>
    public Color ButtonBackground => IsCheckedIn ? Color.FromArgb("#4285f4") : Color.FromArgb("#e8e8e8");

    /// <summary>
    /// Button text color based on check-in status.
    /// </summary>
    public Color ButtonTextColor => IsCheckedIn ? Colors.White : Color.FromArgb("#333333");

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
