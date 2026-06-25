using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace CheckIn.Client.Models;

public class StudentModel : INotifyPropertyChanged
{
    private string _name = "";
    private int _count;
    private string? _firstTime;
    private bool _isCheckedIn;
    private List<string> _history = new();

    public string Name { get => _name; set { _name = value; OnPropertyChanged(); } }
    public int Count { get => _count; set { _count = value; OnPropertyChanged(); } }
    public string? FirstTime { get => _firstTime; set { _firstTime = value; IsCheckedIn = value != null; OnPropertyChanged(); } }
    public bool IsCheckedIn { get => _isCheckedIn; set { _isCheckedIn = value; OnPropertyChanged(); } }
    public List<string> History { get => _history; set { _history = value; OnPropertyChanged(); } }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
