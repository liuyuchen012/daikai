using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace CheckIn.Client.Maui.ViewModels;

/// <summary>
/// ViewModel 基类，提供 INotifyPropertyChanged 实现
/// </summary>
public class BaseViewModel : INotifyPropertyChanged
{
    private bool _isBusy;
    public bool IsBusy
    {
        get => _isBusy;
        set { SetProperty(ref _isBusy, value); }
    }

    private string _title = "";
    public string Title
    {
        get => _title;
        set { SetProperty(ref _title, value); }
    }

    protected bool SetProperty<T>(ref T backingStore, T value,
        [CallerMemberName] string propertyName = "")
    {
        if (EqualityComparer<T>.Default.Equals(backingStore, value))
            return false;

        backingStore = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = "") =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}
