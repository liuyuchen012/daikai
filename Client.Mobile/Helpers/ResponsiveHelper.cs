using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace CheckIn.Client.Mobile.Helpers;

/// <summary>
/// Singleton helper for responsive layout decisions.
/// Monitors device changes and provides properties that XAML can bind to.
/// </summary>
public class ResponsiveHelper : INotifyPropertyChanged
{
    private static readonly Lazy<ResponsiveHelper> _instance = new(() => new ResponsiveHelper());
    public static ResponsiveHelper Instance => _instance.Value;

    private bool _isTablet;
    private double _screenWidthDp;

    /// <summary>
    /// Whether the current device is a tablet (width >= 600dp).
    /// </summary>
    public bool IsTablet
    {
        get => _isTablet;
        set { if (_isTablet != value) { _isTablet = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsPhone)); OnPropertyChanged(nameof(FlyoutBehavior)); OnPropertyChanged(nameof(ColumnsCount)); } }
    }

    /// <summary>
    /// Whether the current device is a phone (width &lt; 600dp).
    /// </summary>
    public bool IsPhone => !_isTablet;

    /// <summary>
    /// Shell flyout behavior: Locked for tablet, Disabled for phone.
    /// </summary>
    public FlyoutBehavior FlyoutBehavior => _isTablet ? FlyoutBehavior.Locked : FlyoutBehavior.Disabled;

    /// <summary>
    /// Number of columns for grid layouts: 3 for tablet, 2 for phone.
    /// </summary>
    public int ColumnsCount => _isTablet ? 3 : 2;

    /// <summary>
    /// Current screen width in dp.
    /// </summary>
    public double ScreenWidthDp
    {
        get => _screenWidthDp;
        set { if (_screenWidthDp != value) { _screenWidthDp = value; OnPropertyChanged(); } }
    }

    private ResponsiveHelper()
    {
        UpdateDeviceInfo();
        DeviceDisplay.MainDisplayInfoChanged += (s, e) => UpdateDeviceInfo();
    }

    private void UpdateDeviceInfo()
    {
        var info = DeviceDisplay.MainDisplayInfo;
        ScreenWidthDp = info.Width / info.Density;
        IsTablet = DeviceInfo.Idiom == DeviceIdiom.Tablet || ScreenWidthDp >= 600;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
