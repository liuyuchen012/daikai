namespace CheckIn.Client.Mobile.Helpers;

/// <summary>
/// Device utility helper for detecting platform and device characteristics.
/// </summary>
public static class DeviceHelper
{
    /// <summary>
    /// Gets whether the current device is a phone (screen width less than 600dp).
    /// </summary>
    public static bool IsPhone => !IsTablet;

    /// <summary>
    /// Gets whether the current device is a tablet (screen width >= 600dp or device idiom is Tablet).
    /// </summary>
    public static bool IsTablet =>
        DeviceInfo.Idiom == DeviceIdiom.Tablet ||
        (DeviceDisplay.MainDisplayInfo.Width / DeviceDisplay.MainDisplayInfo.Density) >= 600;

    /// <summary>
    /// Gets the current screen width in device-independent pixels (dp).
    /// </summary>
    public static double ScreenWidthDp =>
        DeviceDisplay.MainDisplayInfo.Width / DeviceDisplay.MainDisplayInfo.Density;

    /// <summary>
    /// Gets the current screen height in device-independent pixels (dp).
    /// </summary>
    public static double ScreenHeightDp =>
        DeviceDisplay.MainDisplayInfo.Height / DeviceDisplay.MainDisplayInfo.Density;

    /// <summary>
    /// Gets whether the current platform is Android.
    /// </summary>
    public static bool IsAndroid => DeviceInfo.Platform == DevicePlatform.Android;

    /// <summary>
    /// Gets whether the current platform is iOS.
    /// </summary>
    public static bool IsIOS => DeviceInfo.Platform == DevicePlatform.iOS;
}
