using System.Globalization;

namespace CheckIn.Client.Mobile.Converters;

/// <summary>
/// Converts a boolean (IsCheckedIn) to a Color:
/// true -> Primary blue, false -> Gray (not checked in).
/// </summary>
public class BoolToColorConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool isCheckedIn)
        {
            return isCheckedIn ? Color.FromArgb("#4285f4") : Color.FromArgb("#e8e8e8");
        }
        return Color.FromArgb("#e8e8e8");
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// Converts a rank number to a display color:
/// 1 -> Gold, 2 -> Silver, 3 -> Bronze, others -> Dark gray.
/// </summary>
public class RankColorConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is int rank)
        {
            return rank switch
            {
                1 => Color.FromArgb("#FFD700"),
                2 => Color.FromArgb("#C0C0C0"),
                3 => Color.FromArgb("#CD7F32"),
                _ => Color.FromArgb("#333333")
            };
        }
        return Color.FromArgb("#333333");
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// Inverts a boolean value. Used for visibility toggles.
/// </summary>
public class InverseBoolConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is bool b && !b;

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is bool b && !b;
}

/// <summary>
/// Returns true if the string is not null or empty.
/// </summary>
public class StringNotEmptyConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is string s && !string.IsNullOrEmpty(s);

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotImplementedException();
}
