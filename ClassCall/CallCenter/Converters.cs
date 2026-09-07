using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;

namespace CallCenter;

/// <summary>在线状态 → 颜色（在线绿 / 离线灰）</summary>
public class OnlineToBrushConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is true
            ? new SolidColorBrush(Color.FromRgb(0x22, 0xC5, 0x5E))
            : new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF));

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
