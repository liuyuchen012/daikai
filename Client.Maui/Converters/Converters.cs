using System.Globalization;

namespace CheckIn.Client.Maui.Converters;

/// <summary>
/// 排名颜色转换器：根据排名序号返回对应颜色
/// 第1名金色、第2名银色、第3名铜色、其余黑色
/// </summary>
public class RankColorConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is int rank)
        {
            return rank switch
            {
                1 => Color.FromArgb("#FFD700"), // 金
                2 => Color.FromArgb("#C0C0C0"), // 银
                3 => Color.FromArgb("#CD7F32"), // 铜
                _ => Color.FromArgb("#333333")  // 默认黑色
            };
        }
        return Color.FromArgb("#333333");
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// 布尔反向转换器：true 时隐藏，false 时显示
/// </summary>
public class InverseBoolConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is true ? false : true;

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is true ? false : true;
}

/// <summary>
/// 布尔值到文本转换器：用于已打卡/未打卡状态显示
/// </summary>
public class BoolToCheckInTextConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is true ? "已打卡" : "未打卡";

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// 布尔值到背景色转换器：已打卡绿色，未打卡默认色
/// </summary>
public class BoolToBackgroundConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is true ? Color.FromArgb("#E8F5E9") : Color.FromArgb("#F5F5F5");

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// 颜色字符串到 Color 对象转换器（用于 XAML 绑定）
/// </summary>
public class StringToColorConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is string hex && hex.StartsWith("#"))
        {
            return Color.FromArgb(hex);
        }
        return Color.FromArgb("#333333");
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotImplementedException();
}
