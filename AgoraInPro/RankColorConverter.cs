using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace CheckIn.Client;

/// <summary>
/// 排名颜色转换器：根据排名序号返回对应颜色
/// 第1名金色、第2名银色、第3名铜色、其余黑色
/// </summary>
public class RankColorConverter : IValueConverter
{
    /// <summary>
    /// 根据排名值转换为对应颜色的 SolidColorBrush
    /// </summary>
    /// <param name="value">排名序号（int）</param>
    /// <returns>金/银/铜/黑色画刷</returns>
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is int rank)
        {
            return rank switch
            {
                1 => new SolidColorBrush(Color.FromRgb(0xFF, 0xD7, 0x00)), // 金
                2 => new SolidColorBrush(Color.FromRgb(0xC0, 0xC0, 0xC0)), // 银
                3 => new SolidColorBrush(Color.FromRgb(0xCD, 0x7F, 0x32)), // 铜
                _ => new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))  // 默认黑色
            };
        }
        return new SolidColorBrush(Colors.Black);
    }

    /// <summary>
    /// 反向转换不支持，直接抛出异常
    /// </summary>
    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// 布尔反向可见性转换器：true 时隐藏，false 时显示
/// 用于控制"无标签时的占位提示"的 Visibility
/// </summary>
public class InverseBoolToVisibilityConverter : IValueConverter
{
    /// <summary>
    /// 将布尔值反转为 Visibility：true -> Collapsed, false -> Visible
    /// </summary>
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is true ? Visibility.Collapsed : Visibility.Visible;

    /// <summary>
    /// 反向转换不支持，直接抛出异常
    /// </summary>
    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotImplementedException();
}
