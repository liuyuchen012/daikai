using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace CheckIn.Client;

public class RankColorConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is int rank)
        {
            return rank switch
            {
                1 => new SolidColorBrush(Color.FromRgb(0xFF, 0xD7, 0x00)), // 金
                2 => new SolidColorBrush(Color.FromRgb(0xC0, 0xC0, 0xC0)), // 银
                3 => new SolidColorBrush(Color.FromRgb(0xCD, 0x7F, 0x32)), // 铜
                _ => new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))
            };
        }
        return new SolidColorBrush(Colors.Black);
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotImplementedException();
}
