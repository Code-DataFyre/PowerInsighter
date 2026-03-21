using System.Globalization;
using System.Windows.Data;
using System.Windows;

namespace PowerInsighter.Converters;

public class SeverityToIconConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is string severity)
        {
            return severity.ToLower() switch
            {
                "error" => "?", // Red X
                "warning" => "??", // Yellow triangle
                "info" or "information" => "??", // Info circle
                _ => "??"
            };
        }
        return "??";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
