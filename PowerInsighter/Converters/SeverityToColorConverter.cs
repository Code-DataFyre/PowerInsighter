using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace PowerInsighter.Converters;

public class SeverityToColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is string severity)
        {
            return severity.ToLower() switch
            {
                "error" => new SolidColorBrush(Colors.Red),
                "warning" => new SolidColorBrush(Colors.Orange),
                "info" or "information" => new SolidColorBrush(Colors.DodgerBlue),
                _ => new SolidColorBrush(Colors.Gray)
            };
        }
        return new SolidColorBrush(Colors.Gray);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
