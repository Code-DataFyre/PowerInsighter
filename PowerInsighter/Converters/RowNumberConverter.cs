using System.Globalization;
using System.Windows.Controls;
using System.Windows.Data;

namespace PowerInsighter.Converters;

/// <summary>
/// Multi-value converter that calculates the 1-based row number from the DataGrid items.
/// </summary>
public class RowNumberConverter : IMultiValueConverter
{
    public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
    {
        // values[0] = current item
        // values[1] = ItemsSource collection
        if (values.Length >= 2 && values[0] != null && values[1] is System.Collections.IList items)
        {
            int index = items.IndexOf(values[0]);
            if (index >= 0)
            {
                return (index + 1).ToString();
            }
        }
        return "0";
    }

    public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
