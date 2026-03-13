using System.Windows;

namespace PowerInsighter.Helpers;

/// <summary>
/// A proxy class that allows binding to elements that are not part of the visual tree,
/// such as DataGridColumn which doesn't inherit DataContext.
/// </summary>
public class BindingProxy : Freezable
{
    protected override Freezable CreateInstanceCore()
    {
        return new BindingProxy();
    }

    public object Data
    {
        get => GetValue(DataProperty);
        set => SetValue(DataProperty, value);
    }

    public static readonly DependencyProperty DataProperty =
        DependencyProperty.Register(
            nameof(Data),
            typeof(object),
            typeof(BindingProxy),
            new UIPropertyMetadata(null));
}
