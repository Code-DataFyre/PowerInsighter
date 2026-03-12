using System.Windows;
using PowerInsighter.ViewModels;

namespace PowerInsighter;

public partial class MainWindow : Window
{
    public MainWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
    }
}