using System.Windows;
using System.Windows.Input;
using PowerInsighter.Models;
using PowerInsighter.Services;
using PowerInsighter.ViewModels;

namespace PowerInsighter.Views;

public partial class ConnectionDialog : Window
{
    private readonly ConnectionDialogViewModel _viewModel;

    public PowerBIInstance? SelectedInstance => _viewModel.Result;

    public ConnectionDialog(IPowerBIService powerBIService)
    {
        InitializeComponent();
        _viewModel = new ConnectionDialogViewModel(powerBIService, this);
        DataContext = _viewModel;
    }

    private void DataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (_viewModel.ConnectCommand.CanExecute(null))
        {
            _viewModel.ConnectCommand.Execute(null);
        }
    }
}
