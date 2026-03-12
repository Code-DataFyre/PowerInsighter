using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using PowerInsighter.Models;
using PowerInsighter.Services;
using PowerInsighter.Views;

namespace PowerInsighter.ViewModels;

public class MainViewModel : INotifyPropertyChanged
{
    private readonly IPowerBIService _powerBIService;
    private CancellationTokenSource? _cancellationTokenSource;
    private bool _isConnecting;
    private string _statusMessage = "Status: Not Connected";
    private ObservableCollection<ModelMetadata> _metadata = [];

    public event PropertyChangedEventHandler? PropertyChanged;

    public bool IsConnecting
    {
        get => _isConnecting;
        set
        {
            if (_isConnecting != value)
            {
                _isConnecting = value;
                OnPropertyChanged();
                ((RelayCommand)ConnectCommand).RaiseCanExecuteChanged();
            }
        }
    }

    public string StatusMessage
    {
        get => _statusMessage;
        set
        {
            if (_statusMessage != value)
            {
                _statusMessage = value;
                OnPropertyChanged();
            }
        }
    }

    public ObservableCollection<ModelMetadata> Metadata
    {
        get => _metadata;
        set
        {
            if (_metadata != value)
            {
                _metadata = value;
                OnPropertyChanged();
            }
        }
    }

    public ICommand ConnectCommand { get; }

    public MainViewModel(IPowerBIService powerBIService)
    {
        _powerBIService = powerBIService;
        ConnectCommand = new RelayCommand(async () => await ConnectAsync(), () => !IsConnecting);
    }

    private async Task ConnectAsync()
    {
        _cancellationTokenSource?.Cancel();
        _cancellationTokenSource = new CancellationTokenSource();
        var cancellationToken = _cancellationTokenSource.Token;

        try
        {
            StatusMessage = "Opening connection dialog...";

            // Create the dialog first, then create the ViewModel with the dialog reference
            var dialog = new ConnectionDialog(_powerBIService)
            {
                Owner = Application.Current.MainWindow
            };

            var result = dialog.ShowDialog();

            if (result == true && dialog.SelectedInstance != null)
            {
                var selectedInstance = dialog.SelectedInstance;
                IsConnecting = true;
                await ConnectToInstanceAsync(selectedInstance, cancellationToken);
            }
            else
            {
                StatusMessage = "Connection cancelled.";
            }
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Connection cancelled.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Error: {ex.Message}",
                "Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            StatusMessage = "Error occurred.";
            Debug.WriteLine($"Connection error: {ex}");
        }
        finally
        {
            IsConnecting = false;
        }
    }

    private async Task ConnectToInstanceAsync(PowerBIInstance instance, CancellationToken cancellationToken)
    {
        StatusMessage = $"Connecting to {instance.DisplayName}...";

        try
        {
            var metadataList = await _powerBIService.LoadMetadataAsync(instance.Port, cancellationToken);

            Metadata = new ObservableCollection<ModelMetadata>(metadataList);
            StatusMessage = $"? Connected to {instance.DisplayName}! Loaded {metadataList.Count} items";
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Failed to connect to {instance.DisplayName}.\n\nError: {ex.Message}",
                "Connection Failed",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            StatusMessage = $"Failed to connect to {instance.DisplayName}.";
            Debug.WriteLine($"Connection error: {ex}");
            throw;
        }
    }

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

public class RelayCommand : ICommand
{
    private readonly Func<Task> _execute;
    private readonly Func<bool> _canExecute;

    public event EventHandler? CanExecuteChanged;

    public RelayCommand(Func<Task> execute, Func<bool> canExecute)
    {
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute ?? throw new ArgumentNullException(nameof(canExecute));
    }

    public bool CanExecute(object? parameter) => _canExecute();

    public async void Execute(object? parameter) => await _execute();

    public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
}

