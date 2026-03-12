using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using PowerInsighter.Models;
using PowerInsighter.Services;

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
            IsConnecting = true;
            StatusMessage = "Searching for Power BI instance...";

            if (!_powerBIService.IsPowerBIRunning())
            {
                MessageBox.Show(
                    "Power BI Desktop is not running.\n\n" +
                    "Please start Power BI Desktop and open a .pbix file first.",
                    "Power BI Not Found",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                StatusMessage = "Power BI Desktop not running.";
                return;
            }

            StatusMessage = "Scanning for Power BI ports...";
            var ports = await _powerBIService.FindPowerBIPortsAsync(cancellationToken);

            if (ports.Count == 0)
            {
                MessageBox.Show(
                    "Could not find any Power BI Analysis Services instances.\n\n" +
                    "Make sure:\n" +
                    "1. A .pbix file is OPEN in Power BI Desktop\n" +
                    "2. The data model has loaded (check if you see tables in Fields pane)\n" +
                    "3. Wait 10-15 seconds after opening the file",
                    "No Instances Found",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                StatusMessage = "No Power BI instance found.";
                return;
            }

            // Try connecting to each port
            foreach (var port in ports)
            {
                cancellationToken.ThrowIfCancellationRequested();
                StatusMessage = $"Trying port {port}...";
                
                try
                {
                    await ConnectAndLoadMetadataAsync(port, cancellationToken);
                    return; // Success!
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to connect to port {port}: {ex.Message}");
                }
            }

            MessageBox.Show(
                "Found ports but couldn't connect to any.\n\n" +
                $"Ports tried: {string.Join(", ", ports)}",
                "Connection Failed",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            StatusMessage = "Connection failed.";
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

    private async Task ConnectAndLoadMetadataAsync(int port, CancellationToken cancellationToken)
    {
        StatusMessage = $"Connecting to port {port}...";
        
        var metadataList = await _powerBIService.LoadMetadataAsync(port, cancellationToken);

        Metadata = new ObservableCollection<ModelMetadata>(metadataList);
        StatusMessage = $"? Connected to port {port}! Loaded {metadataList.Count} items";
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

