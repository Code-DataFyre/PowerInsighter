using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using PowerInsighter.Models;
using PowerInsighter.Services;

namespace PowerInsighter.ViewModels;

public class ConnectionDialogViewModel : INotifyPropertyChanged
{
    private readonly IPowerBIService _powerBIService;
    private readonly Window _window;
    private ObservableCollection<PowerBIInstance> _instances = [];
    private ObservableCollection<PowerBIInstance> _filteredInstances = [];
    private PowerBIInstance? _selectedInstance;
    private string _searchText = string.Empty;
    private bool _isLoading;
    private string _statusMessage = "Loading Power BI instances...";

    public event PropertyChangedEventHandler? PropertyChanged;

    public ObservableCollection<PowerBIInstance> Instances
    {
        get => _instances;
        set
        {
            _instances = value;
            OnPropertyChanged();
            ApplyFilter();
        }
    }

    public ObservableCollection<PowerBIInstance> FilteredInstances
    {
        get => _filteredInstances;
        set
        {
            _filteredInstances = value;
            OnPropertyChanged();
        }
    }

    public PowerBIInstance? SelectedInstance
    {
        get => _selectedInstance;
        set
        {
            _selectedInstance = value;
            OnPropertyChanged();
            ((RelayCommand)ConnectCommand).RaiseCanExecuteChanged();
        }
    }

    public string SearchText
    {
        get => _searchText;
        set
        {
            _searchText = value;
            OnPropertyChanged();
            ApplyFilter();
        }
    }

    public bool IsLoading
    {
        get => _isLoading;
        set
        {
            _isLoading = value;
            OnPropertyChanged();
        }
    }

    public string StatusMessage
    {
        get => _statusMessage;
        set
        {
            _statusMessage = value;
            OnPropertyChanged();
        }
    }

    public ICommand RefreshCommand { get; }
    public ICommand ConnectCommand { get; }
    public ICommand CancelCommand { get; }

    public PowerBIInstance? Result { get; private set; }

    public ConnectionDialogViewModel(IPowerBIService powerBIService, Window window)
    {
        _powerBIService = powerBIService;
        _window = window;

        RefreshCommand = new RelayCommand(async () => await LoadInstancesAsync(), () => !IsLoading);
        ConnectCommand = new RelayCommand(async () => await ConnectAsync(), () => SelectedInstance != null);
        CancelCommand = new RelayCommand(async () => await Task.Run(() => Cancel()), () => true);

        _ = LoadInstancesAsync();
    }

    private async Task LoadInstancesAsync()
    {
        try
        {
            IsLoading = true;
            StatusMessage = "Searching for Power BI instances...";

            if (!_powerBIService.IsPowerBIRunning())
            {
                StatusMessage = "No Power BI Desktop instances found.\n\nPlease start Power BI Desktop and open a .pbix file.";
                Instances = [];
                return;
            }

            var instances = await _powerBIService.FindPowerBIInstancesAsync();

            if (instances.Count == 0)
            {
                StatusMessage = "No Power BI instances with open models found.\n\n" +
                               "Make sure:\n" +
                               "• A .pbix file is OPEN in Power BI Desktop\n" +
                               "• The data model has loaded\n" +
                               "• Wait 10-15 seconds after opening the file";
            }

            Instances = new ObservableCollection<PowerBIInstance>(instances);
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error loading instances: {ex.Message}";
            Debug.WriteLine($"Error loading instances: {ex}");
        }
        finally
        {
            IsLoading = false;
            ((RelayCommand)RefreshCommand).RaiseCanExecuteChanged();
        }
    }

    private void ApplyFilter()
    {
        if (string.IsNullOrWhiteSpace(SearchText))
        {
            FilteredInstances = new ObservableCollection<PowerBIInstance>(Instances);
        }
        else
        {
            var searchLower = SearchText.ToLowerInvariant();
            var filtered = Instances.Where(i =>
                i.DisplayName.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ||
                i.Port.ToString().Contains(searchLower) ||
                (i.FileName?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (i.WindowTitle?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false)
            ).ToList();

            FilteredInstances = new ObservableCollection<PowerBIInstance>(filtered);
        }
    }

    private async Task ConnectAsync()
    {
        if (SelectedInstance != null)
        {
            Result = SelectedInstance;
            await Task.Run(() => _window.Dispatcher.Invoke(() => _window.DialogResult = true));
        }
    }

    private void Cancel()
    {
        _window.Dispatcher.Invoke(() => _window.DialogResult = false);
    }

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
