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
    private bool _isConnected;
    private string _statusMessage = "Status: Not Connected";
    private ObservableCollection<ModelMetadata> _metadata = [];
    
    // Tab data collections
    private ModelOverview? _modelOverview;
    private ObservableCollection<MeasureInfo> _measures = [];
    private ObservableCollection<MeasureInfo> _filteredMeasures = [];
    private string _measuresSearchText = string.Empty;
    private ObservableCollection<ColumnInfo> _columns = [];
    private ObservableCollection<RelationshipInfo> _relationships = [];
    private ObservableCollection<DependencyInfo> _dependencies = [];
    private ObservableCollection<UnusedObjectInfo> _unusedObjects = [];
    private ObservableCollection<ImpactAnalysisInfo> _impactAnalysis = [];

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

    public bool IsConnected
    {
        get => _isConnected;
        set
        {
            if (_isConnected != value)
            {
                _isConnected = value;
                OnPropertyChanged();
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

    // Tab Properties
    public ModelOverview? ModelOverview
    {
        get => _modelOverview;
        set
        {
            if (_modelOverview != value)
            {
                _modelOverview = value;
                OnPropertyChanged();
            }
        }
    }

    public ObservableCollection<MeasureInfo> Measures
    {
        get => _measures;
        set
        {
            if (_measures != value)
            {
                _measures = value;
                OnPropertyChanged();
                ApplyMeasuresFilter();
            }
        }
    }

    public ObservableCollection<MeasureInfo> FilteredMeasures
    {
        get => _filteredMeasures;
        set
        {
            if (_filteredMeasures != value)
            {
                _filteredMeasures = value;
                OnPropertyChanged();
            }
        }
    }

    public string MeasuresSearchText
    {
        get => _measuresSearchText;
        set
        {
            if (_measuresSearchText != value)
            {
                _measuresSearchText = value;
                OnPropertyChanged();
                ApplyMeasuresFilter();
            }
        }
    }

    private void ApplyMeasuresFilter()
    {
        if (string.IsNullOrWhiteSpace(MeasuresSearchText))
        {
            FilteredMeasures = new ObservableCollection<MeasureInfo>(Measures);
        }
        else
        {
            var searchLower = MeasuresSearchText.ToLowerInvariant();
            var filtered = Measures.Where(m =>
                m.Name.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ||
                m.Table.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ||
                (m.Expression?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (m.Description?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false)
            ).ToList();

            FilteredMeasures = new ObservableCollection<MeasureInfo>(filtered);
        }
    }

    public ObservableCollection<ColumnInfo> Columns
    {
        get => _columns;
        set
        {
            if (_columns != value)
            {
                _columns = value;
                OnPropertyChanged();
            }
        }
    }

    public ObservableCollection<RelationshipInfo> Relationships
    {
        get => _relationships;
        set
        {
            if (_relationships != value)
            {
                _relationships = value;
                OnPropertyChanged();
            }
        }
    }

    public ObservableCollection<DependencyInfo> Dependencies
    {
        get => _dependencies;
        set
        {
            if (_dependencies != value)
            {
                _dependencies = value;
                OnPropertyChanged();
            }
        }
    }

    public ObservableCollection<UnusedObjectInfo> UnusedObjects
    {
        get => _unusedObjects;
        set
        {
            if (_unusedObjects != value)
            {
                _unusedObjects = value;
                OnPropertyChanged();
            }
        }
    }

    public ObservableCollection<ImpactAnalysisInfo> ImpactAnalysis
    {
        get => _impactAnalysis;
        set
        {
            if (_impactAnalysis != value)
            {
                _impactAnalysis = value;
                OnPropertyChanged();
            }
        }
    }

    public ICommand ConnectCommand { get; }
    public ICommand ClearMeasuresSearchCommand { get; }

    public MainViewModel(IPowerBIService powerBIService)
    {
        _powerBIService = powerBIService;
        ConnectCommand = new RelayCommand(async () => await ConnectAsync(), () => !IsConnecting);
        ClearMeasuresSearchCommand = new RelayCommand(async () => await Task.Run(() => MeasuresSearchText = string.Empty), () => true);
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
            
            // Load dummy data for all tabs
            LoadDummyData(instance.DisplayName);
            IsConnected = true;
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

    private void LoadDummyData(string modelName)
    {
        // Model Overview
        ModelOverview = new ModelOverview
        {
            ModelName = modelName,
            TableCount = 12,
            MeasureCount = 45,
            ColumnCount = 156,
            RelationshipCount = 18,
            CalculatedColumnCount = 23,
            CalculatedTableCount = 3,
            ModelSize = 125_000_000,
            LastRefresh = DateTime.Now.AddHours(-2),
            CompatibilityLevel = "1567"
        };

        // Measures
        Measures = new ObservableCollection<MeasureInfo>
        {
            new() { Name = "Total Sales", Table = "Sales", Expression = "SUM(Sales[Amount])", Description = "Sum of all sales amounts", FormatString = "$#,##0.00", IsHidden = false },
            new() { Name = "Total Cost", Table = "Sales", Expression = "SUM(Sales[Cost])", Description = "Sum of all costs", FormatString = "$#,##0.00", IsHidden = false },
            new() { Name = "Profit Margin", Table = "Sales", Expression = "DIVIDE([Total Sales] - [Total Cost], [Total Sales])", Description = "Profit margin percentage", FormatString = "0.00%", IsHidden = false },
            new() { Name = "YTD Sales", Table = "Sales", Expression = "TOTALYTD([Total Sales], 'Date'[Date])", Description = "Year to date sales", FormatString = "$#,##0.00", IsHidden = false },
            new() { Name = "Previous Year Sales", Table = "Sales", Expression = "CALCULATE([Total Sales], SAMEPERIODLASTYEAR('Date'[Date]))", Description = "Sales from previous year", FormatString = "$#,##0.00", IsHidden = false },
            new() { Name = "Sales Growth %", Table = "Sales", Expression = "DIVIDE([Total Sales] - [Previous Year Sales], [Previous Year Sales])", Description = "Year over year sales growth", FormatString = "0.00%", IsHidden = false },
            new() { Name = "Customer Count", Table = "Customers", Expression = "DISTINCTCOUNT(Sales[CustomerID])", Description = "Count of unique customers", FormatString = "#,##0", IsHidden = false },
            new() { Name = "Average Order Value", Table = "Sales", Expression = "DIVIDE([Total Sales], COUNTROWS(Sales))", Description = "Average value per order", FormatString = "$#,##0.00", IsHidden = false },
            new() { Name = "_Helper Measure", Table = "Calculations", Expression = "1", Description = "Internal helper measure", FormatString = "", IsHidden = true },
        };

        // Columns
        Columns = new ObservableCollection<ColumnInfo>
        {
            new() { Name = "ProductID", Table = "Products", DataType = "Int64", IsCalculated = false, IsHidden = false, Description = "Product identifier" },
            new() { Name = "ProductName", Table = "Products", DataType = "String", IsCalculated = false, IsHidden = false, Description = "Name of the product" },
            new() { Name = "Category", Table = "Products", DataType = "String", IsCalculated = false, IsHidden = false, Description = "Product category" },
            new() { Name = "UnitPrice", Table = "Products", DataType = "Decimal", IsCalculated = false, IsHidden = false, Description = "Price per unit" },
            new() { Name = "FullName", Table = "Customers", DataType = "String", Expression = "[FirstName] & \" \" & [LastName]", IsCalculated = true, IsHidden = false, Description = "Customer full name" },
            new() { Name = "OrderDate", Table = "Sales", DataType = "DateTime", IsCalculated = false, IsHidden = false, Description = "Date of the order" },
            new() { Name = "Amount", Table = "Sales", DataType = "Decimal", IsCalculated = false, IsHidden = false, Description = "Order amount" },
            new() { Name = "Year", Table = "Date", DataType = "Int64", Expression = "YEAR([Date])", IsCalculated = true, IsHidden = false, Description = "Year number" },
            new() { Name = "Month", Table = "Date", DataType = "String", Expression = "FORMAT([Date], \"MMMM\")", IsCalculated = true, IsHidden = false, Description = "Month name" },
            new() { Name = "Quarter", Table = "Date", DataType = "String", Expression = "\"Q\" & QUARTER([Date])", IsCalculated = true, IsHidden = false, Description = "Quarter label" },
        };

        // Relationships
        Relationships = new ObservableCollection<RelationshipInfo>
        {
            new() { FromTable = "Sales", FromColumn = "ProductID", ToTable = "Products", ToColumn = "ProductID", Cardinality = "Many to One", CrossFilterDirection = "Single", IsActive = true },
            new() { FromTable = "Sales", FromColumn = "CustomerID", ToTable = "Customers", ToColumn = "CustomerID", Cardinality = "Many to One", CrossFilterDirection = "Single", IsActive = true },
            new() { FromTable = "Sales", FromColumn = "OrderDate", ToTable = "Date", ToColumn = "Date", Cardinality = "Many to One", CrossFilterDirection = "Single", IsActive = true },
            new() { FromTable = "Products", FromColumn = "CategoryID", ToTable = "Categories", ToColumn = "CategoryID", Cardinality = "Many to One", CrossFilterDirection = "Both", IsActive = true },
            new() { FromTable = "Customers", FromColumn = "RegionID", ToTable = "Regions", ToColumn = "RegionID", Cardinality = "Many to One", CrossFilterDirection = "Single", IsActive = true },
            new() { FromTable = "Sales", FromColumn = "ShipDate", ToTable = "Date", ToColumn = "Date", Cardinality = "Many to One", CrossFilterDirection = "Single", IsActive = false },
        };

        // Dependencies
        Dependencies = new ObservableCollection<DependencyInfo>
        {
            new() { ObjectName = "Profit Margin", ObjectType = "Measure", DependsOn = "Total Sales", DependencyType = "Measure Reference" },
            new() { ObjectName = "Profit Margin", ObjectType = "Measure", DependsOn = "Total Cost", DependencyType = "Measure Reference" },
            new() { ObjectName = "YTD Sales", ObjectType = "Measure", DependsOn = "Total Sales", DependencyType = "Measure Reference" },
            new() { ObjectName = "YTD Sales", ObjectType = "Measure", DependsOn = "Date[Date]", DependencyType = "Column Reference" },
            new() { ObjectName = "Sales Growth %", ObjectType = "Measure", DependsOn = "Total Sales", DependencyType = "Measure Reference" },
            new() { ObjectName = "Sales Growth %", ObjectType = "Measure", DependsOn = "Previous Year Sales", DependencyType = "Measure Reference" },
            new() { ObjectName = "FullName", ObjectType = "Calculated Column", DependsOn = "FirstName", DependencyType = "Column Reference" },
            new() { ObjectName = "FullName", ObjectType = "Calculated Column", DependsOn = "LastName", DependencyType = "Column Reference" },
            new() { ObjectName = "Year", ObjectType = "Calculated Column", DependsOn = "Date[Date]", DependencyType = "Column Reference" },
        };

        // Unused Objects
        UnusedObjects = new ObservableCollection<UnusedObjectInfo>
        {
            new() { Name = "_Helper Measure", ObjectType = "Measure", Table = "Calculations", Reason = "Not referenced in any visual or other measure" },
            new() { Name = "OldProductCode", ObjectType = "Column", Table = "Products", Reason = "Hidden column with no references" },
            new() { Name = "TempCalc", ObjectType = "Calculated Column", Table = "Sales", Reason = "Not used in any visual or measure" },
            new() { Name = "Backup_Sales", ObjectType = "Table", Table = "-", Reason = "Hidden table with no active relationships" },
            new() { Name = "LegacyCustomerID", ObjectType = "Column", Table = "Customers", Reason = "Deprecated column, no longer used" },
        };

        // Impact Analysis
        ImpactAnalysis = new ObservableCollection<ImpactAnalysisInfo>
        {
            new() { ObjectName = "Total Sales", ObjectType = "Measure", ImpactedObject = "Profit Margin", ImpactType = "Direct Dependency", Severity = "High" },
            new() { ObjectName = "Total Sales", ObjectType = "Measure", ImpactedObject = "YTD Sales", ImpactType = "Direct Dependency", Severity = "High" },
            new() { ObjectName = "Total Sales", ObjectType = "Measure", ImpactedObject = "Sales Growth %", ImpactType = "Direct Dependency", Severity = "High" },
            new() { ObjectName = "Total Sales", ObjectType = "Measure", ImpactedObject = "Average Order Value", ImpactType = "Direct Dependency", Severity = "High" },
            new() { ObjectName = "Date[Date]", ObjectType = "Column", ImpactedObject = "Year", ImpactType = "Column Reference", Severity = "Medium" },
            new() { ObjectName = "Date[Date]", ObjectType = "Column", ImpactedObject = "Month", ImpactType = "Column Reference", Severity = "Medium" },
            new() { ObjectName = "Date[Date]", ObjectType = "Column", ImpactedObject = "Quarter", ImpactType = "Column Reference", Severity = "Medium" },
            new() { ObjectName = "Products", ObjectType = "Table", ImpactedObject = "Sales", ImpactType = "Relationship", Severity = "Critical" },
            new() { ObjectName = "Customers", ObjectType = "Table", ImpactedObject = "Sales", ImpactType = "Relationship", Severity = "Critical" },
        };
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

