using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using ClosedXML.Excel;
using Microsoft.Win32;
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

    // Measures column visibility settings
    private bool _isColumnSettingsOpen;
    private bool _showNameColumn = true;
    private bool _showTableColumn = true;
    private bool _showExpressionColumn = true;
    private bool _showDescriptionColumn = true;
    private bool _showFormatStringColumn = true;
    private bool _showIsHiddenColumn = true;
    private bool _showDisplayFolderColumn = true;
    private bool _showDataTypeColumn = true;
    
    // Additional column visibility (default unchecked)
    private bool _showDetailRowsExpressionColumn = false;
    private bool _showKPIColumn = false;
    private bool _showStateColumn = false;
    private bool _showErrorMessageColumn = false;
    private bool _showLineageTagColumn = false;
    private bool _showModifiedTimeColumn = false;

    // DAX Viewer properties
    private bool _isDaxViewerOpen;
    private string _selectedDaxExpression = string.Empty;
    private string _selectedMeasureName = string.Empty;
    private bool _isFormatting;

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
                (ExportMeasuresCommand as RelayCommand)?.RaiseCanExecuteChanged();
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
                (ExportMeasuresCommand as RelayCommand)?.RaiseCanExecuteChanged();
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

    // Measures Column Visibility Properties
    public bool IsColumnSettingsOpen
    {
        get => _isColumnSettingsOpen;
        set
        {
            if (_isColumnSettingsOpen != value)
            {
                _isColumnSettingsOpen = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowNameColumn
    {
        get => _showNameColumn;
        set
        {
            if (_showNameColumn != value)
            {
                _showNameColumn = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowTableColumn
    {
        get => _showTableColumn;
        set
        {
            if (_showTableColumn != value)
            {
                _showTableColumn = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowExpressionColumn
    {
        get => _showExpressionColumn;
        set
        {
            if (_showExpressionColumn != value)
            {
                _showExpressionColumn = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowDescriptionColumn
    {
        get => _showDescriptionColumn;
        set
        {
            if (_showDescriptionColumn != value)
            {
                _showDescriptionColumn = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowFormatStringColumn
    {
        get => _showFormatStringColumn;
        set
        {
            if (_showFormatStringColumn != value)
            {
                _showFormatStringColumn = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowIsHiddenColumn
    {
        get => _showIsHiddenColumn;
        set
        {
            if (_showIsHiddenColumn != value)
            {
                _showIsHiddenColumn = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowDisplayFolderColumn
    {
        get => _showDisplayFolderColumn;
        set
        {
            if (_showDisplayFolderColumn != value)
            {
                _showDisplayFolderColumn = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowDataTypeColumn
    {
        get => _showDataTypeColumn;
        set
        {
            if (_showDataTypeColumn != value)
            {
                _showDataTypeColumn = value;
                OnPropertyChanged();
            }
        }
    }

    // Additional column visibility properties (default unchecked)
    public bool ShowDetailRowsExpressionColumn
    {
        get => _showDetailRowsExpressionColumn;
        set
        {
            if (_showDetailRowsExpressionColumn != value)
            {
                _showDetailRowsExpressionColumn = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowKPIColumn
    {
        get => _showKPIColumn;
        set
        {
            if (_showKPIColumn != value)
            {
                _showKPIColumn = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowStateColumn
    {
        get => _showStateColumn;
        set
        {
            if (_showStateColumn != value)
            {
                _showStateColumn = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowErrorMessageColumn
    {
        get => _showErrorMessageColumn;
        set
        {
            if (_showErrorMessageColumn != value)
            {
                _showErrorMessageColumn = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowLineageTagColumn
    {
        get => _showLineageTagColumn;
        set
        {
            if (_showLineageTagColumn != value)
            {
                _showLineageTagColumn = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowModifiedTimeColumn
    {
        get => _showModifiedTimeColumn;
        set
        {
            if (_showModifiedTimeColumn != value)
            {
                _showModifiedTimeColumn = value;
                OnPropertyChanged();
            }
        }
    }

    // DAX Viewer Properties
    public bool IsDaxViewerOpen
    {
        get => _isDaxViewerOpen;
        set
        {
            if (_isDaxViewerOpen != value)
            {
                _isDaxViewerOpen = value;
                OnPropertyChanged();
            }
        }
    }

    public string SelectedDaxExpression
    {
        get => _selectedDaxExpression;
        set
        {
            if (_selectedDaxExpression != value)
            {
                _selectedDaxExpression = value;
                OnPropertyChanged();
                (FormatDaxCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
        }
    }

    public string SelectedMeasureName
    {
        get => _selectedMeasureName;
        set
        {
            if (_selectedMeasureName != value)
            {
                _selectedMeasureName = value;
                OnPropertyChanged();
            }
        }
    }

    public bool IsFormatting
    {
        get => _isFormatting;
        set
        {
            if (_isFormatting != value)
            {
                _isFormatting = value;
                OnPropertyChanged();
                (FormatDaxCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
        }
    }

    public ICommand ConnectCommand { get; }
    public ICommand ClearMeasuresSearchCommand { get; }
    public ICommand ToggleColumnSettingsCommand { get; }
    public ICommand ExportMeasuresCommand { get; }
    public ICommand ViewDaxCommand { get; }
    public ICommand CopyDaxCommand { get; }
    public ICommand CloseDaxViewerCommand { get; }
    public ICommand FormatDaxCommand { get; }

    public MainViewModel(IPowerBIService powerBIService)
    {
        _powerBIService = powerBIService;
        ConnectCommand = new RelayCommand(async () => await ConnectAsync(), () => !IsConnecting);
        ClearMeasuresSearchCommand = new RelayCommand(async () => await Task.Run(() => MeasuresSearchText = string.Empty), () => true);
        ToggleColumnSettingsCommand = new RelayCommand(async () => await Task.Run(() => IsColumnSettingsOpen = !IsColumnSettingsOpen), () => true);
        ExportMeasuresCommand = new RelayCommand(async () => await ExportMeasuresToExcelAsync(), () => IsConnected && FilteredMeasures.Count > 0);
        ViewDaxCommand = new RelayCommandWithParameter(ViewDax);
        CopyDaxCommand = new RelayCommandWithParameter(CopyDax);
        CloseDaxViewerCommand = new RelayCommand(async () => await Task.Run(() => IsDaxViewerOpen = false), () => true);
        FormatDaxCommand = new RelayCommand(async () => await FormatDaxAsync(), () => !IsFormatting && !string.IsNullOrEmpty(SelectedDaxExpression));
    }

    private void ViewDax(object? parameter)
    {
        if (parameter is MeasureInfo measure)
        {
            SelectedMeasureName = measure.Name;
            SelectedDaxExpression = measure.Expression ?? string.Empty;
            IsDaxViewerOpen = true;
        }
    }

    private void CopyDax(object? parameter)
    {
        string? expression = null;
        
        if (parameter is MeasureInfo measure)
        {
            expression = measure.Expression;
        }
        else if (parameter is string str)
        {
            expression = str;
        }
        
        if (!string.IsNullOrEmpty(expression))
        {
            try
            {
                Clipboard.SetText(expression);
                StatusMessage = "? DAX expression copied to clipboard!";
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Copy error: {ex}");
                StatusMessage = "Failed to copy to clipboard.";
            }
        }
    }

    private async Task FormatDaxAsync()
    {
        if (string.IsNullOrWhiteSpace(SelectedDaxExpression))
            return;

        try
        {
            IsFormatting = true;
            StatusMessage = "Formatting DAX expression...";
            
            var formatterService = new DaxFormatterService();
            var formattedDax = await formatterService.FormatDaxAsync(SelectedDaxExpression);
            
            SelectedDaxExpression = formattedDax;
            StatusMessage = "? DAX expression formatted!";
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Format error: {ex}");
            StatusMessage = "Failed to format DAX expression.";
        }
        finally
        {
            IsFormatting = false;
        }
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
            // Load model overview
            StatusMessage = $"Loading model overview from {instance.DisplayName}...";
            ModelOverview = await _powerBIService.GetModelOverviewAsync(instance.Port, cancellationToken);
            
            // Load measures with all properties including Description
            StatusMessage = $"Loading measures from {instance.DisplayName}...";
            var measures = await _powerBIService.GetMeasuresAsync(instance.Port, cancellationToken);
            Measures = new ObservableCollection<MeasureInfo>(measures);
            
            // Load columns
            StatusMessage = $"Loading columns from {instance.DisplayName}...";
            var columns = await _powerBIService.GetColumnsAsync(instance.Port, cancellationToken);
            Columns = new ObservableCollection<ColumnInfo>(columns);
            
            // Load relationships
            StatusMessage = $"Loading relationships from {instance.DisplayName}...";
            var relationships = await _powerBIService.GetRelationshipsAsync(instance.Port, cancellationToken);
            Relationships = new ObservableCollection<RelationshipInfo>(relationships);
            
            // Load other data (dependencies, unused objects, impact analysis - these require more complex analysis)
            LoadAnalysisData();
            
            StatusMessage = $"? Connected to {instance.DisplayName}! Loaded {ModelOverview.MeasureCount} measures, {ModelOverview.ColumnCount} columns, {ModelOverview.RelationshipCount} relationships";
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

    private void LoadAnalysisData()
    {
        // Dependencies - analyze measure expressions for references
        var dependencies = new List<DependencyInfo>();
        foreach (var measure in Measures)
        {
            if (!string.IsNullOrEmpty(measure.Expression))
            {
                // Find measure references in the expression
                foreach (var otherMeasure in Measures.Where(m => m.Name != measure.Name))
                {
                    if (measure.Expression.Contains($"[{otherMeasure.Name}]"))
                    {
                        dependencies.Add(new DependencyInfo
                        {
                            ObjectName = measure.Name,
                            ObjectType = "Measure",
                            DependsOn = otherMeasure.Name,
                            DependencyType = "Measure Reference"
                        });
                    }
                }
            }
        }
        Dependencies = new ObservableCollection<DependencyInfo>(dependencies);

        // Unused Objects - find hidden measures and columns with no references
        var unusedObjects = new List<UnusedObjectInfo>();
        foreach (var measure in Measures.Where(m => m.IsHidden))
        {
            var isReferenced = Measures.Any(m => m.Expression?.Contains($"[{measure.Name}]") == true);
            if (!isReferenced)
            {
                unusedObjects.Add(new UnusedObjectInfo
                {
                    Name = measure.Name,
                    ObjectType = "Measure",
                    Table = measure.Table,
                    Reason = "Hidden measure with no references"
                });
            }
        }
        UnusedObjects = new ObservableCollection<UnusedObjectInfo>(unusedObjects);

        // Impact Analysis - analyze what depends on each measure
        var impactAnalysis = new List<ImpactAnalysisInfo>();
        foreach (var measure in Measures)
        {
            var dependentMeasures = Measures.Where(m => 
                m.Expression?.Contains($"[{measure.Name}]") == true && m.Name != measure.Name);
            
            foreach (var dependent in dependentMeasures)
            {
                impactAnalysis.Add(new ImpactAnalysisInfo
                {
                    ObjectName = measure.Name,
                    ObjectType = "Measure",
                    ImpactedObject = dependent.Name,
                    ImpactType = "Direct Dependency",
                    Severity = "High"
                });
            }
        }
        ImpactAnalysis = new ObservableCollection<ImpactAnalysisInfo>(impactAnalysis);
    }

    private async Task ExportMeasuresToExcelAsync()
    {
        try
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = $"Measures_Export_{DateTime.Now:yyyyMMdd_HHmmss}"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                await Task.Run(() =>
                {
                    using var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add("Measures");

                    // Track column index for visible columns only
                    int col = 1;

                    // Add S.No column (always visible)
                    worksheet.Cell(1, col).Value = "S.No";
                    col++;

                    // Add headers based on visibility settings
                    var columnMappings = new List<(string Header, Func<MeasureInfo, object?> ValueGetter, bool IsVisible)>
                    {
                        ("Name", m => m.Name, ShowNameColumn),
                        ("Table", m => m.Table, ShowTableColumn),
                        ("Expression", m => m.Expression, ShowExpressionColumn),
                        ("Description", m => m.Description, ShowDescriptionColumn),
                        ("Format String", m => m.FormatString, ShowFormatStringColumn),
                        ("Is Hidden", m => m.IsHidden ? "Hidden" : "Visible", ShowIsHiddenColumn),
                        ("Display Folder", m => m.DisplayFolder, ShowDisplayFolderColumn),
                        ("Data Type", m => m.DataType, ShowDataTypeColumn),
                        ("Detail Rows Expression", m => m.DetailRowsExpression, ShowDetailRowsExpressionColumn),
                        ("KPI", m => m.KPI, ShowKPIColumn),
                        ("State", m => m.State, ShowStateColumn),
                        ("Error Message", m => m.ErrorMessage, ShowErrorMessageColumn),
                        ("Lineage Tag", m => m.LineageTag, ShowLineageTagColumn),
                        ("Modified Time", m => m.ModifiedTime?.ToString("g"), ShowModifiedTimeColumn)
                    };

                    // Add visible column headers
                    var visibleColumns = columnMappings.Where(c => c.IsVisible).ToList();
                    foreach (var column in visibleColumns)
                    {
                        worksheet.Cell(1, col).Value = column.Header;
                        col++;
                    }

                    // Style header row
                    var headerRange = worksheet.Range(1, 1, 1, 1 + visibleColumns.Count);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#0078D4");
                    headerRange.Style.Font.FontColor = XLColor.White;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Add data rows
                    int row = 2;
                    int serialNumber = 1;
                    foreach (var measure in FilteredMeasures)
                    {
                        col = 1;

                        // S.No
                        worksheet.Cell(row, col).Value = serialNumber++;
                        col++;

                        // Add visible column values
                        foreach (var column in visibleColumns)
                        {
                            var value = column.ValueGetter(measure);
                            if (value != null)
                            {
                                worksheet.Cell(row, col).Value = value.ToString();
                            }
                            col++;
                        }
                        row++;
                    }

                    // Auto-fit columns
                    worksheet.Columns().AdjustToContents();

                    // Set maximum column width to prevent extremely wide columns
                    foreach (var column in worksheet.ColumnsUsed())
                    {
                        if (column.Width > 50)
                        {
                            column.Width = 50;
                        }
                    }

                    // Add filters
                    worksheet.RangeUsed()?.SetAutoFilter();

                    workbook.SaveAs(saveFileDialog.FileName);
                });

                MessageBox.Show(
                    $"Measures exported successfully to:\n{saveFileDialog.FileName}",
                    "Export Complete",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                // Open the file location
                Process.Start("explorer.exe", $"/select,\"{saveFileDialog.FileName}\"");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Failed to export measures.\n\nError: {ex.Message}",
                "Export Failed",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Debug.WriteLine($"Export error: {ex}");
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

public class RelayCommandWithParameter : ICommand
{
    private readonly Action<object?> _execute;
    private readonly Func<object?, bool>? _canExecute;

    public event EventHandler? CanExecuteChanged;

    public RelayCommandWithParameter(Action<object?> execute, Func<object?, bool>? canExecute = null)
    {
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute;
    }

    public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;

    public void Execute(object? parameter) => _execute(parameter);

    public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
}

