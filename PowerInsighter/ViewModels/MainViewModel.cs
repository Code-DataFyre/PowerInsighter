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
using Microsoft.AnalysisServices.Tabular;
using System.Net.Http;

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
    private ObservableCollection<ColumnInfo> _filteredColumns = [];
    private string _columnsSearchText = string.Empty;
    private ObservableCollection<RelationshipInfo> _relationships = [];
    private ObservableCollection<DependencyInfo> _dependencies = [];
    private ObservableCollection<UnusedObjectInfo> _unusedObjects = [];
    private ObservableCollection<ImpactAnalysisInfo> _impactAnalysis = [];

    private ObservableCollection<BestPracticeViolation> _bestPracticeViolations = [];
    private ObservableCollection<BestPracticeViolation> _filteredBestPracticeViolations = [];
    private string _bestPracticesSearchText = string.Empty;

    // Grouped Best Practices view
    private ObservableCollection<BestPracticeRuleGroup> _bestPracticeRuleGroups = [];
    private int _totalRules;
    private int _violatedRules;
    private int _totalViolationScore;

    private ObservableCollection<DependencyInfo> _filteredDependencies = [];
    private string _dependenciesSearchText = string.Empty;
    private ObservableCollection<UnusedObjectInfo> _filteredUnusedObjects = [];
    private string _unusedObjectsSearchText = string.Empty;
    private ObservableCollection<ImpactAnalysisInfo> _filteredImpactAnalysis = [];
    private string _impactAnalysisSearchText = string.Empty;

    public ObservableCollection<BestPracticeViolation> BestPracticeViolations
    {
        get => _bestPracticeViolations;
        set
        {
            if (_bestPracticeViolations != value)
            {
                _bestPracticeViolations = value;
                OnPropertyChanged();
                ApplyBestPracticesFilter();
                GroupBestPracticeViolations();
            }
        }
    }

    public ObservableCollection<BestPracticeViolation> FilteredBestPracticeViolations
    {
        get => _filteredBestPracticeViolations;
        set
        {
            if (_filteredBestPracticeViolations != value)
            {
                _filteredBestPracticeViolations = value;
                OnPropertyChanged();
            }
        }
    }

    public string BestPracticesSearchText
    {
        get => _bestPracticesSearchText;
        set
        {
            if (_bestPracticesSearchText != value)
            {
                _bestPracticesSearchText = value;
                OnPropertyChanged();
                ApplyBestPracticesFilter();
            }
        }
    }

    private void ApplyBestPracticesFilter()
    {
        if (string.IsNullOrWhiteSpace(BestPracticesSearchText))
        {
            FilteredBestPracticeViolations = new ObservableCollection<BestPracticeViolation>(BestPracticeViolations);
            return;
        }

        var search = BestPracticesSearchText.Trim();
        var filtered = BestPracticeViolations.Where(v =>
            (v.RuleName?.Contains(search, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (v.Category?.Contains(search, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (v.Severity?.Contains(search, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (v.ObjectType?.Contains(search, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (v.ObjectName?.Contains(search, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (v.TableName?.Contains(search, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (v.Description?.Contains(search, StringComparison.OrdinalIgnoreCase) ?? false)
        ).ToList();

        FilteredBestPracticeViolations = new ObservableCollection<BestPracticeViolation>(filtered);
    }

    public ObservableCollection<BestPracticeRuleGroup> BestPracticeRuleGroups
    {
        get => _bestPracticeRuleGroups;
        set
        {
            if (_bestPracticeRuleGroups != value)
            {
                _bestPracticeRuleGroups = value;
                OnPropertyChanged();
            }
        }
    }

    public int TotalRules
    {
        get => _totalRules;
        set
        {
            if (_totalRules != value)
            {
                _totalRules = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(RulesViolatedSummary));
            }
        }
    }

    public int ViolatedRules
    {
        get => _violatedRules;
        set
        {
            if (_violatedRules != value)
            {
                _violatedRules = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(RulesViolatedSummary));
            }
        }
    }

    public int TotalViolationScore
    {
        get => _totalViolationScore;
        set
        {
            if (_totalViolationScore != value)
            {
                _totalViolationScore = value;
                OnPropertyChanged();
            }
        }
    }

    public string RulesViolatedSummary => $"Rules violated: {ViolatedRules} of {TotalRules}";

    private void GroupBestPracticeViolations()
    {
        var groups = BestPracticeViolations
            .GroupBy(v => v.RuleName)
            .Select(g =>
            {
                var firstViolation = g.First();
                var severityLevel = GetSeverityLevel(firstViolation.Severity);
                
                return new BestPracticeRuleGroup
                {
                    RuleName = g.Key ?? "Unknown Rule",
                    Category = firstViolation.Category ?? string.Empty,
                    Description = firstViolation.Description ?? string.Empty,
                    Severity = firstViolation.Severity ?? string.Empty,
                    SeverityLevel = severityLevel,
                    ViolationCount = g.Count(),
                    ViolationScore = g.Count() * severityLevel,
                    Violations = new ObservableCollection<BestPracticeViolation>(g),
                    IsExpanded = false
                };
            })
            .OrderByDescending(g => g.SeverityLevel)
            .ThenByDescending(g => g.ViolationCount)
            .ToList();

        BestPracticeRuleGroups = new ObservableCollection<BestPracticeRuleGroup>(groups);
        ViolatedRules = groups.Count;
        TotalViolationScore = groups.Sum(g => g.ViolationScore);
    }

    private int GetSeverityLevel(string? severity)
    {
        return severity?.ToLower() switch
        {
            "error" => 3,
            "warning" => 2,
            "info" or "information" => 1,
            _ => 1
        };
    }

    private bool _isImpactAnalysisSettingsOpen;
    private bool _showIAObjectColumn = true;
    private bool _showIAObjectTypeColumn = true;
    private bool _showIAImpactedObjectColumn = true;
    private bool _showIAImpactTypeColumn = true;
    private bool _showIASeverityColumn = true;

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

    // Columns tab column visibility settings
    private bool _isColumnsSettingsOpen;
    private bool _showColNameColumn = true;
    private bool _showColTableColumn = true;
    private bool _showColDataTypeColumn = true;
    private bool _showColIsCalculatedColumn = true;
    private bool _showColExpressionColumn = true;
    private bool _showColIsHiddenColumn = true;
    private bool _showColDescriptionColumn = true;
    private bool _showColDisplayFolderColumn = true;
    private bool _showColFormatStringColumn = true;
    // Advanced columns (default unchecked)
    private bool _showColSortByColumnColumn = false;
    private bool _showColIsUniqueColumn = false;
    private bool _showColIsNullableColumn = false;
    private bool _showColIsKeyColumn = false;
    private bool _showColSourceColumnColumn = false;
    private bool _showColDataCategoryColumn = false;
    private bool _showColIsAvailableInMDXColumn = false;
    private bool _showColStateColumn = false;
    private bool _showColErrorMessageColumn = false;
    private bool _showColModifiedTimeColumn = false;
    private bool _showColLineageTagColumn = false;
    private bool _showColSummarizeByColumn = false;
    private bool _showColEncodingColumn = false;

    // Relationships tab fields
    private ObservableCollection<RelationshipInfo> _filteredRelationships = [];
    private string _relationshipsSearchText = string.Empty;
    private bool _isRelationshipsSettingsOpen;
    // Relationships column visibility (main - default checked)
    private bool _showRelNameColumn = true;
    private bool _showRelFromTableColumn = true;
    private bool _showRelFromColumnColumn = true;
    private bool _showRelToTableColumn = true;
    private bool _showRelToColumnColumn = true;
    private bool _showRelCardinalityColumn = true;
    private bool _showRelCrossFilterColumn = true;
    private bool _showRelIsActiveColumn = true;
    // Advanced columns (default unchecked)
    private bool _showRelFromCardinalityColumn = false;
    private bool _showRelToCardinalityColumn = false;
    private bool _showRelSecurityFilterColumn = false;
    private bool _showRelJoinOnDateColumn = false;
    private bool _showRelRelyOnRefIntColumn = false;
    private bool _showRelStateColumn = false;
    private bool _showRelModifiedTimeColumn = false;

    // DAX Viewer properties
    private bool _isDaxViewerOpen;
    private string _selectedDaxExpression = string.Empty;
    private string _selectedMeasureName = string.Empty;
    private bool _isFormatting;

    private int? _connectedPort;

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
                (ExportColumnsCommand as RelayCommand)?.RaiseCanExecuteChanged();
                (ExportRelationshipsCommand as RelayCommand)?.RaiseCanExecuteChanged();
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
                ApplyColumnsFilter();
            }
        }
    }

    public ObservableCollection<ColumnInfo> FilteredColumns
    {
        get => _filteredColumns;
        set
        {
            if (_filteredColumns != value)
            {
                _filteredColumns = value;
                OnPropertyChanged();
                (ExportColumnsCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
        }
    }

    public string ColumnsSearchText
    {
        get => _columnsSearchText;
        set
        {
            if (_columnsSearchText != value)
            {
                _columnsSearchText = value;
                OnPropertyChanged();
                ApplyColumnsFilter();
            }
        }
    }

    private void ApplyColumnsFilter()
    {
        if (string.IsNullOrWhiteSpace(ColumnsSearchText))
        {
            FilteredColumns = new ObservableCollection<ColumnInfo>(Columns);
        }
        else
        {
            var searchLower = ColumnsSearchText.ToLowerInvariant();
            var filtered = Columns.Where(c =>
                c.Name.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ||
                c.Table.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ||
                c.DataType.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ||
                (c.Expression?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (c.Description?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false)
            ).ToList();

            FilteredColumns = new ObservableCollection<ColumnInfo>(filtered);
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
                ApplyRelationshipsFilter();
            }
        }
    }

    public ObservableCollection<RelationshipInfo> FilteredRelationships
    {
        get => _filteredRelationships;
        set
        {
            if (_filteredRelationships != value)
            {
                _filteredRelationships = value;
                OnPropertyChanged();
                (ExportRelationshipsCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
        }
    }

    public string RelationshipsSearchText
    {
        get => _relationshipsSearchText;
        set
        {
            if (_relationshipsSearchText != value)
            {
                _relationshipsSearchText = value;
                OnPropertyChanged();
                ApplyRelationshipsFilter();
            }
        }
    }

    private void ApplyRelationshipsFilter()
    {
        if (string.IsNullOrWhiteSpace(RelationshipsSearchText))
        {
            FilteredRelationships = new ObservableCollection<RelationshipInfo>(Relationships);
        }
        else
        {
            var searchLower = RelationshipsSearchText.ToLowerInvariant();
            var filtered = Relationships.Where(r =>
                r.Name.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ||
                r.FromTable.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ||
                r.FromColumn.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ||
                r.ToTable.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ||
                r.ToColumn.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ||
                r.Cardinality.Contains(searchLower, StringComparison.OrdinalIgnoreCase)
            ).ToList();

            FilteredRelationships = new ObservableCollection<RelationshipInfo>(filtered);
        }
    }

    public bool IsRelationshipsSettingsOpen
    {
        get => _isRelationshipsSettingsOpen;
        set
        {
            if (_isRelationshipsSettingsOpen != value)
            {
                _isRelationshipsSettingsOpen = value;
                OnPropertyChanged();
            }
        }
    }

    // Relationships column visibility properties (main)
    public bool ShowRelNameColumn { get => _showRelNameColumn; set { if (_showRelNameColumn != value) { _showRelNameColumn = value; OnPropertyChanged(); } } }
    public bool ShowRelFromTableColumn { get => _showRelFromTableColumn; set { if (_showRelFromTableColumn != value) { _showRelFromTableColumn = value; OnPropertyChanged(); } } }
    public bool ShowRelFromColumnColumn { get => _showRelFromColumnColumn; set { if (_showRelFromColumnColumn != value) { _showRelFromColumnColumn = value; OnPropertyChanged(); } } }
    public bool ShowRelToTableColumn { get => _showRelToTableColumn; set { if (_showRelToTableColumn != value) { _showRelToTableColumn = value; OnPropertyChanged(); } } }
    public bool ShowRelToColumnColumn { get => _showRelToColumnColumn; set { if (_showRelToColumnColumn != value) { _showRelToColumnColumn = value; OnPropertyChanged(); } } }
    public bool ShowRelCardinalityColumn { get => _showRelCardinalityColumn; set { if (_showRelCardinalityColumn != value) { _showRelCardinalityColumn = value; OnPropertyChanged(); } } }
    public bool ShowRelCrossFilterColumn { get => _showRelCrossFilterColumn; set { if (_showRelCrossFilterColumn != value) { _showRelCrossFilterColumn = value; OnPropertyChanged(); } } }
    public bool ShowRelIsActiveColumn { get => _showRelIsActiveColumn; set { if (_showRelIsActiveColumn != value) { _showRelIsActiveColumn = value; OnPropertyChanged(); } } }
    // Relationships column visibility properties (advanced)
    public bool ShowRelFromCardinalityColumn { get => _showRelFromCardinalityColumn; set { if (_showRelFromCardinalityColumn != value) { _showRelFromCardinalityColumn = value; OnPropertyChanged(); } } }
    public bool ShowRelToCardinalityColumn { get => _showRelToCardinalityColumn; set { if (_showRelToCardinalityColumn != value) { _showRelToCardinalityColumn = value; OnPropertyChanged(); } } }
    public bool ShowRelSecurityFilterColumn { get => _showRelSecurityFilterColumn; set { if (_showRelSecurityFilterColumn != value) { _showRelSecurityFilterColumn = value; OnPropertyChanged(); } } }
    public bool ShowRelJoinOnDateColumn { get => _showRelJoinOnDateColumn; set { if (_showRelJoinOnDateColumn != value) { _showRelJoinOnDateColumn = value; OnPropertyChanged(); } } }
    public bool ShowRelRelyOnRefIntColumn { get => _showRelRelyOnRefIntColumn; set { if (_showRelRelyOnRefIntColumn != value) { _showRelRelyOnRefIntColumn = value; OnPropertyChanged(); } } }
    public bool ShowRelStateColumn { get => _showRelStateColumn; set { if (_showRelStateColumn != value) { _showRelStateColumn = value; OnPropertyChanged(); } } }
    public bool ShowRelModifiedTimeColumn { get => _showRelModifiedTimeColumn; set { if (_showRelModifiedTimeColumn != value) { _showRelModifiedTimeColumn = value; OnPropertyChanged(); } } }

    public ObservableCollection<DependencyInfo> Dependencies
    {
        get => _dependencies;
        set
        {
            if (_dependencies != value)
            {
                _dependencies = value;
                OnPropertyChanged();
                ApplyDependenciesFilter();
            }
        }
    }

    public ObservableCollection<DependencyInfo> FilteredDependencies
    {
        get => _filteredDependencies;
        set
        {
            if (_filteredDependencies != value)
            {
                _filteredDependencies = value;
                OnPropertyChanged();
            }
        }
    }

    public string DependenciesSearchText
    {
        get => _dependenciesSearchText;
        set
        {
            if (_dependenciesSearchText != value)
            {
                _dependenciesSearchText = value;
                OnPropertyChanged();
                ApplyDependenciesFilter();
            }
        }
    }

    private void ApplyDependenciesFilter()
    {
        if (string.IsNullOrWhiteSpace(DependenciesSearchText))
        {
            FilteredDependencies = new ObservableCollection<DependencyInfo>(Dependencies);
            return;
        }

        var searchLower = DependenciesSearchText.ToLowerInvariant();
        var filtered = Dependencies.Where(d =>
            (d.ObjectName?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (d.ObjectType?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (d.DependsOn?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (d.DependencyType?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false)
        ).ToList();

        FilteredDependencies = new ObservableCollection<DependencyInfo>(filtered);
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
                ApplyUnusedObjectsFilter();
            }
        }
    }

    public ObservableCollection<UnusedObjectInfo> FilteredUnusedObjects
    {
        get => _filteredUnusedObjects;
        set
        {
            if (_filteredUnusedObjects != value)
            {
                _filteredUnusedObjects = value;
                OnPropertyChanged();
            }
        }
    }

    public string UnusedObjectsSearchText
    {
        get => _unusedObjectsSearchText;
        set
        {
            if (_unusedObjectsSearchText != value)
            {
                _unusedObjectsSearchText = value;
                OnPropertyChanged();
                ApplyUnusedObjectsFilter();
            }
        }
    }

    private void ApplyUnusedObjectsFilter()
    {
        if (string.IsNullOrWhiteSpace(UnusedObjectsSearchText))
        {
            FilteredUnusedObjects = new ObservableCollection<UnusedObjectInfo>(UnusedObjects);
            return;
        }

        var searchLower = UnusedObjectsSearchText.ToLowerInvariant();
        var filtered = UnusedObjects.Where(u =>
            (u.Name?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (u.ObjectType?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (u.Table?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (u.Reason?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false)
        ).ToList();

        FilteredUnusedObjects = new ObservableCollection<UnusedObjectInfo>(filtered);
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
                ApplyImpactAnalysisFilter();
            }
        }
    }

    public ObservableCollection<ImpactAnalysisInfo> FilteredImpactAnalysis
    {
        get => _filteredImpactAnalysis;
        set
        {
            if (_filteredImpactAnalysis != value)
            {
                _filteredImpactAnalysis = value;
                OnPropertyChanged();
            }
        }
    }

    public string ImpactAnalysisSearchText
    {
        get => _impactAnalysisSearchText;
        set
        {
            if (_impactAnalysisSearchText != value)
            {
                _impactAnalysisSearchText = value;
                OnPropertyChanged();
                ApplyImpactAnalysisFilter();
            }
        }
    }

    private void ApplyImpactAnalysisFilter()
    {
        if (string.IsNullOrWhiteSpace(ImpactAnalysisSearchText))
        {
            FilteredImpactAnalysis = new ObservableCollection<ImpactAnalysisInfo>(ImpactAnalysis);
            return;
        }

        var searchLower = ImpactAnalysisSearchText.ToLowerInvariant();
        var filtered = ImpactAnalysis.Where(i =>
            (i.ObjectName?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (i.ObjectType?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (i.ImpactedObject?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (i.ImpactType?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false) ||
            (i.Severity?.Contains(searchLower, StringComparison.OrdinalIgnoreCase) ?? false)
        ).ToList();

        FilteredImpactAnalysis = new ObservableCollection<ImpactAnalysisInfo>(filtered);
    }

    public bool IsImpactAnalysisSettingsOpen
    {
        get => _isImpactAnalysisSettingsOpen;
        set
        {
            if (_isImpactAnalysisSettingsOpen != value)
            {
                _isImpactAnalysisSettingsOpen = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowIAObjectColumn { get => _showIAObjectColumn; set { if (_showIAObjectColumn != value) { _showIAObjectColumn = value; OnPropertyChanged(); } } }
    public bool ShowIAObjectTypeColumn { get => _showIAObjectTypeColumn; set { if (_showIAObjectTypeColumn != value) { _showIAObjectTypeColumn = value; OnPropertyChanged(); } } }
    public bool ShowIAImpactedObjectColumn { get => _showIAImpactedObjectColumn; set { if (_showIAImpactedObjectColumn != value) { _showIAImpactedObjectColumn = value; OnPropertyChanged(); } } }
    public bool ShowIAImpactTypeColumn { get => _showIAImpactTypeColumn; set { if (_showIAImpactTypeColumn != value) { _showIAImpactTypeColumn = value; OnPropertyChanged(); } } }
    public bool ShowIASeverityColumn { get => _showIASeverityColumn; set { if (_showIASeverityColumn != value) { _showIASeverityColumn = value; OnPropertyChanged(); } } }

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

    // Columns Tab Column Visibility Properties
    public bool IsColumnsSettingsOpen
    {
        get => _isColumnsSettingsOpen;
        set
        {
            if (_isColumnsSettingsOpen != value)
            {
                _isColumnsSettingsOpen = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ShowColNameColumn
    {
        get => _showColNameColumn;
        set { if (_showColNameColumn != value) { _showColNameColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColTableColumn
    {
        get => _showColTableColumn;
        set { if (_showColTableColumn != value) { _showColTableColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColDataTypeColumn
    {
        get => _showColDataTypeColumn;
        set { if (_showColDataTypeColumn != value) { _showColDataTypeColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColIsCalculatedColumn
    {
        get => _showColIsCalculatedColumn;
        set { if (_showColIsCalculatedColumn != value) { _showColIsCalculatedColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColExpressionColumn
    {
        get => _showColExpressionColumn;
        set { if (_showColExpressionColumn != value) { _showColExpressionColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColIsHiddenColumn
    {
        get => _showColIsHiddenColumn;
        set { if (_showColIsHiddenColumn != value) { _showColIsHiddenColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColDescriptionColumn
    {
        get => _showColDescriptionColumn;
        set { if (_showColDescriptionColumn != value) { _showColDescriptionColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColDisplayFolderColumn
    {
        get => _showColDisplayFolderColumn;
        set { if (_showColDisplayFolderColumn != value) { _showColDisplayFolderColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColFormatStringColumn
    {
        get => _showColFormatStringColumn;
        set { if (_showColFormatStringColumn != value) { _showColFormatStringColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColSortByColumnColumn
    {
        get => _showColSortByColumnColumn;
        set { if (_showColSortByColumnColumn != value) { _showColSortByColumnColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColIsUniqueColumn
    {
        get => _showColIsUniqueColumn;
        set { if (_showColIsUniqueColumn != value) { _showColIsUniqueColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColIsNullableColumn
    {
        get => _showColIsNullableColumn;
        set { if (_showColIsNullableColumn != value) { _showColIsNullableColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColIsKeyColumn
    {
        get => _showColIsKeyColumn;
        set { if (_showColIsKeyColumn != value) { _showColIsKeyColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColSourceColumnColumn
    {
        get => _showColSourceColumnColumn;
        set { if (_showColSourceColumnColumn != value) { _showColSourceColumnColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColDataCategoryColumn
    {
        get => _showColDataCategoryColumn;
        set { if (_showColDataCategoryColumn != value) { _showColDataCategoryColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColIsAvailableInMDXColumn
    {
        get => _showColIsAvailableInMDXColumn;
        set { if (_showColIsAvailableInMDXColumn != value) { _showColIsAvailableInMDXColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColStateColumn
    {
        get => _showColStateColumn;
        set { if (_showColStateColumn != value) { _showColStateColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColErrorMessageColumn
    {
        get => _showColErrorMessageColumn;
        set { if (_showColErrorMessageColumn != value) { _showColErrorMessageColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColModifiedTimeColumn
    {
        get => _showColModifiedTimeColumn;
        set { if (_showColModifiedTimeColumn != value) { _showColModifiedTimeColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColLineageTagColumn
    {
        get => _showColLineageTagColumn;
        set { if (_showColLineageTagColumn != value) { _showColLineageTagColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColSummarizeByColumn
    {
        get => _showColSummarizeByColumn;
        set { if (_showColSummarizeByColumn != value) { _showColSummarizeByColumn = value; OnPropertyChanged(); } }
    }

    public bool ShowColEncodingColumn
    {
        get => _showColEncodingColumn;
        set { if (_showColEncodingColumn != value) { _showColEncodingColumn = value; OnPropertyChanged(); } }
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
    
    // Columns tab commands
    public ICommand ClearColumnsSearchCommand { get; }
    public ICommand ExportColumnsCommand { get; }
    public ICommand ViewColumnExpressionCommand { get; }
    public ICommand CopyColumnExpressionCommand { get; }
    
    // Relationships tab commands
    public ICommand ClearRelationshipsSearchCommand { get; }
    public ICommand ExportRelationshipsCommand { get; }

    // Dependencies tab commands
    public ICommand ClearDependenciesSearchCommand { get; }

    // Unused Objects tab commands
    public ICommand ClearUnusedObjectsSearchCommand { get; }

    // Impact Analysis tab commands
    public ICommand ClearImpactAnalysisSearchCommand { get; }

    // Best Practices tab commands
    public ICommand ClearBestPracticesSearchCommand { get; }
    public ICommand ExpandAllBPRulesCommand { get; }
    public ICommand CollapseAllBPRulesCommand { get; }

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
        
        // Columns tab commands
        ClearColumnsSearchCommand = new RelayCommand(async () => await Task.Run(() => ColumnsSearchText = string.Empty), () => true);
        ExportColumnsCommand = new RelayCommand(async () => await ExportColumnsToExcelAsync(), () => IsConnected && FilteredColumns.Count > 0);
        ViewColumnExpressionCommand = new RelayCommandWithParameter(ViewColumnExpression);
        CopyColumnExpressionCommand = new RelayCommandWithParameter(CopyColumnExpression);
        
        // Relationships tab commands
        ClearRelationshipsSearchCommand = new RelayCommand(async () => await Task.Run(() => RelationshipsSearchText = string.Empty), () => true);
        ExportRelationshipsCommand = new RelayCommand(async () => await ExportRelationshipsToExcelAsync(), () => IsConnected && FilteredRelationships.Count > 0);

        // Dependencies tab commands
        ClearDependenciesSearchCommand = new RelayCommand(async () => await Task.Run(() => DependenciesSearchText = string.Empty), () => true);

        // Unused Objects tab commands
        ClearUnusedObjectsSearchCommand = new RelayCommand(async () => await Task.Run(() => UnusedObjectsSearchText = string.Empty), () => true);

        // Impact Analysis tab commands
        ClearImpactAnalysisSearchCommand = new RelayCommand(async () => await Task.Run(() => ImpactAnalysisSearchText = string.Empty), () => true);

        // Best Practices tab commands
        ClearBestPracticesSearchCommand = new RelayCommand(async () => await Task.Run(() => BestPracticesSearchText = string.Empty), () => true);
        ExpandAllBPRulesCommand = new RelayCommand(async () => await Task.Run(() => ExpandAllBPRules()), () => true);
        CollapseAllBPRulesCommand = new RelayCommand(async () => await Task.Run(() => CollapseAllBPRules()), () => true);
    }

    private void ExpandAllBPRules()
    {
        foreach (var group in BestPracticeRuleGroups)
        {
            group.IsExpanded = true;
        }
    }

    private void CollapseAllBPRules()
    {
        foreach (var group in BestPracticeRuleGroups)
        {
            group.IsExpanded = false;
        }
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
                StatusMessage = "DAX expression copied to clipboard!";
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
            StatusMessage = "DAX expression formatted!";
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

    private void ViewColumnExpression(object? parameter)
    {
        if (parameter is ColumnInfo column && !string.IsNullOrEmpty(column.Expression))
        {
            SelectedMeasureName = $"{column.Table}[{column.Name}]";
            SelectedDaxExpression = column.Expression;
            IsDaxViewerOpen = true;
        }
    }

    private void CopyColumnExpression(object? parameter)
    {
        string? expression = null;
        
        if (parameter is ColumnInfo column)
        {
            expression = column.Expression;
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
                StatusMessage = "Column expression copied to clipboard!";
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Copy error: {ex}");
                StatusMessage = "Failed to copy to clipboard.";
            }
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

            // Load unused objects from the model
            StatusMessage = $"Analyzing unused objects in {instance.DisplayName}...";
            var unusedObjects = await _powerBIService.GetUnusedObjectsAsync(instance.Port, cancellationToken);
            UnusedObjects = new ObservableCollection<UnusedObjectInfo>(unusedObjects);

            // Load Best Practice Analyzer results (BPA) - Using local BPARules.json file
            StatusMessage = $"Loading Best Practice rules...";
            try
            {
                Debug.WriteLine("BPA: Loading rules from local BPARules.json file");
                var rules = await LocalBestPracticeRulesLoader.LoadRulesAsync();
                Debug.WriteLine($"BPA: Loaded {rules.Count} rules from local file");

                TotalRules = rules.Count; // Set total rules count

                if (rules.Count == 0)
                {
                    Debug.WriteLine("BPA: Warning - No rules loaded from file");
                    BestPracticeViolations = new ObservableCollection<BestPracticeViolation>();
                }
                else
                {
                    StatusMessage = $"Running Best Practice Analyzer with {rules.Count} rules...";
                    var bpa = new BestPracticeAnalyzer(rules);
                    var violations = await Task.Run(() =>
                    {
                        using var server = new Server();
                        server.Connect($"DataSource=localhost:{instance.Port}");
                        var database = server.Databases[0];
                        var model = database.Model;
                        Debug.WriteLine($"BPA: Analyzing model '{model.Name}' with {model.Tables.Count} tables");
                        return bpa.AnalyzeModel(model);
                    }, cancellationToken);

                    BestPracticeViolations = new ObservableCollection<BestPracticeViolation>(violations);
                    ApplyBestPracticesFilter();

                    Debug.WriteLine($"BPA: Analysis complete. Found {BestPracticeViolations.Count} violations");
                    if (BestPracticeViolations.Count == 0)
                    {
                        Debug.WriteLine("BPA: No violations found - model follows all best practices!");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"BPA: Error during analysis: {ex.Message}");
                Debug.WriteLine($"BPA: Stack trace: {ex.StackTrace}");
                BestPracticeViolations = new ObservableCollection<BestPracticeViolation>();
            }
            
            // Load other data (dependencies, unused objects, impact analysis - these require more complex analysis)
            LoadAnalysisData();
            
            StatusMessage = $"Connected to {instance.DisplayName}! Loaded {ModelOverview.MeasureCount} measures, {ModelOverview.ColumnCount} columns, {ModelOverview.RelationshipCount} relationships";
            IsConnected = true;
            _connectedPort = instance.Port;
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

        // Unused Objects are loaded from the service (model-driven)

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

    private async Task ExportColumnsToExcelAsync()
    {
        try
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = $"Columns_Export_{DateTime.Now:yyyyMMdd_HHmmss}"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                await Task.Run(() =>
                {
                    using var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add("Columns");

                    int col = 1;

                    // S.No column (always visible)
                    worksheet.Cell(1, col).Value = "S.No";
                    col++;

                    // Column mappings with visibility
                    var columnMappings = new List<(string Header, Func<ColumnInfo, object?> ValueGetter, bool IsVisible)>
                    {
                        ("Name", c => c.Name, ShowColNameColumn),
                        ("Table", c => c.Table, ShowColTableColumn),
                        ("Data Type", c => c.DataType, ShowColDataTypeColumn),
                        ("Is Calculated", c => c.IsCalculated ? "Yes" : "No", ShowColIsCalculatedColumn),
                        ("Expression", c => c.Expression, ShowColExpressionColumn),
                        ("Is Hidden", c => c.IsHidden ? "Hidden" : "Visible", ShowColIsHiddenColumn),
                        ("Description", c => c.Description, ShowColDescriptionColumn),
                        ("Display Folder", c => c.DisplayFolder, ShowColDisplayFolderColumn),
                        ("Format String", c => c.FormatString, ShowColFormatStringColumn),
                        ("Sort By Column", c => c.SortByColumn, ShowColSortByColumnColumn),
                        ("Is Unique", c => c.IsUnique ? "Yes" : "No", ShowColIsUniqueColumn),
                        ("Is Nullable", c => c.IsNullable ? "Yes" : "No", ShowColIsNullableColumn),
                        ("Is Key", c => c.IsKey ? "Yes" : "No", ShowColIsKeyColumn),
                        ("Source Column", c => c.SourceColumn, ShowColSourceColumnColumn),
                        ("Data Category", c => c.DataCategory, ShowColDataCategoryColumn),
                        ("Available In MDX", c => c.IsAvailableInMDX ? "Yes" : "No", ShowColIsAvailableInMDXColumn),
                        ("State", c => c.State, ShowColStateColumn),
                        ("Error Message", c => c.ErrorMessage, ShowColErrorMessageColumn),
                        ("Modified Time", c => c.ModifiedTime?.ToString("g"), ShowColModifiedTimeColumn),
                        ("Lineage Tag", c => c.LineageTag, ShowColLineageTagColumn),
                        ("Summarize By", c => c.SummarizeBy, ShowColSummarizeByColumn),
                        ("Encoding", c => c.Encoding, ShowColEncodingColumn)
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
                    headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#5C2D91");
                    headerRange.Style.Font.FontColor = XLColor.White;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Add data rows
                    int row = 2;
                    int serialNumber = 1;
                    foreach (var columnInfo in FilteredColumns)
                    {
                        col = 1;

                        // S.No
                        worksheet.Cell(row, col).Value = serialNumber++;
                        col++;

                        // Add visible column values
                        foreach (var column in visibleColumns)
                        {
                            var value = column.ValueGetter(columnInfo);
                            if (value != null)
                            {
                                worksheet.Cell(row, col).Value = value.ToString();
                            }
                            col++;
                        }
                        row++;
                    }

                    // Auto-fit columns with max width
                    worksheet.Columns().AdjustToContents();
                    foreach (var column in worksheet.ColumnsUsed())
                    {
                        if (column.Width > 50)
                        {
                            column.Width = 50;
                        }
                    }

                    worksheet.RangeUsed()?.SetAutoFilter();
                    workbook.SaveAs(saveFileDialog.FileName);
                });

                MessageBox.Show(
                    $"Columns exported successfully to:\n{saveFileDialog.FileName}",
                    "Export Complete",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                Process.Start("explorer.exe", $"/select,\"{saveFileDialog.FileName}\"");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Failed to export columns.\n\nError: {ex.Message}",
                "Export Failed",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Debug.WriteLine($"Export error: {ex}");
        }
    }

    private async Task ExportRelationshipsToExcelAsync()
    {
        try
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = $"Relationships_Export_{DateTime.Now:yyyyMMdd_HHmmss}"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                await Task.Run(() =>
                {
                    using var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add("Relationships");

                    int col = 1;

                    // S.No column (always visible)
                    worksheet.Cell(1, col).Value = "S.No";
                    col++;

                    // Column mappings with visibility
                    var columnMappings = new List<(string Header, Func<RelationshipInfo, object?> ValueGetter, bool IsVisible)>
                    {
                        ("Name", r => r.Name, ShowRelNameColumn),
                        ("From Table", r => r.FromTable, ShowRelFromTableColumn),
                        ("From Column", r => r.FromColumn, ShowRelFromColumnColumn),
                        ("To Table", r => r.ToTable, ShowRelToTableColumn),
                        ("To Column", r => r.ToColumn, ShowRelToColumnColumn),
                        ("Cardinality", r => r.Cardinality, ShowRelCardinalityColumn),
                        ("Cross Filter", r => r.CrossFilterDirection, ShowRelCrossFilterColumn),
                        ("Is Active", r => r.IsActive ? "Yes" : "No", ShowRelIsActiveColumn),
                        ("From Cardinality", r => r.FromCardinality, ShowRelFromCardinalityColumn),
                        ("To Cardinality", r => r.ToCardinality, ShowRelToCardinalityColumn),
                        ("Security Filter", r => r.SecurityFilteringBehavior, ShowRelSecurityFilterColumn),
                        ("Join On Date", r => r.JoinOnDateBehavior, ShowRelJoinOnDateColumn),
                        ("Rely On Ref. Integrity", r => r.RelyOnReferentialIntegrity ? "Yes" : "No", ShowRelRelyOnRefIntColumn),
                        ("State", r => r.State, ShowRelStateColumn),
                        ("Modified Time", r => r.ModifiedTime?.ToString("g"), ShowRelModifiedTimeColumn)
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
                    headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#E81123");
                    headerRange.Style.Font.FontColor = XLColor.White;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Add data rows
                    int row = 2;
                    int serialNumber = 1;
                    foreach (var relationship in FilteredRelationships)
                    {
                        col = 1;

                        // S.No
                        worksheet.Cell(row, col).Value = serialNumber++;
                        col++;

                        // Add visible column values
                        foreach (var column in visibleColumns)
                        {
                            var value = column.ValueGetter(relationship);
                            if (value != null)
                            {
                                worksheet.Cell(row, col).Value = value.ToString();
                            }
                            col++;
                        }
                        row++;
                    }

                    // Auto-fit columns with max width
                    worksheet.Columns().AdjustToContents();
                    foreach (var column in worksheet.ColumnsUsed())
                    {
                        if (column.Width > 50)
                        {
                            column.Width = 50;
                        }
                    }

                    worksheet.RangeUsed()?.SetAutoFilter();
                    workbook.SaveAs(saveFileDialog.FileName);
                });

                MessageBox.Show(
                    $"Relationships exported successfully to:\n{saveFileDialog.FileName}",
                    "Export Complete",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                Process.Start("explorer.exe", $"/select,\"{saveFileDialog.FileName}\"");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Failed to export relationships.\n\nError: {ex.Message}",
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

