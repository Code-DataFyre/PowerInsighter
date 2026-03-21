using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace PowerInsighter.Models;

/// <summary>
/// Represents a grouped view of best practice violations by rule
/// </summary>
public class BestPracticeRuleGroup : INotifyPropertyChanged
{
    private bool _isExpanded;

    public string RuleName { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Severity { get; set; } = string.Empty;
    public int SeverityLevel { get; set; }
    public int ViolationCount { get; set; }
    public int ViolationScore { get; set; }
    
    public ObservableCollection<BestPracticeViolation> Violations { get; set; } = new();

    public bool IsExpanded
    {
        get => _isExpanded;
        set
        {
            if (_isExpanded != value)
            {
                _isExpanded = value;
                OnPropertyChanged();
            }
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
