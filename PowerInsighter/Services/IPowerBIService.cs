using PowerInsighter.Models;

namespace PowerInsighter.Services;

public interface IPowerBIService
{
    bool IsPowerBIRunning();
    Task<List<int>> FindPowerBIPortsAsync(CancellationToken cancellationToken = default);
    Task<List<PowerBIInstance>> FindPowerBIInstancesAsync(CancellationToken cancellationToken = default);
    Task<List<ModelMetadata>> LoadMetadataAsync(int port, CancellationToken cancellationToken = default);
    Task<ModelOverview> GetModelOverviewAsync(int port, CancellationToken cancellationToken = default);
    Task<List<MeasureInfo>> GetMeasuresAsync(int port, CancellationToken cancellationToken = default);
    Task<List<ColumnInfo>> GetColumnsAsync(int port, CancellationToken cancellationToken = default);
    Task<List<RelationshipInfo>> GetRelationshipsAsync(int port, CancellationToken cancellationToken = default);
    Task<List<UnusedObjectInfo>> GetUnusedObjectsAsync(int port, CancellationToken cancellationToken = default);
}
