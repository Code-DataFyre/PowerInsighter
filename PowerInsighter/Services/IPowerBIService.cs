using PowerInsighter.Models;

namespace PowerInsighter.Services;

public interface IPowerBIService
{
    bool IsPowerBIRunning();
    Task<List<int>> FindPowerBIPortsAsync(CancellationToken cancellationToken = default);
    Task<List<ModelMetadata>> LoadMetadataAsync(int port, CancellationToken cancellationToken = default);
}
