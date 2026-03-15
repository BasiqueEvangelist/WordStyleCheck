using System.Text.Json;
using WordStyleCheckService.Worker;

namespace WordStyleCheckService.Foreman;

public class CleanupDbService(Db db) : BackgroundService
{
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        var errorExpiredNotTaken = JsonSerializer.Serialize(new {
            Code = "expiredNotTaken",
            Message = "Task was not taken within expiry period"
        });
        
        var errorProcessingTimeout = JsonSerializer.Serialize(new {
            Code = "processingTimeout",
            Message = "Task was not processed within timeout"
        });
        
        while (!stoppingToken.IsCancellationRequested)
        {
            await db.MaintainTaskQueue(errorExpiredNotTaken, errorProcessingTimeout);
            
            await Task.Delay(TimeSpan.FromSeconds(30), stoppingToken);
        }
    }
}