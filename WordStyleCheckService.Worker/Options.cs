namespace WordStyleCheckService.Worker;

public class Options
{
    public required string ConnectionString { get; set; }

    public int PoolSize { get; set; } = Environment.ProcessorCount;
}