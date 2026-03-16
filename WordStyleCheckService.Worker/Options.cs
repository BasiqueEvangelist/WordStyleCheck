namespace WordStyleCheckService.Worker;

public class Options
{
    public required string ConnectionString { get; set; }
    
    public required string S3AccessKeyId { get; set; }
    public required string S3SecretAccessKey { get; set; }
    public required string S3Region { get; set; }
    public required string S3EndpointUrl { get; set; }
    public required string S3IngressBucket { get; set; }
    public required string S3EgressBucket { get; set; }
    public required string S3EgressBucketBaseUrl { get; set; }

    public int PoolSize { get; set; } = Environment.ProcessorCount;
}