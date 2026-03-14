namespace WordStyleCheckService.Worker;

public class TaskInputs
{
    public required string DownloadUrl { get; set; }
}

public class TaskOutputs
{
    public string Code { get; set; } = "success";
    
    public required string DownloadUrl { get; set; }
}