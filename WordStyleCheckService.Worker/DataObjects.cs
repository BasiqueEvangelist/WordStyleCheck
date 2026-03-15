namespace WordStyleCheckService.Worker;

public class TaskInputs
{
    public required string InputObject { get; set; }
}

public class TaskOutputs
{
    public string Code { get; set; } = "success";
    
    public string? OutputObject { get; set; }
}