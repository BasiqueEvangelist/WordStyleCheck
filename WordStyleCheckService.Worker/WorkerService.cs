using System.Text.Json;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using WordStyleCheck;

namespace WordStyleCheckService.Worker;

public class WorkerService(ILogger<WorkerService> logger, Db db, IOptions<Options> options)
    : BackgroundService
{
    private static readonly XmlTranslationsFile _translations = XmlTranslationsFile.LoadEmbedded();

    private readonly LinterThreadPool _pool = new(options.Value.PoolSize);

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            var taskInfo = await db.TryDequeuePendingTask();
            if (taskInfo is null)
            {
                await Task.Delay(1000, stoppingToken);
                continue;
            }
        
            logger.LogInformation("Processing task {TaskId} as {TaskType}", taskInfo.TaskId, taskInfo.TaskType);
            try
            {
                var result = await RunTask(taskInfo.TaskId, taskInfo.TaskType, taskInfo.TaskData);
                await db.ReportTaskSuccess(taskInfo.PendingTaskId, result);
                logger.LogInformation("Reported task {TaskId} success", taskInfo.TaskId);
            }
            catch (Exception e)
            {
                logger.LogError("Processing task {TaskId} failed", taskInfo.TaskId);
                var error = new {
                    code = "exception",
                    message = "Task failed with an exception",
                    exception = e.ToString()
                };
                await db.ReportTaskFailure(taskInfo.PendingTaskId, JsonSerializer.Serialize(error));
                logger.LogInformation("Reported task {TaskId} failure", taskInfo.TaskId);
            }
        }
    }

    private async Task<string> RunTask(uint taskId, string taskType, string taskData)
    {
        if (taskType != "LintFile")
        {
            throw new NotImplementedException();
        }
        
        TaskInputs inputs = JsonSerializer.Deserialize<TaskInputs>(taskData)!;

        LintTask task = new LintTask(inputs.DownloadUrl, x => true, false, _translations);
        _pool.AddTask(task);
        var linter = await task.Result;

        var file = linter.SaveTemp();

        TaskOutputs outputs = new TaskOutputs()
        {
            DownloadUrl = file
        };

        return JsonSerializer.Serialize(outputs);
    } 
}