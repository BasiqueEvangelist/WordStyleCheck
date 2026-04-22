using System.Security.Cryptography;
using System.Text.Json;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Minio;
using Minio.DataModel.Args;
using WordStyleCheck;
using WordStyleCheck.Gost7_32;

namespace WordStyleCheckService.Worker;

public class WorkerService(ILogger<WorkerService> logger, Db db, IOptionsMonitor<Options> options, IMinioClient s3)
    : BackgroundService
{
    private static readonly XmlTranslationsFile _translations = XmlTranslationsFile.LoadEmbedded();

    private readonly LinterThreadPool _pool = new(options.CurrentValue.PoolSize);

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
                    Code = "exception",
                    Message = "Task failed with an exception",
                    Exception = e.ToString()
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

        MemoryStream stream = new MemoryStream();
        await s3.GetObjectAsync(new GetObjectArgs()
            .WithBucket(options.CurrentValue.S3IngressBucket)
            .WithObject(inputs.InputObject)
            .WithCallbackStream((x, token) => x.CopyToAsync(stream, token)));
        
        stream.Seek(0, SeekOrigin.Begin);

        LintTask task = new LintTask(stream, new Gost7_32Profile(), x => true, false, _translations);
        _pool.AddTask(task);
        var linter = await task.Result;

        var file = linter.Save();

        string objectKey = RandomNumberGenerator.GetHexString(32) + ".docx";
        await s3.PutObjectAsync(new PutObjectArgs()
            .WithBucket(options.CurrentValue.S3EgressBucket)
            .WithObject(objectKey)
            .WithStreamData(file)
            .WithObjectSize(file.Length)
            .WithContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
        
        TaskOutputs outputs = new TaskOutputs()
        {
            OutputObject = objectKey
        };

        return JsonSerializer.Serialize(outputs);
    } 
}