using WordStyleCheck;

namespace WordStyleCheckWeb;

public class DocumentProcessingService : BackgroundService
{
    private readonly Dictionary<Guid, DocumentTask> _tasks = new();
    
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        try
        {
            while (true)
            {
                await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken);

                foreach (var taskId in _tasks.Where(x => x.Value.IsExpired).Select(x => x.Key).ToList())
                {
                    _tasks[taskId].Dispose();
                    _tasks.Remove(taskId);
                }
            }
        }
        catch (TaskCanceledException)
        {
            foreach (var task in _tasks.Values)
            {
                task?.Dispose();
            }    
        }
    }

    public class DocumentTask : IDisposable
    {
        private DocumentLinter? _linter;
        private DateTime? finishedAt;
        private TaskCompletionSource finishedSrc;
        
        public DocumentTask(string name, string tempPath)
        {
            Name = name;
            finishedSrc = new TaskCompletionSource();
            
            var thread = new Thread(() =>
            {
                _linter = new DocumentLinter(tempPath, true);
                _linter.RunLints();
                
                var translations = DiagnosticTranslationsFile.LoadEmbedded();

                foreach (var message in _linter!.Diagnostics)
                {
                    _linter.DocumentAnalysis.WriteComment(message, translations);
                }

                finishedAt = DateTime.Now;
                finishedSrc.SetResult();
            })
            {
                Name = "Linter thread for " + name
            };

            thread.Start();
        }

        public Guid Id { get; } = Guid.NewGuid();
        
        public string Name { get; }

        public int? StyleErrorCount => finishedAt != null ? _linter!.Diagnostics.Count : null;

        public Task Finished => finishedSrc.Task;

        public bool IsExpired => finishedAt != null && (DateTime.Now - finishedAt.Value).TotalMinutes > 5;
        
        public void Dispose()
        {
            _linter?.Dispose();
        }

        public string GetPath()
        {
            return _linter.SaveTemp();
        }
    }

    public DocumentTask? GetTask(Guid id)
    {
        return _tasks.GetValueOrDefault(id);
    }
    
    public DocumentTask StartTask(string name, string tempPath)
    {
        var task = new DocumentTask(name, tempPath);
        
        _tasks.Add(task.Id, task);

        return task;
    }
}