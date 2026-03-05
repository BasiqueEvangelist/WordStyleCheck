using WordStyleCheck;

namespace WordStyleCheckWeb;

public class DocumentProcessingService : BackgroundService
{
    private readonly Dictionary<Guid, DocumentTask> _tasks = new();
    private readonly LinterThreadPool _pool = new(Environment.ProcessorCount);
    private readonly DiagnosticTranslationsFile _translations = DiagnosticTranslationsFile.LoadEmbedded();
    
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
            
            _pool.Dispose();
        }
    }

    public class DocumentTask : IDisposable
    {
        private DocumentLinter? _linter;
        private DateTime? finishedAt;
        private Task finished;
        
        public DocumentTask(string name, string tempPath, LinterThreadPool pool, DiagnosticTranslationsFile translations)
        {
            Name = name;

            async Task RunThing()
            {
                var task = new LintTask(tempPath, _ => true, true, translations);
                pool.AddTask(task);
                _linter = await task.Result;
                finishedAt = DateTime.Now;
            }

            finished = RunThing();
        }

        public Guid Id { get; } = Guid.NewGuid();
        
        public string Name { get; }

        public int? StyleErrorCount => finishedAt != null ? _linter!.Diagnostics.Count : null;

        public Task Finished => finished;

        public bool IsExpired => finishedAt != null && (DateTime.Now - finishedAt.Value).TotalMinutes > 5;
        
        public void Dispose()
        {
            _linter?.Dispose();
        }

        public string GetPath()
        {
            return _linter!.SaveTemp();
        }
    }

    public DocumentTask? GetTask(Guid id)
    {
        return _tasks.GetValueOrDefault(id);
    }
    
    public DocumentTask StartTask(string name, string tempPath)
    {
        var task = new DocumentTask(name, tempPath, _pool, _translations);
        
        _tasks.Add(task.Id, task);

        return task;
    }
}