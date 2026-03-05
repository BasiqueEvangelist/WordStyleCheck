using System.Collections.Concurrent;

namespace WordStyleCheck;

public class LinterThreadPool : IDisposable
{
    private readonly BlockingCollection<LintTask> _tasks = new();
    
    public LinterThreadPool(int numThreads)
    {
        Thread[] threads = new Thread[numThreads];

        for (int i = 0; i < numThreads; i++)
        {
            threads[i] = new Thread(RunLoop);
            threads[i].Name = "Linter Pool Thread #" + (i + 1);
            threads[i].IsBackground = true;
            threads[i].Start();
        }
    }

    public void AddTask(LintTask task)
    {
        _tasks.Add(task);
    }
    
    private void RunLoop()
    {
        while (true)
        {
            LintTask task;
            try
            {
                task = _tasks.Take();
            }
            catch (InvalidOperationException)
            {
                return;
            }
            
            var linter = new DocumentLinter(task.Path, task.TakeOwnership);
            linter.LintIdFilter = task.LintIdFilter;
            linter.RunLints();
            
            task._resultSrc.SetResult(linter);
        }
    }

    public void Dispose()
    {
        _tasks.CompleteAdding();
    }
}

public class LintTask(string path, Predicate<string> lintIdFilter, bool takeOwnership)
{
    public string Path { get; } = path;
    public bool TakeOwnership { get; } = takeOwnership;
    public Predicate<string> LintIdFilter { get; } = lintIdFilter;

    internal TaskCompletionSource<DocumentLinter> _resultSrc = new();

    public Task<DocumentLinter> Result => _resultSrc.Task;
}
