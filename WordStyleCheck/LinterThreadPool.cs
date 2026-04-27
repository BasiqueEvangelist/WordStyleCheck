using System.Collections.Concurrent;
using WordStyleCheck.Profiles;

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

            try
            {
                var linter = new DocumentLinter(task.Stream, task.Profile);
                linter.LintIdFilter = task.LintIdFilter;
                linter.RunLints(task.AutoFix);

                if (task.CommentsTranslations != null)
                {
                    linter.ApplyDiagnostics(task.CommentsTranslations);
                }

                task._resultSrc.SetResult(linter);
            }
            catch (Exception e)
            {
                task._resultSrc.SetException(e);
            }
        }
    }

    public void Dispose()
    {
        _tasks.CompleteAdding();
    }
}

public class LintTask(
    Stream stream,
    IProfile profile,
    Predicate<string> lintIdFilter,
    XmlTranslationsFile? commentsTranslations,
    bool autoFix = false)
{
    public Stream Stream { get; } = stream;
    public IProfile Profile { get; } = profile;
    public bool AutoFix { get; } = autoFix;
    public Predicate<string> LintIdFilter { get; } = lintIdFilter;
    public XmlTranslationsFile? CommentsTranslations { get; } = commentsTranslations;

    internal TaskCompletionSource<DocumentLinter> _resultSrc = new();

    public Task<DocumentLinter> Result => _resultSrc.Task;
}
