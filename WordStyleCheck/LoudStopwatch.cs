using System.Diagnostics;

namespace WordStyleCheck;

public readonly struct LoudStopwatch(string name) : IDisposable
{
    private readonly long _timestamp = Stopwatch.GetTimestamp();

    public void Dispose()
    {
        Console.WriteLine($"{name} took {Stopwatch.GetElapsedTime(_timestamp, Stopwatch.GetTimestamp()).TotalMilliseconds}ms");
    }
}