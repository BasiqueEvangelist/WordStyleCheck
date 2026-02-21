using System.Diagnostics;

namespace WordStyleCheck;

public readonly struct LoudStopwatch(string name) : IDisposable
{
    public static bool Enabled = false;
    
    private readonly long _timestamp = Stopwatch.GetTimestamp();

    public void Dispose()
    {
        if (!Enabled) return;
        
        Console.WriteLine($"{name} took {Stopwatch.GetElapsedTime(_timestamp, Stopwatch.GetTimestamp()).TotalMilliseconds}ms");
    }
}