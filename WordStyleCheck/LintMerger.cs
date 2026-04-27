namespace WordStyleCheck;

public static class LintMerger
{
    public static void Run(List<LintDiagnostic> messages)
    {
        // TODO: make this not rely on lints being order in execution order.
        for (int i = 1; i < messages.Count; i++)
        {
            if (messages[i - 1].Id != messages[i].Id) continue;
            if (messages[i - 1].Type != messages[i].Type) continue;
            // TODO: make this more efficient.
            if ((messages[i - 1].Parameters == null) != (messages[i].Parameters == null)) continue;
            if (messages[i - 1].Parameters != null && messages[i - 1].Parameters!.Except(messages[i].Parameters!).Any()) continue;

            var newContext = messages[i].Context.TryMerge(messages[i - 1].Context);

            if (newContext == null) continue;
            
            messages[i - 1] = new LintDiagnostic(
                messages[i].Id,
                messages[i].Type,
                newContext,
                messages[i].Parameters
            );
            messages.RemoveAt(i);
            i -= 1;
        }
    }
}