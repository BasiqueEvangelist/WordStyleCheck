namespace WordStyleCheck;

public static class LintMerger
{
    public static void Run(List<LintMessage> messages)
    {
        // TODO: make this not rely on lints being order in execution order.
        for (int i = 1; i < messages.Count; i++)
        {
            if (messages[i - 1].Id != messages[i].Id) continue;
            // TODO: check for equal parameters, actually
            //if (messages[i - 1].Parameters != messages[i].Parameters) continue;

            var newContext = messages[i].Context.TryMerge(messages[i - 1].Context);

            if (newContext == null) continue;
            
            var prevAutofix = messages[i - 1].AutoFix;
            var curAutofix = messages[i].AutoFix;
            
            messages[i - 1] = new LintMessage(
                messages[i].Id,
                newContext,
                messages[i].Parameters,
                prevAutofix != null || curAutofix != null 
                    ? () => { prevAutofix?.Invoke(); curAutofix?.Invoke(); }
                    : null
            );
            messages.RemoveAt(i);
            i -= 1;
        }
    }
}