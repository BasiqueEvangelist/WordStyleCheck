using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Context;

namespace WordStyleCheck;

public static class RunLintMerger
{
    public static void Run(List<LintMessage> messages)
    {
        // TODO: make this not rely on lints being order in execution order.
        for (int i = 1; i < messages.Count; i++)
        {
            if (!(messages[i - 1].Context is RunDiagnosticContext prev &&
                  messages[i].Context is RunDiagnosticContext next))
            {
                continue;
            }

            if (messages[i - 1].Message != messages[i].Message) continue;
            if (messages[i - 1].AutoFixed != messages[i].AutoFixed) continue;
            if (messages[i - 1].Values != messages[i].Values) continue;
            
            if (prev.Runs[^1].NextSibling() != next.Runs[0]) continue;

            RunDiagnosticContext newContext = new RunDiagnosticContext([..prev.Runs, ..next.Runs]);

            messages[i - 1] = new LintMessage(messages[i].Message, messages[i].Values, messages[i].AutoFixed, newContext);
            messages.RemoveAt(i);
            i -= 1;
        }
    }
}