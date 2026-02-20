using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck;

public static class RunLintMerger
{
    public static void Run(List<LintMessage> messages)
    {
        // TODO: make this not rely on lints being order in execution order.
        for (int i = 1; i < messages.Count; i++)
        {
            if (!(messages[i - 1].Context.ContextObjects[^1] is Run prevRun &&
                  messages[i].Context.ContextObjects[0] is Run nextRun))
            {
                continue;
            }

            if (messages[i - 1].Message != messages[i].Message) continue;
            if (messages[i - 1].AutoFixed != messages[i].AutoFixed) continue;
            
            if (prevRun.NextSibling() != nextRun) continue;

            Context newContext = new Context(messages[i - 1].Context.Type,
                [..messages[i - 1].Context.ContextObjects, ..messages[i].Context.ContextObjects],
                messages[i - 1].Context.Before, messages[i - 1].Context.Text + messages[i].Context.Text,
                messages[i].Context.After);

            messages[i - 1] = new LintMessage(messages[i].Message, messages[i].AutoFixed, newContext);
            messages.RemoveAt(i);
            i -= 1;
        }
    }
}