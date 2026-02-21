using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck;

public static class FieldStackTracker
{
    private static readonly ConditionalWeakTable<OpenXmlElement, List<FieldStackEntry>> FieldStackMappings = new();

    public static List<FieldStackEntry> GetContextFor(OpenXmlElement el) => FieldStackMappings.GetOrCreateValue(el);
    
    public static void Run(Document doc)
    {
        if (doc.Body == null) return;
        
        List<FieldStackEntry> stack = [];

        Run(stack, doc.Body);
    }

    private static void Run(List<FieldStackEntry> stack, OpenXmlElement element)
    {
        if (element is FieldChar fldChar)
        {
            if (fldChar.FieldCharType?.Value == FieldCharValues.Begin)
            {
                stack.Add(new FieldStackEntry());
            }
            else if (fldChar.FieldCharType?.Value == FieldCharValues.End)
            {
                stack.RemoveAt(stack.Count - 1);
            }
        } else if (element is FieldCode instrText)
        {
            stack[^1].InstrText = instrText.Text;
        }

        if (stack.Count > 0 && element is Paragraph)
        {
            FieldStackMappings.Add(element, [..stack]);
        }

        foreach (var child in element.ChildElements)
        {
            Run(stack, child);
        }
        
        if (stack.Count > 0 && element is Paragraph)
        {
            var mapping = FieldStackMappings.GetOrCreateValue(element);

            foreach (var el in stack)
            {
                if (!mapping.Contains(el))
                    mapping.Add(el);
            }
        }
    }

    public class FieldStackEntry
    {
        public string? InstrText;
    }
}