using System.Text;
using WordStyleCheck.Context;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Ntk;

public class SupervisorFixerLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["IncorrectSupervisorFormatting"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (!tool.GetFeature(NtkParagraphData.Key)!.IsSupervisorData) continue;
            
            RunAssociatedText rat = RunAssociatedText.FromParagraph(tool);

            int start = rat.Text.IndexOf(':') + 1;
            
            while (char.IsWhiteSpace(rat.Text[start])) start++;

            string name = rat.Text.Substring(start);

            List<string> doctors = [];
            List<string> candidates = [];
            bool docent = false;
            bool professor = false;
            bool assistant = false;
            bool lecturer = false;
            bool seniorLecturer = false;

            int i = 0;

            bool ConsumeIfNext(params string[] args)
            {
                foreach (var option in args)
                {
                    if (name.Length - i < option.Length) continue;

                    if (name.Substring(i, option.Length).Equals(option, StringComparison.InvariantCultureIgnoreCase))
                    {
                        i += option.Length;
                        return true;
                    }
                }

                return false;
            }

            void ConsumeWhitespace()
            {
                while (i < name.Length && char.IsWhiteSpace(name[i])) i += 1;
            }

            void HandleOfSciences(List<string> addTo)
            {
                if (ConsumeIfNext("в.", "военных", "воен."))
                {
                    addTo.Add("в.");
                }
                else if (ConsumeIfNext("т.", "технических", "техн."))
                {
                    addTo.Add("т.");
                }
                else if (ConsumeIfNext("ф.-м.", "физико-математических", "физ.-мат."))
                {
                    addTo.Add("ф.-м.");
                }
                else if (ConsumeIfNext("э.", "экономических", "экон."))
                {
                    addTo.Add("э.");
                }
                else if (ConsumeIfNext("х.", "химических", "хим."))
                {
                    addTo.Add("х.");
                }
                else if (ConsumeIfNext("и.", "исторических", "ист."))
                {
                    addTo.Add("и.");
                }
                else if (ConsumeIfNext("ю.", "юридических", "юрид."))
                {
                    addTo.Add("ю.");
                }
                else if (ConsumeIfNext("с.", "социологических", "социол."))
                {
                    addTo.Add("с.");
                }
                else if (ConsumeIfNext("фарм.", "фармацевтических", "фармацевт."))
                {
                    addTo.Add("фарм.");
                }
            }
            
            while (i < name.Length)
            {
                if (char.IsWhiteSpace(name[i]))
                {
                    i += 1;
                }
                else if (ConsumeIfNext("д.", "д-р", "доктор"))
                {
                    ConsumeWhitespace();
                    HandleOfSciences(doctors);
                }
                else if (ConsumeIfNext("к.", "канд.", "кандидат"))
                {
                    ConsumeWhitespace();
                    HandleOfSciences(candidates);
                }
                else if (ConsumeIfNext("асс.", "ассистент"))
                {
                    assistant = true;
                }
                else if (ConsumeIfNext("препо.", "преп.", "преподаватель"))
                {
                    lecturer = true;
                }
                else if (ConsumeIfNext("ст.преп.", "ст. преп.", "старший преподаватель"))
                {
                    seniorLecturer = true;
                }
                else if (ConsumeIfNext("доц.", "доцент"))
                {
                    docent = true;
                }
                else if (ConsumeIfNext("проф.", "профессор"))
                {
                    professor = true;
                }
                else
                {
                    i += 1;
                }
            }

            List<string> names = StripAuthorJunkLint.FindNames(name);
            
            if (names.Count < 1) continue;

            StringBuilder correct = new();

            bool first = true;

            foreach (var doctor in doctors)
            {
                if (!first) correct.Append(", ");

                first = false;

                correct.Append("д." + doctor + "н.");
            }
            
            foreach (var candidate in candidates)
            {
                if (!first) correct.Append(", ");

                first = false;

                correct.Append("к." + candidate + "н.");
            }

            if (docent)
            {
                if (!first) correct.Append(", ");

                first = false;

                correct.Append("доц.");
            }
            
            if (professor)
            {
                if (!first) correct.Append(", ");

                first = false;

                correct.Append("проф.");
            }
            
            if (assistant)
            {
                if (!first) correct.Append(", ");

                first = false;

                correct.Append("асс.");
            }
            
            if (lecturer)
            {
                if (!first) correct.Append(", ");

                first = false;

                correct.Append("препо.");
            }
            
            if (seniorLecturer)
            {
                if (!first) correct.Append(", ");

                first = false;

                correct.Append("ст.преп.");
            }

            if (!first) correct.Append(" ");

            correct.Append(names[0]);
            
            if (name == correct.ToString()) continue;
            
            var span = rat.GetSpan(start, rat.Text.Length - start);
            
            if (!ctx.AutomaticallyFix)
            {
                ctx.AddMessage(new LintDiagnostic("IncorrectSupervisorFormatting", DiagnosticType.FormattingError, new RunSpanDiagnosticContext(span))
                {
                    Parameters = new()
                    {
                        ["Expected"] = correct.ToString(),
                        ["Actual"] = name
                    }
                });
            }
            else
            {
                ctx.MarkAutoFixed();
                    
                span.Replace(correct.ToString());
            }
        }
    }
}