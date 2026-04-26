using System.Collections.Generic;
using WordStyleCheck;
using WordStyleCheck.Context;

namespace WordStyleCheckGui.ViewModels
{
    public class DiagnosticViewModel : ViewModelBase
    {
        public string TranslatedDescription { get; }

        public List<DiagnosticContextLine> Context { get; }

        public DiagnosticViewModel(XmlTranslationsFile translations, LintDiagnostic diagnostic)
        {
            TranslatedDescription = Utils.ToPlainText(translations.Translate(diagnostic.Id, diagnostic.Parameters ?? new(), null));
            Context = diagnostic.Context.Lines;
        }
    }
}
