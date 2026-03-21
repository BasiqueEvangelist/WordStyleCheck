using System.Collections.Generic;
using WordStyleCheck;
using WordStyleCheck.Context;

namespace WordStyleCheckGui.ViewModels
{
    public class DiagnosticViewModel : ViewModelBase
    {
        public string TranslatedDescription { get; }

        public List<DiagnosticContextLine> Context { get; }

        public DiagnosticViewModel(XmlTranslationsFile translations, LintMessage message)
        {
            TranslatedDescription = Utils.ToPlainText(translations.Translate(message.Id, message.Parameters ?? new(), null));
            Context = message.Context.Lines;
        }
    }
}
