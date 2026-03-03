using System;
using System.Collections.Generic;
using System.Text;
using Tmds.DBus.Protocol;
using WordStyleCheck;
using WordStyleCheck.Context;

namespace WordStyleCheckGui.ViewModels
{
    public class DiagnosticViewModel : ViewModelBase
    {
        public string TranslatedDescription { get; }

        public List<DiagnosticContextLine> Context { get; }

        public DiagnosticViewModel(DiagnosticTranslationsFile translations, LintMessage message)
        {
            TranslatedDescription = Utils.ToPlainText(translations.Translate(message.Id, message.Parameters ?? new()));
            Context = message.Context.Lines;
        }
    }
}
