using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using AvaloniaEdit.TextMate;
using TextMateSharp.Grammars;
using WscDebugTool.ViewModels;

namespace WscDebugTool.Views;

public partial class DocumentFileView : UserControl
{
    public DocumentFileView()
    {
        InitializeComponent();
        
        var registryOptions = new RegistryOptions(ThemeName.DarkPlus);
        var textMateInstallation = CodeEditor.InstallTextMate(registryOptions);
        textMateInstallation.SetGrammar(registryOptions.GetScopeByLanguageId(registryOptions.GetLanguageByExtension(".xml").Id));
    }
    
    protected override void OnDataContextChanged(EventArgs e)
    {
        base.OnDataContextChanged(e);

        if (DataContext is DocumentFileViewModel vm)
        {
            CodeEditor.Text = vm.FileText;

            vm.PropertyChanged += (_, args) =>
            {
                if (args.PropertyName == nameof(XmlFileViewModel.FileText))
                {
                    CodeEditor.Text = vm.FileText;
                }
            };

            vm.FocusParagraph += (p) =>
            {
                CodeEditor.TextArea.Caret.Line = p.Line;
                CodeEditor.TextArea.Caret.Column = p.Column;
                CodeEditor.ScrollTo(p.Line, p.Column);
            };
        }
    }
}