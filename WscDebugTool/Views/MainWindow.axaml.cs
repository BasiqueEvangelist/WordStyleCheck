using System;
using Avalonia.Controls;
using AvaloniaEdit.TextMate;
using TextMateSharp.Grammars;
using WscDebugTool.ViewModels;

namespace WscDebugTool.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();

        var registryOptions = new RegistryOptions(ThemeName.DarkPlus);
        var textMateInstallation = CodeEditor.InstallTextMate(registryOptions);
        textMateInstallation.SetGrammar(registryOptions.GetScopeByLanguageId(registryOptions.GetLanguageByExtension(".xml").Id));
    }

    protected override void OnDataContextChanged(EventArgs e)
    {
        base.OnDataContextChanged(e);

        if (DataContext is MainWindowViewModel vm)
        {
            CodeEditor.Text = vm.FileText;

            vm.PropertyChanged += (_, args) =>
            {
                if (args.PropertyName == nameof(MainWindowViewModel.FileText))
                {
                    CodeEditor.Text = vm.FileText;
                }
            };
        }
    }
}