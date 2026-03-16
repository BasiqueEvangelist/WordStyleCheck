using Avalonia.Controls;
using AvaloniaEdit;
using AvaloniaEdit.TextMate;
using TextMateSharp.Grammars;
using WscDebugTool.ViewModels;

namespace WscDebugTool.Views;

public partial class MainWindow : Window
{
    public MainWindow(MainWindowViewModel vm)
    {
        InitializeComponent();

        DataContext = vm;
        CodeEditor.Text = vm.FileText;

        vm.PropertyChanged += (sender, args) =>
        {
            if (args.PropertyName == nameof(MainWindowViewModel.FileText))
            {
                CodeEditor.Text = vm.FileText;
            }
        };

        var _registryOptions = new RegistryOptions(ThemeName.DarkPlus);

        var _textMateInstallation = CodeEditor.InstallTextMate(_registryOptions);

        _textMateInstallation.SetGrammar(_registryOptions.GetScopeByLanguageId(_registryOptions.GetLanguageByExtension(".xml").Id));
    }
}