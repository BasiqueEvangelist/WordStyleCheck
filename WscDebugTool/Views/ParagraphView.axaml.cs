using System;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Markup.Xaml;
using DocumentFormat.OpenXml.Wordprocessing;
using WscDebugTool.ViewModels;

namespace WscDebugTool.Views;

public partial class ParagraphView : UserControl
{
    public ParagraphView()
    {
        InitializeComponent();
    }

    protected override void OnDataContextChanged(EventArgs e)
    {
        base.OnDataContextChanged(e);

        if (DataContext is not ParagraphViewModel vm) return;

        Block.Inlines = new InlineCollection();
        
        Block.Inlines.AddRange(vm.Runs.Select(x => x.PrepareRun()));

        if (vm.Tool.Justification == JustificationValues.Center)
        {
            Block.HorizontalAlignment = HorizontalAlignment.Center;
        }
        
        if (vm.Tool.Justification == JustificationValues.Right)
        {
            Block.HorizontalAlignment = HorizontalAlignment.Right;
        }
        
        Block.AddHandler(PointerPressedEvent, (_, _) =>
        {
            vm.Focus();
        });
    }
}