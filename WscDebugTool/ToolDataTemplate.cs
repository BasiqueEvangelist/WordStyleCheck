using System;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using WscDebugTool.ViewModels;
using WscDebugTool.Views;

namespace WscDebugTool;

public class ToolDataTemplate : IDataTemplate
{
    public Control? Build(object? param)
    {
        if (param is ParagraphViewModel)
        {
            return new ParagraphView();
        }
        
        if (param is DocumentFileViewModel)
        {
            return new DocumentFileView();
        }
        
        if (param is XmlFileViewModel)
        {
            return new XmlFileView();
        }

        if (param is ImageFileViewModel)
        {
            return new ImageFileView();
        }

        throw new NotImplementedException();
    }

    public bool Match(object? data)
    {
        return data is ViewModelBase;
    }
}