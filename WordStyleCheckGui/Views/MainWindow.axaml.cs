using System;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Remote.Protocol.Input;
using WordStyleCheckGui.ViewModels;

namespace WordStyleCheckGui.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        
        DragDrop.SetAllowDrop(ContainerPanel, true);
        ContainerPanel.AddHandler(DragDrop.DragOverEvent, (_, args) =>
        {
            Console.WriteLine(args.DataTransfer);
            
            args.DragEffects &= DragDropEffects.Copy;

            if (!args.DataTransfer.Contains(DataFormat.File))
                args.DragEffects = DragDropEffects.None;
        });
        ContainerPanel.AddHandler(DragDrop.DropEvent, (_, args) =>
        {
            if (DataContext == null) return;

            args.DragEffects = DragDropEffects.Copy;
            
            foreach (var item in args.DataTransfer.Items)
            {
                if (item.TryGetValue(DataFormat.File) is { } file)
                {
                    ((MainWindowViewModel)DataContext).AddDocument(file);
                }
            }
        });
    }
}