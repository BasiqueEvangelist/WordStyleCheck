using System;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Remote.Protocol.Input;
using WordStyleCheckGui.ViewModels;
using MouseButton = Avalonia.Input.MouseButton;

namespace WordStyleCheckGui.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        
        DragDrop.SetAllowDrop(DropZone, true);
        DropZone.AddHandler(DragDrop.DragOverEvent, (_, args) =>
        {
            Console.WriteLine(args.DataTransfer);
            
            args.DragEffects &= DragDropEffects.Copy;

            if (!args.DataTransfer.Contains(DataFormat.File))
                args.DragEffects = DragDropEffects.None;
        });
        DropZone.AddHandler(DragDrop.DropEvent, (_, args) =>
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
        DropZone.AddHandler(PointerReleasedEvent, (_, args) =>
        {
            if (args.InitialPressMouseButton != MouseButton.Left) return;
            
            if (DataContext == null) return;

            ((MainWindowViewModel) DataContext).OpenDialog();
        });
    }
}