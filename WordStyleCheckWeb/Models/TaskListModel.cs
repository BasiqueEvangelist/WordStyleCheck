namespace WordStyleCheckWeb.Models;

public class TaskListModel
{
    public required List<DocumentProcessingService.DocumentTask> Tasks { get; init; }
}