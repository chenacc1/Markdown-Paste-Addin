namespace MarkdownPasteAddin.Models;

public class TaskItem
{
    public bool Checked { get; set; }
    public string Text { get; set; } = "";
}

public class TaskListChunk : ContentChunk
{
    public List<TaskItem> Items { get; set; }

    public TaskListChunk(List<TaskItem> items)
    {
        Type = ChunkType.TaskList;
        Items = items;
    }
}
