namespace MarkdownPasteAddin.Models;

public class TableChunk : ContentChunk
{
    public List<string> Headers { get; }
    public List<ColumnAlignment> Alignments { get; }
    public List<List<string>> Rows { get; }

    public int ColumnCount => Headers.Count;
    public int RowCount => Rows.Count;

    public TableChunk(List<string> headers, List<ColumnAlignment> alignments, List<List<string>> rows)
    {
        Type = ChunkType.Table;
        Headers = headers;
        Alignments = alignments;
        Rows = rows;
    }
}

public enum ColumnAlignment
{
    Left,
    Center,
    Right
}
