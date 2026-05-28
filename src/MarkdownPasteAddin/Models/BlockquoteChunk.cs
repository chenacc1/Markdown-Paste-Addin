namespace MarkdownPasteAddin.Models;

public class BlockquoteChunk : ContentChunk
{
    public string Text { get; set; }

    public BlockquoteChunk(string text)
    {
        Type = ChunkType.Blockquote;
        Text = text;
    }
}
