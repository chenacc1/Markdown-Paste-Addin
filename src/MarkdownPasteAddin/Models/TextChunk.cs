namespace MarkdownPasteAddin.Models;

public class TextChunk : ContentChunk
{
    public string Text { get; }

    public TextChunk(string text)
    {
        Type = ChunkType.Text;
        Text = text;
    }

    public override string ToString() => Text;
}

public class HeadingChunk : ContentChunk
{
    public int Level { get; set; }
    public string Text { get; set; }

    public HeadingChunk(int level, string text)
    {
        Type = ChunkType.Heading;
        Level = level;
        Text = text;
    }
}
