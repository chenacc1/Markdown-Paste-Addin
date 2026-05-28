namespace MarkdownPasteAddin.Models;

public class MathChunk : ContentChunk
{
    public string Latex { get; set; }
    public bool Display { get; set; }

    public MathChunk(string latex, bool display = false)
    {
        Type = ChunkType.Math;
        Latex = latex;
        Display = display;
    }
}
