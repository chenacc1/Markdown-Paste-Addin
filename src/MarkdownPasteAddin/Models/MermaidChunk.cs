namespace MarkdownPasteAddin.Models;

public class MermaidChunk : ContentChunk
{
    public string Code { get; }

    public MermaidChunk(string code)
    {
        Type = ChunkType.Mermaid;
        Code = code;
    }

    public override string ToString() => Code;
}
