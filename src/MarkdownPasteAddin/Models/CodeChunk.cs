namespace MarkdownPasteAddin.Models;

public class CodeChunk : ContentChunk
{
    public string Language { get; set; }
    public string Code { get; set; }

    public CodeChunk(string language, string code)
    {
        Type = ChunkType.Code;
        Language = language;
        Code = code;
    }
}
