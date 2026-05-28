namespace MarkdownPasteAddin.Models;

public class ImageChunk : ContentChunk
{
    public string Source { get; set; }
    public string AltText { get; set; }

    public ImageChunk(string source, string altText = "")
    {
        Type = ChunkType.Image;
        Source = source;
        AltText = altText;
    }
}
