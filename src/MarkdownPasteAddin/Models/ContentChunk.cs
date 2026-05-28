namespace MarkdownPasteAddin.Models;

public enum ChunkType
{
    Text,
    Table,
    Mermaid,
    Image,
    Code,
    Math,
    TaskList,
    Blockquote,
    HorizontalRule,
    Heading
}

public abstract class ContentChunk
{
    public ChunkType Type { get; protected set; }
}
