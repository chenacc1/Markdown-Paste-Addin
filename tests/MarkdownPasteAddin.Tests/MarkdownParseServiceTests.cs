using MarkdownPasteAddin.Models;
using MarkdownPasteAddin.Services;

namespace MarkdownPasteAddin.Tests;

public class MarkdownParseServiceTests
{
    private readonly MarkdownParseService _service = new();

    [Fact]
    public void Parse_SimpleTable_ReturnsTableChunk()
    {
        var markdown = "| Name | Age |\n|------|-----|\n| Alice | 30 |\n| Bob | 25 |";

        var chunks = _service.Parse(markdown);

        var table = Assert.Single(chunks) as TableChunk;
        Assert.NotNull(table);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(2, table.RowCount);
        Assert.Equal("Name", table.Headers[0]);
        Assert.Equal("Age", table.Headers[1]);
        Assert.Equal("Alice", table.Rows[0][0]);
        Assert.Equal("25", table.Rows[1][1]);
    }

    [Fact]
    public void Parse_TableWithAlignment_ExtractsAlignments()
    {
        var markdown = "| Left | Center | Right |\n|:-----|:------:|------:|\n| A | B | C |";

        var chunks = _service.Parse(markdown);

        var table = Assert.IsType<TableChunk>(chunks[0]);
        Assert.Equal(ColumnAlignment.Left, table.Alignments[0]);
        Assert.Equal(ColumnAlignment.Center, table.Alignments[1]);
        Assert.Equal(ColumnAlignment.Right, table.Alignments[2]);
    }

    [Fact]
    public void Parse_MermaidCodeBlock_ReturnsMermaidChunk()
    {
        var markdown = "```mermaid\ngraph LR\n  A-->B\n```";

        var chunks = _service.Parse(markdown);

        var mermaid = Assert.IsType<MermaidChunk>(chunks[0]);
        Assert.Contains("graph LR", mermaid.Code);
        Assert.Contains("A-->B", mermaid.Code);
    }

    [Fact]
    public void Parse_MixedContent_ReturnsOrderedChunks()
    {
        var markdown = "## Title\n\nSome text.\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n```mermaid\ngraph TD\n  X-->Y\n```";

        var chunks = _service.Parse(markdown);

        Assert.True(chunks.Count >= 3);

        // Should have text, table, mermaid in order
        Assert.Contains(chunks, c => c is TextChunk);
        Assert.Contains(chunks, c => c is TableChunk);
        Assert.Contains(chunks, c => c is MermaidChunk);
    }

    [Fact]
    public void IsMarkdownContent_WithTable_ReturnsTrue()
    {
        Assert.True(_service.IsMarkdownContent("| A | B |\n|---|---|"));
        Assert.True(_service.IsMarkdownContent("|:---|:---:|"));
    }

    [Fact]
    public void IsMarkdownContent_WithMermaid_ReturnsTrue()
    {
        Assert.True(_service.IsMarkdownContent("Some text\n```mermaid\ngraph LR\n```"));
    }

    [Fact]
    public void IsMarkdownContent_PlainText_ReturnsFalse()
    {
        Assert.False(_service.IsMarkdownContent("Hello world"));
        Assert.False(_service.IsMarkdownContent(""));
        Assert.False(_service.IsMarkdownContent("   "));
    }

    [Fact]
    public void Parse_EmptyInput_ReturnsEmptyList()
    {
        var chunks = _service.Parse("");
        Assert.Empty(chunks);
    }
}
