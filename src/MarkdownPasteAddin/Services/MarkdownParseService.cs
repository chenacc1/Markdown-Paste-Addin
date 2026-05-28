using System.Text.RegularExpressions;
using Markdig;
using Markdig.Syntax;
using MarkdownPasteAddin.Models;

namespace MarkdownPasteAddin.Services;

public class MarkdownParseService
{
    public List<ContentChunk> Parse(string markdown)
    {
        var document = Markdown.Parse(markdown);
        var chunks = new List<ContentChunk>();

        foreach (var block in document)
        {
            var chunk = ConvertBlock(block);
            if (chunk != null)
                chunks.Add(chunk);
        }

        return chunks;
    }

    public bool IsMarkdownContent(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var hasTable = text.Contains("|---") || text.Contains("| ---") ||
                       text.Contains("|:---") || text.Contains("| :---");
        var hasMermaid = text.Contains("```mermaid") || text.Contains("``` mermaid");
        var hasCode = Regex.IsMatch(text, @"```\w+");
        var hasMath = text.Contains("$$") || text.Contains("$");
        var hasTask = Regex.IsMatch(text, @"^\s*[-*+]\s+\[[ xX]\]", RegexOptions.Multiline);

        return hasTable || hasMermaid || hasCode || hasMath || hasTask;
    }

    private ContentChunk? ConvertBlock(Block block)
    {
        switch (block)
        {
            case Markdig.Syntax.Table table:
                return ConvertTable(table);

            case FencedCodeBlock fenced when IsMermaidBlock(fenced):
                return ConvertMermaid(fenced);

            case FencedCodeBlock fenced:
                return ConvertCode(fenced);

            case ParagraphBlock para:
                return ConvertParagraph(para);

            case HeadingBlock heading:
                return new HeadingChunk(heading.Level,
                    heading.Inline?.FirstChild?.ToString() ?? "");

            case QuoteBlock quote:
                return ConvertBlockquote(quote);

            case ThematicBreakBlock:
                return new ContentChunk { Type = ChunkType.HorizontalRule };

            default:
                return ConvertGenericBlock(block);
        }
    }

    private ContentChunk? ConvertParagraph(ParagraphBlock para)
    {
        var text = para.Inline?.FirstChild?.ToString() ?? "";
        if (string.IsNullOrEmpty(text))
        {
            var raw = para.ToPositionText();
            if (string.IsNullOrWhiteSpace(raw))
                return null;
            text = raw;
        }

        // Detect task list: - [ ] or - [x]
        var taskMatch = Regex.Match(text, @"^\s*[-*+]\s+\[([ xX])\]\s+(.*)");
        if (taskMatch.Success)
        {
            // Collect all task items that follow
            return new TaskListChunk(new List<TaskItem>());
        }

        // Detect math: $$...$$
        if (text.Contains("$$"))
        {
            var mathMatch = Regex.Match(text, @"\$\$(.+?)\$\$", RegexOptions.Singleline);
            if (mathMatch.Success)
                return new MathChunk(mathMatch.Groups[1].Value.Trim(), display: true);
        }
        if (text.StartsWith("$") && text.EndsWith("$"))
            return new MathChunk(text.Trim('$'), display: false);

        // Detect blockquote continuation
        if (text.StartsWith("> "))
            return new BlockquoteChunk(text.TrimStart('>', ' '));

        return new TextChunk(text);
    }

    private ContentChunk? ConvertBlockquote(QuoteBlock quote)
    {
        var lines = new List<string>();
        foreach (var block in quote)
        {
            if (block is ParagraphBlock pb)
                lines.Add(pb.Inline?.FirstChild?.ToString() ?? "");
        }
        return new BlockquoteChunk(string.Join("\n", lines));
    }

    private ContentChunk ConvertTable(Markdig.Syntax.Table table)
    {
        var alignments = new List<ColumnAlignment>();
        foreach (var col in table.ColumnDefinitions)
        {
            alignments.Add(col.Alignment switch
            {
                Markdig.Syntax.TableColumnAlign.Left => ColumnAlignment.Left,
                Markdig.Syntax.TableColumnAlign.Center => ColumnAlignment.Center,
                Markdig.Syntax.TableColumnAlign.Right => ColumnAlignment.Right,
                _ => ColumnAlignment.Left
            });
        }

        var headers = new List<string>();
        var rows = new List<List<string>>();

        int rowIdx = 0;
        foreach (var row in table)
        {
            var cells = new List<string>();
            foreach (var cell in row)
                cells.Add(ExtractCellText(cell));

            if (rowIdx == 0 && table.HeaderRowCount > 0)
                headers = cells;
            else
                rows.Add(cells);
            rowIdx++;
        }

        if (headers.Count == 0 && rows.Count > 0)
        {
            headers = rows[0];
            rows.RemoveAt(0);
        }

        var colCount = headers.Count;
        while (alignments.Count < colCount)
            alignments.Add(ColumnAlignment.Left);

        return new TableChunk(headers, alignments, rows);
    }

    private ContentChunk ConvertMermaid(FencedCodeBlock fenced)
    {
        var code = string.Join("\n", fenced.Lines);
        return new MermaidChunk(code);
    }

    private ContentChunk ConvertCode(FencedCodeBlock fenced)
    {
        var lang = fenced.Info?.Trim() ?? "";
        var code = string.Join("\n", fenced.Lines);
        return new CodeChunk(lang, code);
    }

    private ContentChunk? ConvertGenericBlock(Block block)
    {
        if (block is LeafBlock leaf && leaf.Lines.Count > 0)
        {
            var text = string.Join("\n", leaf.Lines);
            if (!string.IsNullOrWhiteSpace(text))
                return new TextChunk(text);
        }
        return null;
    }

    private string ExtractCellText(Block cell)
    {
        if (cell is ParagraphBlock para)
            return para.Inline?.FirstChild?.ToString() ?? "";
        return "";
    }

    private static bool IsMermaidBlock(FencedCodeBlock fenced)
    {
        var info = fenced.Info;
        if (string.IsNullOrEmpty(info))
            return false;
        return info.Trim().Equals("mermaid", StringComparison.OrdinalIgnoreCase)
            || info.StartsWith("mermaid", StringComparison.OrdinalIgnoreCase);
    }
}
