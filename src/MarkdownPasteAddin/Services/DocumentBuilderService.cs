using System.Drawing;
using System.IO;
using MarkdownPasteAddin.Models;
using Word = Microsoft.Office.Interop.Word;

namespace MarkdownPasteAddin.Services;

public class DocumentBuilderService
{
    private readonly MarkdownParseService _parser;
    private readonly TableBuilderService _tableBuilder;
    private readonly MermaidRenderService _mermaidRenderer;
    private readonly ImageDownloadService _imageDownloader;
    private int _figureCount;
    private int _tableCount;

    public DocumentBuilderService(
        MarkdownParseService parser,
        TableBuilderService tableBuilder,
        MermaidRenderService mermaidRenderer)
    {
        _parser = parser;
        _tableBuilder = tableBuilder;
        _mermaidRenderer = mermaidRenderer;
        _imageDownloader = new ImageDownloadService();
        _figureCount = 0;
        _tableCount = 0;
    }

    public void Build(Word.Range startRange, string markdown)
    {
        var chunks = _parser.Parse(markdown);
        var currentRange = startRange;
        _figureCount = 0;
        _tableCount = 0;

        foreach (var chunk in chunks)
        {
            try
            {
                currentRange = InsertChunk(currentRange, chunk);
            }
            catch (Exception ex)
            {
                // Error recovery: insert error placeholder and continue
                var errorRange = GetDocEnd(currentRange);
                errorRange.Text = $"[Error: {ex.Message}]";
                errorRange.Font.Color = Word.WdColor.wdColorRed;
                errorRange.InsertParagraphAfter();
                currentRange = GetDocEnd(currentRange);
            }
        }
    }

    private Word.Range InsertChunk(Word.Range range, ContentChunk chunk)
    {
        switch (chunk)
        {
            case TextChunk textChunk:
                InsertText(range, textChunk.Text);
                break;

            case HeadingChunk headingChunk:
                InsertHeading(range, headingChunk);
                break;

            case TableChunk tableChunk:
                InsertTable(range, tableChunk);
                break;

            case MermaidChunk mermaidChunk:
                InsertMermaidImage(range, mermaidChunk);
                break;

            case ImageChunk imageChunk:
                InsertImage(range, imageChunk);
                break;

            case CodeChunk codeChunk:
                InsertCode(range, codeChunk);
                break;

            case TaskListChunk taskChunk:
                InsertTaskList(range, taskChunk);
                break;

            case BlockquoteChunk quoteChunk:
                InsertBlockquote(range, quoteChunk);
                break;

            case ContentChunk hr when hr.Type == ChunkType.HorizontalRule:
                InsertHorizontalRule(range);
                break;
        }

        return GetDocEnd(range);
    }

    private void InsertText(Word.Range range, string text)
    {
        // Handle markdown heading patterns
        if (text.StartsWith("#"))
        {
            var level = text.TakeWhile(c => c == '#').Count();
            var headingText = text.Substring(level).Trim();
            var para = range.Document.Paragraphs.Add(range);
            para.Range.Text = headingText;
            para.set_Style($"Heading {Math.Min(level, 9)}");
            para.Range.InsertParagraphAfter();
            return;
        }

        // Handle lists (UL/OL)
        if (text.StartsWith("- ") || text.StartsWith("* "))
        {
            var para = range.Document.Paragraphs.Add(range);
            para.Range.Text = text.Substring(2);
            para.set_Style("List Bullet");
            para.Range.InsertParagraphAfter();
            return;
        }

        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+[.)]\s"))
        {
            var content = System.Text.RegularExpressions.Regex.Replace(text, @"^\d+[.)]\s*", "");
            var para = range.Document.Paragraphs.Add(range);
            para.Range.Text = content;
            para.set_Style("List Number");
            para.Range.InsertParagraphAfter();
            return;
        }

        // Plain text
        range.Text = text;
        range.InsertParagraphAfter();
    }

    private void InsertHeading(Word.Range range, HeadingChunk chunk)
    {
        var para = range.Document.Paragraphs.Add(range);
        para.Range.Text = chunk.Text;
        para.set_Style($"Heading {Math.Min(chunk.Level, 9)}");
        para.Range.InsertParagraphAfter();
    }

    private void InsertTable(Word.Range range, TableChunk chunk)
    {
        _tableCount++;
        _tableBuilder.CreateTable(range, chunk);
        range.InsertParagraphAfter();
    }

    private void InsertMermaidImage(Word.Range range, MermaidChunk chunk)
    {
        var imagePath = _mermaidRenderer.Render(chunk.Code);

        if (imagePath != null)
        {
            var shape = range.InlineShapes.AddPicture(
                imagePath, LinkToFile: false, SaveWithDocument: true);
            _figureCount++;
            range.InsertParagraphAfter();

            // Add caption
            var captionRange = GetDocEnd(range);
            captionRange.Text = $"图 {_figureCount}";
            captionRange.Font.Size = 10;
            captionRange.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
            captionRange.InsertParagraphAfter();

            _mermaidRenderer.Cleanup(imagePath);
        }
        else
        {
            range.Text = $"[Mermaid diagram could not be rendered]\n```mermaid\n{chunk.Code}\n```";
            range.InsertParagraphAfter();
        }
    }

    private void InsertImage(Word.Range range, ImageChunk chunk)
    {
        var imgBytes = _imageDownloader.Download(chunk.Source);
        if (imgBytes != null && imgBytes.Length > 0)
        {
            var tmpPath = Path.Combine(Path.GetTempPath(), $"mdpaste_{Guid.NewGuid():N}.png");
            File.WriteAllBytes(tmpPath, imgBytes);

            try
            {
                range.InlineShapes.AddPicture(tmpPath,
                    LinkToFile: false, SaveWithDocument: true);
                _figureCount++;
                range.InsertParagraphAfter();

                if (!string.IsNullOrEmpty(chunk.AltText))
                {
                    var captionRange = GetDocEnd(range);
                    captionRange.Text = $"图 {_figureCount}：{chunk.AltText}";
                    captionRange.Font.Size = 10;
                    captionRange.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    captionRange.InsertParagraphAfter();
                }
            }
            finally
            {
                try { File.Delete(tmpPath); } catch { }
            }
        }
        else
        {
            range.Text = $"[Image: {chunk.AltText ?? chunk.Source}]";
            range.InsertParagraphAfter();
        }
    }

    private void InsertCode(Word.Range range, CodeChunk chunk)
    {
        foreach (var line in chunk.Code.Split('\n'))
        {
            var para = range.Document.Paragraphs.Add(range);
            para.Range.Text = line;
            para.Range.Font.Name = "Consolas";
            para.Range.Font.Size = 9;

            // Light background
            para.Range.Shading.BackgroundPatternColor =
                Word.WdColor.wdColorGray05;

            para.Range.InsertParagraphAfter();
        }
        range.InsertParagraphAfter();
    }

    private void InsertTaskList(Word.Range range, TaskListChunk chunk)
    {
        foreach (var item in chunk.Items)
        {
            var para = range.Document.Paragraphs.Add(range);
            para.Range.Text = $"{(item.Checked ? "☒" : "☐")}  {item.Text}";
            para.Range.Font.Size = 11;
            para.Range.InsertParagraphAfter();
        }
    }

    private void InsertBlockquote(Word.Range range, BlockquoteChunk chunk)
    {
        var para = range.Document.Paragraphs.Add(range);
        para.Range.Text = chunk.Text;
        para.Range.Font.Italic = 1;
        para.Range.Font.Color = Word.WdColor.wdColorGray50;
        para.Range.Font.Size = 10.5f;
        para.Range.ParagraphFormat.LeftIndent = 36; // ~0.5 inch
        para.Range.Shading.BackgroundPatternColor =
            Word.WdColor.wdColorGray05;
        para.Range.InsertParagraphAfter();
    }

    private void InsertHorizontalRule(Word.Range range)
    {
        var para = range.Document.Paragraphs.Add(range);
        para.Range.Text = "─".PadRight(60, '─');
        para.Range.Font.Color = Word.WdColor.wdColorGray50;
        para.Range.ParagraphFormat.Alignment =
            Word.WdParagraphAlignment.wdAlignParagraphCenter;
        para.Range.InsertParagraphAfter();
    }

    private Word.Range GetDocEnd(Word.Range afterRange)
    {
        return afterRange.Document.Range(
            afterRange.Document.Content.End - 1,
            afterRange.Document.Content.End - 1);
    }
}
