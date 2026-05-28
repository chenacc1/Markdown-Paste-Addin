using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MarkdownPasteAddin.Services;

public class ClipboardContent
{
    public string? Format { get; set; }  // "html" or "text"
    public string? Content { get; set; }
}

public class ClipboardService
{
    public ClipboardContent? GetContent()
    {
        try
        {
            // Try HTML first
            if (Clipboard.ContainsText(TextDataFormat.Html))
            {
                var html = Clipboard.GetText(TextDataFormat.Html);
                return new ClipboardContent
                {
                    Format = "html",
                    Content = DecodeClipboardHtml(html)
                };
            }

            if (Clipboard.ContainsText(TextDataFormat.UnicodeText))
            {
                return new ClipboardContent
                {
                    Format = "text",
                    Content = Clipboard.GetText(TextDataFormat.UnicodeText)
                };
            }

            if (Clipboard.ContainsText(TextDataFormat.Text))
            {
                return new ClipboardContent
                {
                    Format = "text",
                    Content = Clipboard.GetText(TextDataFormat.Text)
                };
            }

            return null;
        }
        catch
        {
            return null;
        }
    }

    public string? GetText()
    {
        return Clipboard.ContainsText() ? Clipboard.GetText() : null;
    }

    public bool ContainsText()
    {
        return Clipboard.ContainsText();
    }

    public void SetText(string text)
    {
        Clipboard.SetText(text);
    }

    private static string DecodeClipboardHtml(string raw)
    {
        // CF_HTML format: strip the header
        var match = Regex.Match(raw, @"<html[^>]*>", RegexOptions.IgnoreCase);
        if (!match.Success)
            match = Regex.Match(raw, @"<[^>]+>");
        if (match.Success)
            return raw.Substring(match.Index);
        return raw;
    }
}
