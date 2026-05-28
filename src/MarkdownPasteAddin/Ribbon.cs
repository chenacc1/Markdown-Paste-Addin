using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace MarkdownPasteAddin;

[ComVisible(true)]
public class Ribbon : Office.IRibbonExtensibility
{
    private Word.Application _wordApp;
    private ThisAddIn? _addin;

    public Ribbon(Word.Application wordApp, ThisAddIn? addin = null)
    {
        _wordApp = wordApp;
        _addin = addin;
    }

    public string GetCustomUI(string ribbonID)
    {
        var assembly = Assembly.GetExecutingAssembly();
        using var stream = assembly.GetManifestResourceStream("MarkdownPasteAddin.Ribbon.xml");
        if (stream == null)
            return string.Empty;

        using var reader = new System.IO.StreamReader(stream);
        return reader.ReadToEnd();
    }

    public void Ribbon_Load(Office.IRibbonUI ribbonUI)
    {
        // Ribbon loaded - can store ribbonUI reference for invalidation
    }

    public void OnSmartPasteClick(Office.IRibbonControl control)
    {
        _addin?.SmartPaste();
    }

    public void OnSettingsClick(Office.IRibbonControl control)
    {
        // Open settings dialog (simplified - could show a WPF/Forms dialog)
        System.Windows.Forms.MessageBox.Show(
            "Markdown Smart Paste Add-in v2.0\n\n" +
            "Shortcut: Ctrl+Shift+V\n\n" +
            "Supports:\n" +
            "- Markdown/HTML tables → Word tables\n" +
            "- Mermaid diagrams → embedded images\n" +
            "- LaTeX math → Word equations\n" +
            "- Code blocks with syntax highlighting\n" +
            "- Task lists, blockquotes, horizontal rules\n" +
            "- Images (URL/download) → embedded\n" +
            "- Auto figure/table numbering\n" +
            "- Nested lists, multi-level headings",
            "Markdown Smart Paste");
    }
}
