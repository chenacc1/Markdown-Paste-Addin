using System.Reflection;
using System.Runtime.InteropServices;
using MarkdownPasteAddin.Services;
using Serilog;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace MarkdownPasteAddin;

[ComVisible(true)]
[Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
[ProgId("MarkdownPasteAddin.ThisAddIn")]
public class ThisAddIn : IDisposable
{
    private Word.Application? _wordApp;
    private DocumentBuilderService? _documentBuilder;
    private ClipboardService? _clipboardService;
    private MarkdownParseService? _markdownParser;

    public void Initialize(Word.Application application)
    {
        _wordApp = application;

        ConfigureLogging();

        _clipboardService = new ClipboardService();
        _markdownParser = new MarkdownParseService();

        var tableBuilder = new TableBuilderService();
        var mermaidRenderer = new MermaidRenderService();

        _documentBuilder = new DocumentBuilderService(
            _markdownParser, tableBuilder, mermaidRenderer);

        Log.Information("MarkdownPasteAddin initialized successfully");
    }

    public void SmartPaste()
    {
        if (_wordApp?.ActiveDocument == null)
        {
            Log.Warning("SmartPaste: No active document");
            return;
        }

        try
        {
            var clipboardText = _clipboardService?.GetText();
            var selection = _wordApp.Selection;

            if (string.IsNullOrEmpty(clipboardText))
            {
                // No text in clipboard, do regular paste
                selection.Paste();
                return;
            }

            if (_markdownParser == null || !_markdownParser.IsMarkdownContent(clipboardText))
            {
                // Not markdown, do regular paste
                selection.Paste();
                return;
            }

            // It's markdown content - process with smart paste
            _documentBuilder?.Build(selection.Range, clipboardText);

            Log.Information("SmartPaste completed successfully");
        }
        catch (Exception ex)
        {
            Log.Error(ex, "SmartPaste failed");
            // Fallback to regular paste on error
            try { _wordApp?.Selection.Paste(); } catch { }
        }
    }

    private void ConfigureLogging()
    {
        var logDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MarkdownPasteAddin", "Logs");

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.File(
                Path.Combine(logDir, "addin-.log"),
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 7)
            .CreateLogger();
    }

    public void Dispose()
    {
        Log.Information("MarkdownPasteAddin shutting down");
        Log.CloseAndFlush();
        _wordApp = null;
    }
}
