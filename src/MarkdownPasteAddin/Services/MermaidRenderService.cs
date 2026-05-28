using System.Diagnostics;
using System.IO;

namespace MarkdownPasteAddin.Services;

public class MermaidRenderService
{
    private readonly string _outputDir;
    private readonly string _mmdcPath;
    private readonly int _scale;
    private readonly int _timeoutMs;

    public MermaidRenderService(string? outputDir = null, string? mmdcPath = null,
        int scale = 2, int timeoutMs = 30000)
    {
        _outputDir = outputDir ?? Path.Combine(Path.GetTempPath(), "MarkdownPasteAddin");
        _mmdcPath = mmdcPath ?? "mmdc";
        _scale = scale;
        _timeoutMs = timeoutMs;
    }

    public async Task<string?> RenderAsync(string mermaidCode)
    {
        Directory.CreateDirectory(_outputDir);

        var inputFile = Path.Combine(_outputDir, $"{Guid.NewGuid():N}.mmd");
        var outputFile = Path.Combine(_outputDir, $"{Guid.NewGuid():N}.png");

        try
        {
            await File.WriteAllTextAsync(inputFile, mermaidCode);

            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = _mmdcPath,
                    Arguments = $"-i \"{inputFile}\" -o \"{outputFile}\" -s {_scale} -b transparent --puppeteerConfigFile \"{GetPuppeteerConfig()}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true
                }
            };

            process.Start();
            var completed = await Task.Run(() => process.WaitForExit(_timeoutMs));

            if (!completed)
            {
                process.Kill();
                return null;
            }

            if (process.ExitCode != 0)
            {
                var error = await process.StandardError.ReadToEndAsync();
                Debug.WriteLine($"mmdc error (exit {process.ExitCode}): {error}");
                return null;
            }

            if (File.Exists(outputFile) && new FileInfo(outputFile).Length > 0)
                return outputFile;

            return null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Mermaid render failed: {ex.Message}");
            return null;
        }
        finally
        {
            // Clean up input file
            try { File.Delete(inputFile); } catch { }
        }
    }

    public string? Render(string mermaidCode)
    {
        return RenderAsync(mermaidCode).GetAwaiter().GetResult();
    }

    private string GetPuppeteerConfig()
    {
        var configPath = Path.Combine(_outputDir, "puppeteer-config.json");
        if (!File.Exists(configPath))
        {
            File.WriteAllText(configPath, @"{""args"": [""--no-sandbox"", ""--disable-setuid-sandbox""]}");
        }
        return configPath;
    }

    public void Cleanup(string imagePath)
    {
        try
        {
            if (File.Exists(imagePath))
                File.Delete(imagePath);
        }
        catch { }
    }
}
