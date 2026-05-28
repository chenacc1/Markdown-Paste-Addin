using System.Net.Http;

namespace MarkdownPasteAddin.Services;

public class ImageDownloadService
{
    private static readonly HttpClient _client = new()
    {
        Timeout = TimeSpan.FromSeconds(15)
    };

    static ImageDownloadService()
    {
        _client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0");
    }

    public byte[]? Download(string url)
    {
        try
        {
            // Handle data: URIs
            if (url.StartsWith("data:image/"))
            {
                var commaIdx = url.IndexOf(',');
                if (commaIdx > 0)
                {
                    var b64 = url.Substring(commaIdx + 1);
                    return Convert.FromBase64String(b64);
                }
                return null;
            }

            return _client.GetByteArrayAsync(url).GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Image download failed: {url} ({ex.Message})");
            return null;
        }
    }
}
