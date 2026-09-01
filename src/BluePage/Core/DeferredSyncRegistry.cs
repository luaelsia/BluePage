using System.Security.Cryptography;
using System.Text;

namespace Microsoft365OfficeWebLauncher.Core;

/// <summary>프로세스가 달라도 공유되는 파일별 백그라운드 동기화 보류 상태.</summary>
public sealed class DeferredSyncRegistry
{
    private readonly string _rootDirectory;

    public DeferredSyncRegistry(string? rootDirectory = null)
    {
        _rootDirectory = rootDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft365OfficeWebLauncher",
            "deferred-sync");
    }

    public bool IsDeferred(string localFilePath) => File.Exists(GetMarkerPath(localFilePath));

    public void Defer(string localFilePath)
    {
        Directory.CreateDirectory(_rootDirectory);
        File.WriteAllText(GetMarkerPath(localFilePath), Path.GetFullPath(localFilePath), Encoding.UTF8);
    }

    public void Resume(string localFilePath)
    {
        var markerPath = GetMarkerPath(localFilePath);
        if (File.Exists(markerPath)) File.Delete(markerPath);
    }

    private string GetMarkerPath(string localFilePath)
    {
        var normalizedPath = Path.GetFullPath(localFilePath).ToLowerInvariant();
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(normalizedPath)))[..16];
        return Path.Combine(_rootDirectory, $"{hash}.deferred");
    }
}
