using Microsoft365OfficeWebLauncher.OneDrive;

namespace Microsoft365OfficeWebLauncher.Cloud;

public enum CloudProvider
{
    Microsoft,
    Google
}

public sealed record CloudFileMetadata(
    string Id,
    string WebUrl,
    DateTimeOffset ModifiedUtc);

public interface ICloudDriveService
{
    CloudProvider Provider { get; }
    string DisplayName { get; }

    /// <summary>원격에 새 항목을 만든다. remoteFileName은 RemoteFileNaming.Build로 만든 충돌하지 않는 이름이다.</summary>
    Task<CloudFileMetadata> CreateAsync(string localFilePath, string remoteFileName, CancellationToken ct);
    Task<CloudFileMetadata> UpdateAsync(string remoteFileId, string localFilePath, CancellationToken ct);
    Task<CloudFileMetadata> GetMetadataAsync(string remoteFileId, CancellationToken ct);
    Task DownloadAsync(string remoteFileId, string destinationPath, CancellationToken ct);
    Task<string> GetBluePageFolderWebUrlAsync(CancellationToken ct);
    Task<RemoteLockState> GetLockStateAsync(string remoteFileId, CancellationToken ct);
}

public static class CloudProviderNames
{
    public static string DisplayName(this CloudProvider provider) => provider switch
    {
        CloudProvider.Google => "Google Workspace",
        _ => "Microsoft 365"
    };

    public static bool TryParse(string? value, out CloudProvider provider) =>
        Enum.TryParse(value, ignoreCase: true, out provider);
}
