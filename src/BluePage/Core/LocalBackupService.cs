using System.Security.Cryptography;
using System.Text;
using Microsoft365OfficeWebLauncher.Config;

namespace Microsoft365OfficeWebLauncher.Core;

/// <summary>로컬 원본 폴더를 건드리지 않고 덮어쓰기 전 파일과 충돌 사본을 한곳에 보관한다.</summary>
public static class LocalBackupService
{
    public static string BackupRootDirectory =>
        Path.Combine(ConfigLoader.UserConfigDirectory, "Backups");

    public static string BackupBeforeOverwrite(string localFilePath)
    {
        var destination = BuildVersionPath(localFilePath, "덮어쓰기 전 로컬 백업");
        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
        File.Copy(localFilePath, destination, overwrite: false);
        return destination;
    }

    public static string BuildConflictCopyPath(string localFilePath)
    {
        var destination = BuildVersionPath(localFilePath, "온라인 충돌 사본");
        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
        return destination;
    }

    private static string BuildVersionPath(string localFilePath, string label)
    {
        var fullPath = Path.GetFullPath(localFilePath);
        var fileName = Path.GetFileNameWithoutExtension(fullPath);
        var extension = Path.GetExtension(fullPath);
        var pathHash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(fullPath.ToLowerInvariant())))[..8];
        var documentDirectory = Path.Combine(BackupRootDirectory, $"{fileName} ({pathHash})");
        var timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss-fff");
        return Path.Combine(documentDirectory, $"{fileName} ({label}, {timestamp}){extension}");
    }
}
