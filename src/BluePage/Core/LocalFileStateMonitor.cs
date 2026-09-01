namespace Microsoft365OfficeWebLauncher.Core;

public readonly record struct LocalFileState(long Length, DateTime LastWriteUtc);

public static class LocalFileStateMonitor
{
    public static bool TryRead(string localFilePath, out LocalFileState state)
    {
        var file = new FileInfo(localFilePath);
        if (!file.Exists)
        {
            state = default;
            return false;
        }

        state = new LocalFileState(file.Length, file.LastWriteTimeUtc);
        return true;
    }
}
