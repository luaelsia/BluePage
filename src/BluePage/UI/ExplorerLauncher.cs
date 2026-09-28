using System.Diagnostics;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>탐색기로 폴더를 열거나, 파일이 선택된 상태로 그 파일이 있는 폴더를 연다.</summary>
public static class ExplorerLauncher
{
    public static void OpenFolder(string directory)
    {
        Directory.CreateDirectory(directory);
        Process.Start(new ProcessStartInfo("explorer.exe", $"\"{directory}\"") { UseShellExecute = true });
    }

    public static void RevealFile(string filePath)
    {
        if (!File.Exists(filePath))
        {
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
            {
                OpenFolder(directory);
            }
            return;
        }
        Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{filePath}\"") { UseShellExecute = true });
    }
}
