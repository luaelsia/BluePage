namespace Microsoft365OfficeWebLauncher.Registry;

/// <summary>
/// Windows 로그인 시 자동 실행을 HKEY_CURRENT_USER\...\Run에 등록/해제한다. 관리자 권한이 필요 없다.
/// 별도 config.json 필드를 두지 않고, 이 레지스트리 값 자체를 유일한 상태 저장소로 삼는다.
/// </summary>
public sealed class StartupRegistrar
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "BluePage";
    private const string LegacyValueName = "WebOfficeLauncher";
    private const string MinimizedArg = "--minimized";

    public bool IsAutoStartEnabled()
    {
        using var runKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
        return runKey?.GetValue(ValueName) is string || runKey?.GetValue(LegacyValueName) is string;
    }

    public bool IsStartMinimized()
    {
        using var runKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
        var command = runKey?.GetValue(ValueName) as string ?? runKey?.GetValue(LegacyValueName) as string;
        return command?.Contains(MinimizedArg, StringComparison.OrdinalIgnoreCase) == true;
    }

    public void SetAutoStart(bool enabled, bool startMinimized, string exePath)
    {
        using var runKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RunKeyPath, writable: true)
            ?? throw new InvalidOperationException("HKCU Run 키를 열 수 없습니다.");

        if (!enabled)
        {
            runKey.DeleteValue(ValueName, throwOnMissingValue: false);
            runKey.DeleteValue(LegacyValueName, throwOnMissingValue: false);
            return;
        }

        var command = startMinimized ? $"\"{exePath}\" {MinimizedArg}" : $"\"{exePath}\"";
        runKey.SetValue(ValueName, command);
        runKey.DeleteValue(LegacyValueName, throwOnMissingValue: false);
    }
}
