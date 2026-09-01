using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Microsoft365OfficeWebLauncher.Config;
using Microsoft365OfficeWebLauncher.Logging;

namespace Microsoft365OfficeWebLauncher.Registry;

/// <summary>
/// HKEY_CURRENT_USER에만 등록하므로 관리자 권한이 필요 없다.
/// 주의: Windows 8 이후 정책상 앱은 스스로를 "기본 프로그램"으로 강제 지정할 수 없다.
/// 이 클래스는 "연결 프로그램(Open with)" 후보 및 설정 앱의 "기본 앱" 목록에 나타나도록 등록하는 것까지만 수행하며,
/// 실제 기본값 지정은 사용자가 Windows 설정에서 1회 확인해야 한다.
/// </summary>
public sealed class FileAssociationRegistrar
{
    private const string AppRegistryName = "BluePage";
    private const string ProgIdPrefix = "BluePage";
    private const string ExecutableName = "BluePage.exe";
    private const string LegacyAppRegistryName = "Microsoft365OfficeWebLauncher";

    [DllImport("shell32.dll")]
    private static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);

    private const int SHCNE_ASSOCCHANGED = 0x08000000;
    private const int SHCNF_IDLIST = 0x0000;

    private readonly FileLogger _logger;

    public FileAssociationRegistrar(FileLogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Windows 설정 앱의 "기본 앱" 페이지를 이 앱에 맞춰 바로 연다.
    /// (예전 IApplicationAssociationRegistrationUI COM API는 Windows 10부터 실제 창을 열지 않고
    /// "설정 > 앱 > 기본 앱으로 이동하세요"라는 안내 메시지만 띄우도록 바뀌어서, 대신 ms-settings 딥링크를 사용한다.
    /// registeredAppUser 파라미터는 HKCU\Software\RegisteredApplications에 등록한 값 이름을 그대로 사용한다.
    /// Windows 11 21H2(2023-04 누적 업데이트)/22H2/23H2 이상에서는 이 앱의 설정 페이지로 바로 이동하고,
    /// 그보다 오래된 버전에서는 매개변수를 무시하고 "기본 앱" 목록 페이지가 열린다.)
    /// </summary>
    public void OpenAdvancedAssociationUI()
    {
        var uri = "ms-settings:defaultapps?registeredAppUser=" + Uri.EscapeDataString(AppRegistryName);
        Process.Start(new ProcessStartInfo(uri) { UseShellExecute = true });
    }

    /// <summary>현재 파일 연결이 등록된 상태인지 확인한다(GUI 상태 표시용).</summary>
    public bool IsRegistered()
    {
        using var registeredApps = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\RegisteredApplications", writable: false);
        var value = registeredApps?.GetValue(AppRegistryName) as string;
        return !string.IsNullOrEmpty(value);
    }

    public void RegisterAll(AppConfig config, string exePath)
    {
        using var classesRoot = OpenClassesRoot();

        var extensions = config.DocumentTypes
            .SelectMany(d => d.Extensions.Select(e => (Ext: Normalize(e), d.OfficeApp)))
            .DistinctBy(e => e.Ext, StringComparer.OrdinalIgnoreCase)
            .ToList();

        RegisterApplicationEntry(classesRoot, exePath, extensions.Select(e => e.Ext));

        using var capabilities = OpenOrCreate(Microsoft.Win32.Registry.CurrentUser, $@"Software\{AppRegistryName}\Capabilities");

        capabilities.SetValue("ApplicationName", AppBrand.Name);
        capabilities.SetValue("ApplicationDescription", AppBrand.Description);

        using (var fileAssoc = OpenOrCreate(Microsoft.Win32.Registry.CurrentUser, $@"Software\{AppRegistryName}\Capabilities\FileAssociations"))
        {
            foreach (var (ext, officeApp) in extensions)
            {
                var progId = ProgIdFor(ext);
                RegisterProgId(classesRoot, progId, ext, officeApp, exePath);
                RegisterOpenWithEntry(classesRoot, ext, progId);
                RegisterOpenWithExecutable(classesRoot, ext);
                fileAssoc.SetValue(ext, progId);
                _logger.Info($"확장자 등록: {ext} -> {progId} ({officeApp})");
            }
        }

        using (var registeredApps = OpenOrCreate(Microsoft.Win32.Registry.CurrentUser, @"Software\RegisteredApplications"))
        {
            registeredApps.SetValue(AppRegistryName, $@"Software\{AppRegistryName}\Capabilities");
        }

        NotifyShell();
        _logger.Info("파일 연결 등록 완료. 탐색기에서 '연결 프로그램'으로 선택하거나, 설정 > 앱 > 기본 앱에서 지정할 수 있습니다.");
    }

    public void UnregisterAll(AppConfig config)
    {
        using var classesRoot = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Classes", writable: true);

        Microsoft.Win32.Registry.CurrentUser.DeleteSubKeyTree($@"Software\{AppRegistryName}", throwOnMissingSubKey: false);
        Microsoft.Win32.Registry.CurrentUser.DeleteSubKeyTree($@"Software\{LegacyAppRegistryName}", throwOnMissingSubKey: false);

        using (var registeredApps = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\RegisteredApplications", writable: true))
        {
            registeredApps?.DeleteValue(AppRegistryName, throwOnMissingValue: false);
            registeredApps?.DeleteValue(LegacyAppRegistryName, throwOnMissingValue: false);
        }

        if (classesRoot is not null)
        {
            classesRoot.DeleteSubKeyTree($@"Applications\{ExecutableName}", throwOnMissingSubKey: false);
            classesRoot.DeleteSubKeyTree(@"Applications\Launcher.exe", throwOnMissingSubKey: false);

            var extensions = config.DocumentTypes.SelectMany(d => d.Extensions).Select(Normalize);
            foreach (var ext in extensions)
            {
                var progId = ProgIdFor(ext);
                using (var openWithProgIds = classesRoot.OpenSubKey($@"{ext}\OpenWithProgids", writable: true))
                {
                    openWithProgIds?.DeleteValue(progId, throwOnMissingValue: false);
                }

                classesRoot.DeleteSubKeyTree($@"{ext}\OpenWithList\{ExecutableName}", throwOnMissingSubKey: false);
                classesRoot.DeleteSubKeyTree(progId, throwOnMissingSubKey: false);
                _logger.Info($"확장자 등록 해제: {ext}");
            }
        }

        NotifyShell();
        _logger.Info("파일 연결 등록을 모두 해제했습니다.");
    }

    private void RegisterApplicationEntry(RegistryKey classesRoot, string exePath, IEnumerable<string> extensions)
    {
        using var appKey = OpenOrCreate(classesRoot, $@"Applications\{ExecutableName}");
        appKey.SetValue("FriendlyAppName", AppBrand.Name);

        using (var supportedTypes = OpenOrCreate(appKey, "SupportedTypes"))
        {
            foreach (var ext in extensions)
            {
                supportedTypes.SetValue(ext, string.Empty);
            }
        }

        using var shellOpenCommand = OpenOrCreate(classesRoot, $@"Applications\{ExecutableName}\shell\open\command");
        shellOpenCommand.SetValue(string.Empty, $"\"{exePath}\" \"%1\"");
    }

    private static void RegisterOpenWithEntry(RegistryKey classesRoot, string ext, string progId)
    {
        using var openWithProgIds = OpenOrCreate(classesRoot, $@"{ext}\OpenWithProgids");
        openWithProgIds.SetValue(progId, string.Empty);
    }

    private static void RegisterOpenWithExecutable(RegistryKey classesRoot, string ext)
    {
        using var _ = OpenOrCreate(classesRoot, $@"{ext}\OpenWithList\{ExecutableName}");
    }

    private static void RegisterProgId(RegistryKey classesRoot, string progId, string ext, string officeApp, string exePath)
    {
        using var progIdKey = OpenOrCreate(classesRoot, progId);
        progIdKey.SetValue(string.Empty, $"{officeApp} 문서 (Office Web)");

        using var iconKey = OpenOrCreate(classesRoot, $@"{progId}\DefaultIcon");
        iconKey.SetValue(string.Empty, $"{exePath},0");

        using var applicationKey = OpenOrCreate(classesRoot, $@"{progId}\Application");
        applicationKey.SetValue("ApplicationName", AppBrand.Name);
        applicationKey.SetValue("ApplicationDescription", AppBrand.Description);
        applicationKey.SetValue("ApplicationIcon", $"{exePath},0");

        using var commandKey = OpenOrCreate(classesRoot, $@"{progId}\shell\open\command");
        commandKey.SetValue(string.Empty, $"\"{exePath}\" \"%1\"");
    }

    private static RegistryKey OpenClassesRoot() =>
        Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Classes", writable: true)
            ?? throw new InvalidOperationException("HKCU\\Software\\Classes를 열 수 없습니다.");

    private static RegistryKey OpenOrCreate(RegistryKey root, string subKeyPath) =>
        root.CreateSubKey(subKeyPath, writable: true)
            ?? throw new InvalidOperationException($"레지스트리 키를 생성할 수 없습니다: {subKeyPath}");

    private static string ProgIdFor(string ext) => $"{ProgIdPrefix}{ext}";

    private static string Normalize(string ext) => ext.StartsWith('.') ? ext : "." + ext;

    private static void NotifyShell() => SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero);
}
