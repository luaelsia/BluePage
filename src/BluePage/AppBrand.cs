namespace Microsoft365OfficeWebLauncher;

internal static class AppBrand
{
    public const string Name = "Blue Page";
    public const string Publisher = "MiniWhaleLabs";
    public const string ContactEmail = "miniwhalelabs@gmail.com";
    public const string Description = "로컬 Office 문서를 Microsoft 365 또는 Google Workspace에서 여는 문서 런처";

    public static string Version
    {
        get
        {
            var version = typeof(AppBrand).Assembly.GetName().Version;
            return version is null
                ? "알 수 없음"
                : $"{version.Major}.{version.Minor}.{Math.Max(0, version.Build)}";
        }
    }
}
