using Microsoft365OfficeWebLauncher.Cloud;

namespace Microsoft365OfficeWebLauncher.Core;

/// <summary>확장자 하나에 대응하는 웹 Office 앱과 지원 서비스 정보.</summary>
public sealed record OfficeAppDefinition(
    string OfficeApp,
    IReadOnlyList<string> Extensions,
    IReadOnlySet<CloudProvider> SupportedProviders)
{
    public bool Supports(CloudProvider provider) => SupportedProviders.Contains(provider);
}
