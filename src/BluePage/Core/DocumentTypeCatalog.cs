using Microsoft365OfficeWebLauncher.Config;
using Microsoft365OfficeWebLauncher.Cloud;

namespace Microsoft365OfficeWebLauncher.Core;

/// <summary>
/// 확장자 → Office 앱 매핑. appsettings의 documentTypes 배열에 항목을 추가하는 것만으로
/// 새로운 확장자/Office 앱(Outlook, Visio, OneNote 등)을 코드 변경 없이 지원할 수 있다.
/// </summary>
public sealed class DocumentTypeCatalog
{
    private readonly Dictionary<string, OfficeAppDefinition> _byExtension;

    public DocumentTypeCatalog(AppConfig config)
    {
        _byExtension = new Dictionary<string, OfficeAppDefinition>(StringComparer.OrdinalIgnoreCase);

        foreach (var docType in config.DocumentTypes)
        {
            var providers = docType.SupportedProviders.Count == 0
                ? new HashSet<CloudProvider> { CloudProvider.Microsoft, CloudProvider.Google }
                : docType.SupportedProviders
                    .Select(value => CloudProviderNames.TryParse(value, out var provider) ? provider : (CloudProvider?)null)
                    .Where(provider => provider.HasValue)
                    .Select(provider => provider!.Value)
                    .ToHashSet();
            var definition = new OfficeAppDefinition(docType.OfficeApp, docType.Extensions, providers);
            foreach (var ext in docType.Extensions)
            {
                _byExtension[Normalize(ext)] = definition;
            }
        }
    }

    public bool TryResolve(string filePath, out OfficeAppDefinition definition)
    {
        var ext = Path.GetExtension(filePath);
        return _byExtension.TryGetValue(ext, out definition!);
    }

    public IReadOnlyCollection<string> SupportedExtensions => _byExtension.Keys;

    private static string Normalize(string ext) => ext.StartsWith('.') ? ext : "." + ext;
}
