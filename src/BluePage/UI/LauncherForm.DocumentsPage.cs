using Microsoft365OfficeWebLauncher.Cloud;
using Microsoft365OfficeWebLauncher.Config;
using Microsoft365OfficeWebLauncher.Core;
using Microsoft365OfficeWebLauncher.UI.Controls;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>문서 연결 화면: 기본 열기 방식과 확장자별 열기 방식.</summary>
public sealed partial class LauncherForm
{
    private SelectBox _providerSelect = null!;
    private readonly Dictionary<string, SelectBox> _extensionProviderSelects = new(StringComparer.OrdinalIgnoreCase);

    private Control BuildDocumentsPage()
    {
        var page = new PageStack("문서 연결", "탐색기에서 연 문서를 어느 서비스의 웹 편집기로 열지 정합니다");

        _providerSelect = new SelectBox("열 때마다 선택", "Microsoft 365", "Google Workspace (테스트)");
        _providerSelect.SelectedIndex = _config.PreferredCloudProvider.ToLowerInvariant() switch
        {
            "microsoft" => 1,
            "google" => 2,
            _ => 0
        };
        _providerSelect.SelectedIndexChanged += (_, _) =>
        {
            _config.PreferredCloudProvider = _providerSelect.SelectedIndex switch { 1 => "Microsoft", 2 => "Google", _ => "Ask" };
            ConfigLoader.Save(_config);
        };
        page.Add(BuildRowsCard(new Control[]
        {
            BuildSettingRow("기본 열기 방식", "아래 확장자별 열기 방식에서 따로 정하지 않은 형식은 이 방식으로 엽니다.", _providerSelect)
        }, SettingRowHeight));

        var catalog = new DocumentTypeCatalog(_config);
        var definitions = _config.DocumentTypes
            .SelectMany(type => type.Extensions.Select(NormalizeExtension))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(extension =>
            {
                catalog.TryResolve(extension, out var type);
                return (Extension: extension, Type: type);
            })
            .Where(item => item.Type is not null)
            .ToList();

        // Word, Excel, PowerPoint, Google 순서로 묶는다. Google 묶음은 Google에서만 열 수 있는 형식이다.
        var groups = definitions
            .GroupBy(item => GroupNameOf(item.Type!.OfficeApp))
            .OrderBy(group => GroupOrder(group.Key))
            .ThenBy(group => group.Key, StringComparer.CurrentCulture);
        foreach (var group in groups)
        {
            page.AddSection(group.Key);
            // 서비스를 고를 수 없는 전용 형식(예: Microsoft 365 전용)은 묶음의 맨 아래로 보낸다.
            var rows = group
                .OrderBy(item => item.Type!.SupportedProviders.Count == 1)
                .ThenBy(item => item.Extension, StringComparer.OrdinalIgnoreCase)
                .Select(item => BuildExtensionRow(item.Extension, item.Type!, showApp: group.Key == "Google"))
                .ToList();
            page.Add(BuildRowsCard(rows, UiDraw.S(50)));
        }

        return page.Root;
    }

    private static string GroupNameOf(string officeApp) =>
        officeApp.StartsWith("Google", StringComparison.OrdinalIgnoreCase) ? "Google" : officeApp;

    private static int GroupOrder(string group) => group switch
    {
        "Word" => 0,
        "Excel" => 1,
        "PowerPoint" => 2,
        "Google" => 3,
        _ => 4
    };

    /// <summary>showApp이 false면 묶음 제목과 같은 앱 이름을 되풀이하지 않는다.</summary>
    private Control BuildExtensionRow(string extension, OfficeAppDefinition type, bool showApp)
    {
        var row = new TableLayoutPanel { ColumnCount = 3, RowCount = 1, Margin = new Padding(0) };
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, UiDraw.S(96)));
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        row.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        row.Controls.Add(new Label { Text = extension, AutoSize = true, Font = CardTitleFont, Anchor = AnchorStyles.Left, Margin = new Padding(0) }, 0, 0);
        row.Controls.Add(new Label { Text = showApp ? type.OfficeApp : string.Empty, AutoSize = true, Tag = ThemeApplier.SecondaryTag, Anchor = AnchorStyles.Left, Margin = new Padding(0) }, 1, 0);

        if (type.SupportedProviders.Count == 1)
        {
            // 한 서비스에서만 열 수 있는 형식은 바꿀 수 없으므로 선택 상자 대신 배지로 보여 준다.
            _config.DocumentProviderPreferences.Remove(extension);
            var badge = new StatusBadge { Anchor = AnchorStyles.Right };
            badge.Set(BadgeKind.Neutral, $"{type.SupportedProviders.Single().DisplayName()} 전용");
            row.Controls.Add(badge, 2, 0);
            return row;
        }

        var select = new SelectBox("기본값 따름", "열 때마다 선택", "Microsoft 365", "Google Workspace (테스트)") { Anchor = AnchorStyles.Right };
        var stored = _config.DocumentProviderPreferences.TryGetValue(extension, out var value) ? value : "Default";
        select.SelectedIndex = stored.ToLowerInvariant() switch
        {
            "ask" => 1,
            "microsoft" => 2,
            "google" => 3,
            _ => 0
        };
        select.SelectedIndexChanged += (_, _) =>
        {
            var preference = select.SelectedIndex switch { 1 => "Ask", 2 => "Microsoft", 3 => "Google", _ => "Default" };
            if (preference == "Default") _config.DocumentProviderPreferences.Remove(extension);
            else _config.DocumentProviderPreferences[extension] = preference;
            ConfigLoader.Save(_config);
        };
        _extensionProviderSelects[extension] = select;
        row.Controls.Add(select, 2, 0);
        return row;
    }

    private static string NormalizeExtension(string extension) =>
        (extension.StartsWith('.') ? extension : "." + extension).ToLowerInvariant();
}
