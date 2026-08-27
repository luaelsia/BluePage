using System.Text.Json.Serialization;

namespace Microsoft365OfficeWebLauncher.Config;

public sealed class AppConfig
{
    [JsonPropertyName("auth")]
    public AuthConfig Auth { get; set; } = new();

    [JsonPropertyName("sharedPcMode")]
    public bool SharedPcMode { get; set; }

    /// <summary>백그라운드 자동 동기화 주기(초). 1~600(10분) 범위로 GUI에서 조절 가능. 기본값 180(3분).</summary>
    [JsonPropertyName("backgroundSyncIntervalSeconds")]
    public int BackgroundSyncIntervalSeconds { get; set; } = 180;

    /// <summary>동기화 중/완료를 화면 우하단에 토스트로 알릴지 여부. GUI 체크박스로 조절 가능.</summary>
    [JsonPropertyName("showSyncToast")]
    public bool ShowSyncToast { get; set; } = true;

    /// <summary>"System"(Windows 설정 따라가기) / "Light" / "Dark". GUI 드롭다운으로 조절 가능.</summary>
    [JsonPropertyName("theme")]
    public string Theme { get; set; } = "System";

    [JsonPropertyName("appFolderDisplayName")]
    public string AppFolderDisplayName { get; set; } = "Blue Page";

    [JsonPropertyName("documentTypes")]
    public List<DocumentTypeConfig> DocumentTypes { get; set; } = new();

    [JsonPropertyName("logging")]
    public LoggingConfig Logging { get; set; } = new();
}

public sealed class AuthConfig
{
    [JsonPropertyName("clientId")]
    public string ClientId { get; set; } = string.Empty;

    [JsonPropertyName("authority")]
    public string Authority { get; set; } = "https://login.microsoftonline.com/common";

    [JsonPropertyName("scopes")]
    public List<string> Scopes { get; set; } = new() { "Files.ReadWrite.AppFolder" };
}

public sealed class DocumentTypeConfig
{
    [JsonPropertyName("extensions")]
    public List<string> Extensions { get; set; } = new();

    [JsonPropertyName("officeApp")]
    public string OfficeApp { get; set; } = string.Empty;
}

public sealed class LoggingConfig
{
    [JsonPropertyName("level")]
    public string Level { get; set; } = "Info";

    [JsonPropertyName("retainDays")]
    public int RetainDays { get; set; } = 14;
}
