using System.Reflection;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft365OfficeWebLauncher.OneDrive;

namespace Microsoft365OfficeWebLauncher.Core;

/// <summary>--url 실패 사유. 종료 코드와 JSON의 error 값으로 함께 쓴다.</summary>
public enum AiUrlError
{
    None = 0,
    Failed = 1,
    NotRegistered = 2,
    Conflict = 3,
    Locked = 4,
    AuthRequired = 5,
    NotFound = 6,
    Unsupported = 7
}

/// <summary>--url 결과. 그대로 JSON으로 직렬화해 stdout에 쓴다.</summary>
public sealed class AiUrlResult
{
    [JsonPropertyName("ok")]
    public bool Ok { get; init; }

    [JsonPropertyName("error")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Error { get; init; }

    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }

    [JsonPropertyName("path")]
    public string Path { get; init; } = string.Empty;

    [JsonPropertyName("provider")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Provider { get; init; }

    [JsonPropertyName("syncState")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SyncStateName { get; init; }

    [JsonPropertyName("url")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Url { get; init; }

    [JsonPropertyName("viewOnly")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? ViewOnly { get; init; }

    [JsonPropertyName("note")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; init; }

    [JsonPropertyName("rules")]
    public IReadOnlyList<string> Rules { get; init; } = AiAccess.Rules;

    [JsonPropertyName("guide")]
    public string Guide { get; init; } = AiAccess.GuideCommand;

    [JsonIgnore]
    public AiUrlError ErrorCode { get; init; }

    [JsonIgnore]
    public int ExitCode => (int)ErrorCode;

    public static AiUrlResult Success(string path, string provider, SyncState state, string webUrl)
    {
        var viewUrl = AiAccess.ToViewUrl(webUrl, out var viewOnly);
        return new AiUrlResult
        {
            Ok = true,
            Path = path,
            Provider = provider,
            SyncStateName = state.ToString(),
            Url = viewUrl,
            ViewOnly = viewOnly,
            Note = state == SyncState.RemoteOnlyChanged
                ? "온라인 사본이 더 새로워서 로컬 파일을 온라인 내용으로 덮어썼습니다. 로컬 파일을 다시 읽고 작업하세요."
                : null
        };
    }

    public static AiUrlResult Fail(string path, AiUrlError error, string message) => new()
    {
        Ok = false,
        Path = path,
        ErrorCode = error,
        Error = AiAccess.ErrorName(error),
        Message = message
    };
}

/// <summary>
/// AI가 BluePage를 통해 문서의 온라인 사본을 볼 때 필요한 공용 규칙, 보기 모드 주소 변환, 지침 원문.
/// 규칙은 --url 출력마다 함께 내보내서, 주소를 받은 AI가 규칙을 못 보고 지나칠 수 없게 한다.
/// </summary>
public static class AiAccess
{
    public const string GuideCommand = "BluePage.exe --ai-guide";

    public static readonly IReadOnlyList<string> Rules =
    [
        "이 주소는 보기 전용이다. 웹 Office/Google 문서에서 내용을 편집하지 않는다.",
        "브라우저 탭이 열려 있는 동안 로컬 파일을 수정하지 않는다. 확인이 끝나면 탭을 닫고 로컬을 수정한다.",
        "로컬 파일을 수정한 뒤에는 다시 --url을 실행해 동기화하고 새 주소로 확인한다. 예전 탭을 새로고침하지 않는다.",
        "BluePage 데이터 폴더(%LOCALAPPDATA%\\Microsoft365OfficeWebLauncher)를 읽지 않는다. 토큰이 들어 있다.",
        "ok가 false이면 주소를 추측하거나 다른 경로로 문서를 찾지 말고, error와 message를 사용자에게 알린다.",
        "로그인 화면이 나오면 로그인하지 말고 사용자에게 로그인을 요청한다."
    ];

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    public static string ToJson(AiUrlResult result) => JsonSerializer.Serialize(result, JsonOptions);

    public static string ErrorName(AiUrlError error) => error switch
    {
        AiUrlError.NotRegistered => "not_registered",
        AiUrlError.Conflict => "conflict",
        AiUrlError.Locked => "locked",
        AiUrlError.AuthRequired => "auth_required",
        AiUrlError.NotFound => "not_found",
        AiUrlError.Unsupported => "unsupported",
        AiUrlError.None => "none",
        _ => "failed"
    };

    /// <summary>
    /// 매니페스트에 온라인 사본이 기록된 문서만 허용한다. 사용자가 BluePage로 한 번도 연 적 없는 파일을
    /// AI가 임의로 클라우드에 새로 올리지 못하게 막는 조건이다.
    /// </summary>
    public static bool IsRegistered(ManifestEntry? entry) =>
        entry is not null && !string.IsNullOrWhiteSpace(entry.DriveItemId) && !string.IsNullOrWhiteSpace(entry.WebUrl);

    /// <summary>
    /// Office Online의 Doc.aspx/WopiFrame.aspx 주소면 action=view로 바꿔 보기 모드로 열리게 한다.
    /// 그 밖의 형식(개인 OneDrive 단축 주소, Google 문서 등)은 바꾸지 않고 viewOnly=false로 알린다.
    /// </summary>
    public static string ToViewUrl(string webUrl, out bool viewOnly)
    {
        viewOnly = false;
        if (!Uri.TryCreate(webUrl, UriKind.Absolute, out var uri))
        {
            return webUrl;
        }

        var page = uri.AbsolutePath;
        if (!page.EndsWith("/Doc.aspx", StringComparison.OrdinalIgnoreCase) &&
            !page.EndsWith("/WopiFrame.aspx", StringComparison.OrdinalIgnoreCase))
        {
            return webUrl;
        }

        var parts = uri.Query.TrimStart('?')
            .Split('&', StringSplitOptions.RemoveEmptyEntries)
            .Where(part => !part.StartsWith("action=", StringComparison.OrdinalIgnoreCase))
            .Append("action=view");

        viewOnly = true;
        return new UriBuilder(uri) { Query = string.Join("&", parts) }.Uri.AbsoluteUri;
    }

    /// <summary>exe에 포함된 AI-GUIDE.md 원문. 설치된 버전과 지침이 항상 일치하도록 리소스로 넣었다.</summary>
    public static string ReadGuide()
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("BluePage.AI-GUIDE.md")
            ?? throw new InvalidOperationException("AI-GUIDE.md 리소스를 찾을 수 없습니다.");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
