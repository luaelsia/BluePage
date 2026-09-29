using Microsoft365OfficeWebLauncher.Core;
using Microsoft365OfficeWebLauncher.Config;
using Microsoft365OfficeWebLauncher.Cloud;
using Microsoft365OfficeWebLauncher.OneDrive;
using System.Text.Json;

var testRoot = Path.Combine(Path.GetTempPath(), "BluePage-SmokeTests", Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(testRoot);

try
{
    var source = Path.Combine(AppContext.BaseDirectory, "sample.docx");
    var sample = Path.Combine(testRoot, "sample.docx");
    File.Copy(source, sample);

    var registry = new DeferredSyncRegistry(Path.Combine(testRoot, "deferred-sync"));
    registry.Defer(sample);
    Assert(registry.IsDeferred(sample), "동기화 보류 상태가 기록되지 않았습니다.");

    registry.Resume(sample);
    Assert(!registry.IsDeferred(sample), "문서 재열기 후 동기화 보류가 해제되지 않았습니다.");

    Assert(LocalFileStateMonitor.TryRead(sample, out var state) && state.Length > 0,
        "sample.docx 상태를 읽지 못했습니다.");
    File.Delete(sample);
    Assert(!LocalFileStateMonitor.TryRead(sample, out _),
        "삭제된 파일을 사용 가능한 파일로 잘못 감지했습니다.");

    var manifestEntry = new ManifestEntry
    {
        Provider = "Microsoft",
        DriveItemId = "ms-file",
        WebUrl = "https://microsoft.example/file",
        LastKnownLocalWriteUtc = DateTimeOffset.UtcNow,
        LastKnownRemoteModifiedUtc = DateTimeOffset.UtcNow
    };
    Assert(!manifestEntry.Activate("Google"), "처음 선택한 Google 원격 항목이 이미 있다고 잘못 판단했습니다.");
    manifestEntry.DriveItemId = "google-file";
    manifestEntry.WebUrl = "https://google.example/file";
    manifestEntry.SaveActiveRemote();
    Assert(manifestEntry.Activate("Microsoft") && manifestEntry.DriveItemId == "ms-file",
        "Microsoft 원격 항목을 공급자 전환 후 복원하지 못했습니다.");
    Assert(manifestEntry.Activate("Google") && manifestEntry.DriveItemId == "google-file",
        "Google 원격 항목을 공급자 전환 후 복원하지 못했습니다.");

    // 서로 다른 폴더의 동명 문서가 같은 원격 파일 이름을 쓰지 않아야 한다.
    var documentIdA = RemoteFileNaming.NewDocumentId();
    var documentIdB = RemoteFileNaming.NewDocumentId();
    Assert(documentIdA != documentIdB, "문서 ID가 매번 새로 발급되지 않았습니다.");
    var remoteNameA = RemoteFileNaming.Build(@"C:\FolderA\plan.docx", documentIdA);
    var remoteNameB = RemoteFileNaming.Build(@"C:\FolderB\plan.docx", documentIdB);
    Assert(remoteNameA != remoteNameB, "동명 문서가 같은 원격 파일 이름을 갖습니다.");
    Assert(remoteNameA.EndsWith(".docx", StringComparison.Ordinal), "원격 파일 이름이 확장자를 잃었습니다.");

    var freshEntry = new ManifestEntry();
    var issued = freshEntry.EnsureDocumentId();
    Assert(!string.IsNullOrWhiteSpace(issued), "문서 ID가 발급되지 않았습니다.");
    Assert(freshEntry.EnsureDocumentId() == issued, "이미 발급된 문서 ID가 다시 발급되었습니다.");

    var defaultConfigPath = Path.Combine(AppContext.BaseDirectory, "appsettings.default.json");
    var defaultConfig = JsonSerializer.Deserialize<AppConfig>(File.ReadAllText(defaultConfigPath))
        ?? throw new InvalidOperationException("기본 설정을 읽지 못했습니다.");
    var documentTypes = new DocumentTypeCatalog(defaultConfig);
    Assert(documentTypes.TryResolve("sample.csv", out var csvType) && csvType.OfficeApp == "Excel",
        "CSV가 Excel 문서 형식으로 등록되지 않았습니다.");
    Assert(documentTypes.TryResolve("sample.ods", out var odsType) &&
           odsType.Supports(CloudProvider.Microsoft) && odsType.Supports(CloudProvider.Google),
        "ODS가 양쪽 스프레드시트 서비스에 등록되지 않았습니다.");
    Assert(documentTypes.TryResolve("sample.odp", out var odpType) &&
           odpType.Supports(CloudProvider.Microsoft) && odpType.Supports(CloudProvider.Google),
        "ODP가 양쪽 프레젠테이션 서비스에 등록되지 않았습니다.");
    Assert(documentTypes.TryResolve("sample.xlsb", out var xlsbType) &&
           xlsbType.Supports(CloudProvider.Microsoft) && !xlsbType.Supports(CloudProvider.Google),
        "XLSB의 Microsoft 전용 제한이 적용되지 않았습니다.");
    Assert(documentTypes.TryResolve("sample.odt", out var odtType) &&
           !odtType.Supports(CloudProvider.Microsoft) && odtType.Supports(CloudProvider.Google),
        "ODT의 Google 전용 제한이 적용되지 않았습니다.");

    // --url: 보기 모드 주소 변환
    var editUrl = "https://contoso-my.sharepoint.com/personal/u/_layouts/15/Doc.aspx?sourcedoc=%7BA%7D&file=a.docx&action=default&mobileredirect=true";
    var viewUrl = AiAccess.ToViewUrl(editUrl, out var viewOnly);
    Assert(viewOnly && viewUrl.Contains("action=view") && !viewUrl.Contains("action=default"),
        "Doc.aspx 주소가 보기 모드로 바뀌지 않았습니다.");
    Assert(viewUrl.Contains("sourcedoc=%7BA%7D") && viewUrl.Contains("file=a.docx"),
        "보기 모드 변환에서 기존 파라미터가 사라졌습니다.");
    var shortUrl = "https://1drv.ms/w/c/abc/EXAMPLE";
    Assert(AiAccess.ToViewUrl(shortUrl, out var shortViewOnly) == shortUrl && !shortViewOnly,
        "변환할 수 없는 주소를 바꾸거나 보기 전용으로 잘못 표시했습니다.");

    // --url: 등록되지 않은 문서는 거부
    Assert(!AiAccess.IsRegistered(null), "매니페스트에 없는 문서를 등록된 문서로 판단했습니다.");
    Assert(!AiAccess.IsRegistered(new ManifestEntry()), "온라인 사본이 없는 항목을 등록된 문서로 판단했습니다.");
    Assert(AiAccess.IsRegistered(manifestEntry), "온라인 사본이 있는 항목을 등록되지 않은 문서로 판단했습니다.");

    // --url: 충돌이면 창 없이 건너뛴다
    Assert(new NonInteractiveConflictResolver().Resolve(new ConflictInfo(sample, DateTimeOffset.UtcNow, DateTimeOffset.UtcNow))
           == ConflictResolutionChoice.Skip, "창 없는 충돌 처리기가 Skip 이외의 동작을 골랐습니다.");

    // --url: JSON 형식 (성공이면 주소와 규칙, 실패면 주소 없이 오류 코드)
    var success = AiUrlResult.Success(sample, "Microsoft", SyncState.LocalOnlyChanged, editUrl);
    using (var json = JsonDocument.Parse(AiAccess.ToJson(success)))
    {
        var root = json.RootElement;
        Assert(root.GetProperty("ok").GetBoolean() && root.GetProperty("url").GetString() == viewUrl,
            "성공 JSON에 보기 모드 주소가 들어가지 않았습니다.");
        Assert(root.GetProperty("syncState").GetString() == "LocalOnlyChanged", "성공 JSON의 syncState가 잘못되었습니다.");
        Assert(root.GetProperty("rules").GetArrayLength() == AiAccess.Rules.Count, "성공 JSON에 규칙이 빠졌습니다.");
        Assert(!root.TryGetProperty("error", out _), "성공 JSON에 error가 들어갔습니다.");
    }
    var failure = AiUrlResult.Fail(sample, AiUrlError.Conflict, "충돌");
    Assert(failure.ExitCode == 3, "충돌 종료 코드가 3이 아닙니다.");
    using (var json = JsonDocument.Parse(AiAccess.ToJson(failure)))
    {
        var root = json.RootElement;
        Assert(!root.GetProperty("ok").GetBoolean() && root.GetProperty("error").GetString() == "conflict",
            "실패 JSON의 오류 코드가 잘못되었습니다.");
        Assert(!root.TryGetProperty("url", out _), "실패 JSON에 주소가 들어갔습니다.");
        Assert(!root.TryGetProperty("ExitCode", out _), "JSON에 내부용 ExitCode가 들어갔습니다.");
        Assert(root.GetProperty("rules").GetArrayLength() > 0, "실패 JSON에 규칙이 빠졌습니다.");
    }

    Assert(AiAccess.ReadGuide().Contains("--url"), "AI 지침 리소스를 읽지 못했습니다.");

    Console.WriteLine("PASS: 동기화 보류/재개, 삭제 감지, 클라우드 공급자 전환, 원격 파일 이름 충돌 방지, 확장자별 서비스 제한, AI 주소 요청 테스트");
    return 0;
}
finally
{
    if (Directory.Exists(testRoot)) Directory.Delete(testRoot, recursive: true);
}

static void Assert(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}
