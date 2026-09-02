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

    Console.WriteLine("PASS: 동기화 보류/재개, 삭제 감지, 클라우드 공급자 전환, 확장자별 서비스 제한 테스트");
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
