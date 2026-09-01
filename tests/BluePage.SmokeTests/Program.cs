using Microsoft365OfficeWebLauncher.Core;
using Microsoft365OfficeWebLauncher.OneDrive;

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

    Console.WriteLine("PASS: 동기화 보류/재개, 삭제 감지, 클라우드 공급자 전환 테스트");
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
