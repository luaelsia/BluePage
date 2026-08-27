using Microsoft.Graph.Models;
using Microsoft365OfficeWebLauncher.Core;
using Microsoft365OfficeWebLauncher.Logging;

namespace Microsoft365OfficeWebLauncher.OneDrive;

public enum SyncState
{
    /// <summary>매니페스트에 없던 파일 — 최초 업로드.</summary>
    NewUpload,
    /// <summary>로컬/온라인 모두 마지막 동기화 이후 변경 없음.</summary>
    NoChange,
    /// <summary>Office Web에서만 수정됨(또는 충돌 해결로 온라인을 선택) — 온라인 사본을 로컬로 내려받음.</summary>
    RemoteOnlyChanged,
    /// <summary>로컬 파일만 수정됨(또는 충돌 해결로 로컬을 선택) — 로컬 내용을 온라인으로 다시 올림.</summary>
    LocalOnlyChanged,
    /// <summary>양쪽 모두 변경됐고, 사용자가 "사본 생성"을 선택함 — 충돌 사본을 별도 생성.</summary>
    Conflict
}

public sealed record SyncResult(SyncState State, string WebUrl, string? ConflictCopyPath);

/// <summary>
/// 계획서 8장(Sync-back 설계)의 판정 로직: 로컬 mtime과 온라인 lastModifiedDateTime을
/// 마지막으로 알려진 동기화 시각과 비교해 상태를 분류하고 그에 맞는 동작을 수행한다.
/// 로컬/온라인이 동시에 변경된 충돌 상황은 IConflictResolver에게 물어 사용자가 고른 방식대로 처리한다.
/// </summary>
public sealed class SyncCoordinator
{
    // 파일시스템/네트워크 타임스탬프 해상도 차이로 인한 오탐을 막기 위한 여유 시간
    private static readonly TimeSpan ChangeTolerance = TimeSpan.FromSeconds(2);

    private readonly OneDriveUploadService _uploadService;
    private readonly UploadManifest _manifest;
    private readonly IConflictResolver _conflictResolver;
    private readonly FileLogger _logger;
    private ISyncActivityReporter _activityReporter = NullSyncActivityReporter.Instance;

    public SyncCoordinator(OneDriveUploadService uploadService, UploadManifest manifest, IConflictResolver conflictResolver, FileLogger logger)
    {
        _uploadService = uploadService;
        _manifest = manifest;
        _conflictResolver = conflictResolver;
        _logger = logger;
    }

    /// <summary>
    /// GUI 프로세스에서만 실제 토스트 알림 구현을 붙인다(생성 순서 문제로 생성자 주입 대신
    /// 나중에 붙이는 방식 — LauncherForm이 자기 자신의 UI 컴포넌트를 만든 뒤 연결한다).
    /// 헤드리스 오픈/CLI는 이 메서드를 호출하지 않으므로 계속 NullSyncActivityReporter로 동작한다.
    /// </summary>
    public void AttachActivityReporter(ISyncActivityReporter reporter) => _activityReporter = reporter;

    /// <summary>
    /// 파일을 열기 직전에 호출: 필요한 업로드/다운로드를 수행하고 열어야 할 webUrl을 반환한다.
    /// 충돌이면 백그라운드 폴링을 포함해 항상 사용자에게 물어본다(진짜 충돌은 자주 발생하지 않고,
    /// 조용히 임의로 처리하기보다는 사용자가 직접 고르는 게 맞다는 판단). 과거에는 프로세스 간 매니페스트가
    /// 최신화되지 않아 가짜 충돌이 반복 발생했었는데, 이는 호출 전 UploadManifest.Reload()로 해결했다.
    /// </summary>
    public async Task<SyncResult> PrepareForOpenAsync(string localFilePath, CancellationToken ct)
    {
        var entry = _manifest.Get(localFilePath);

        if (entry is null)
        {
            _activityReporter.ReportStarted(localFilePath);
            try
            {
                var created = await _uploadService.CreateInAppFolderAsync(localFilePath, ct);
                var newEntry = new ManifestEntry
                {
                    DriveItemId = created.Id!,
                    WebUrl = created.WebUrl ?? string.Empty,
                    LastKnownLocalWriteUtc = File.GetLastWriteTimeUtc(localFilePath),
                    LastKnownRemoteModifiedUtc = created.LastModifiedDateTime ?? DateTimeOffset.UtcNow
                };
                _manifest.Set(localFilePath, newEntry);
                _manifest.Save();
                _logger.Info($"신규 업로드 완료: {localFilePath} -> {newEntry.WebUrl}");
                _activityReporter.ReportCompleted(localFilePath, true);
                return new SyncResult(SyncState.NewUpload, newEntry.WebUrl, null);
            }
            catch
            {
                _activityReporter.ReportCompleted(localFilePath, false);
                throw;
            }
        }

        var localWriteUtc = File.GetLastWriteTimeUtc(localFilePath);
        var remoteMetadata = await _uploadService.GetMetadataAsync(entry.DriveItemId, ct);
        var remoteModifiedUtc = remoteMetadata.LastModifiedDateTime ?? entry.LastKnownRemoteModifiedUtc;

        var localChanged = localWriteUtc - entry.LastKnownLocalWriteUtc > ChangeTolerance;
        var remoteChanged = remoteModifiedUtc - entry.LastKnownRemoteModifiedUtc > ChangeTolerance;

        if (!localChanged && !remoteChanged)
        {
            _logger.Info($"변경 없음: {localFilePath}");
            return new SyncResult(SyncState.NoChange, entry.WebUrl, null);
        }

        if (remoteChanged && !localChanged)
        {
            return await DownloadOnlineOverLocalAsync(localFilePath, entry, remoteMetadata, remoteModifiedUtc, ct);
        }

        if (localChanged && !remoteChanged)
        {
            return await UploadLocalOverOnlineAsync(localFilePath, entry, localWriteUtc, ct);
        }

        // 충돌: 양쪽 다 변경됨 — 항상 사용자에게 어떻게 처리할지 묻는다(기본값: 사본 생성).
        var choice = _conflictResolver.Resolve(new ConflictInfo(localFilePath, localWriteUtc, remoteModifiedUtc));
        _logger.Info($"충돌 감지: {localFilePath} — 사용자 선택: {choice}");

        switch (choice)
        {
            case ConflictResolutionChoice.KeepLocal:
                return await UploadLocalOverOnlineAsync(localFilePath, entry, localWriteUtc, ct);

            case ConflictResolutionChoice.KeepRemote:
                return await DownloadOnlineOverLocalAsync(localFilePath, entry, remoteMetadata, remoteModifiedUtc, ct);

            case ConflictResolutionChoice.KeepNewer:
                return localWriteUtc >= remoteModifiedUtc
                    ? await UploadLocalOverOnlineAsync(localFilePath, entry, localWriteUtc, ct)
                    : await DownloadOnlineOverLocalAsync(localFilePath, entry, remoteMetadata, remoteModifiedUtc, ct);

            case ConflictResolutionChoice.CreateCopy:
            default:
                return await CreateConflictCopyAsync(localFilePath, entry, remoteMetadata, localWriteUtc, remoteModifiedUtc, ct);
        }
    }

    /// <summary>
    /// 동기화 검토 창용: 업로드/다운로드 없이 상태만 조회한다.
    /// 매니페스트에 없는(추적 안 하는) 파일은 대상이 아니므로 호출하지 않는다.
    /// </summary>
    public async Task<SyncDetection> DetectAsync(string localFilePath, CancellationToken ct)
    {
        var entry = _manifest.Get(localFilePath)
            ?? throw new InvalidOperationException($"추적 중이 아닌 파일입니다: {localFilePath}");

        var localWriteUtc = File.GetLastWriteTimeUtc(localFilePath);
        var remoteMetadata = await _uploadService.GetMetadataAsync(entry.DriveItemId, ct);
        var remoteModifiedUtc = remoteMetadata.LastModifiedDateTime ?? entry.LastKnownRemoteModifiedUtc;

        var localChanged = localWriteUtc - entry.LastKnownLocalWriteUtc > ChangeTolerance;
        var remoteChanged = remoteModifiedUtc - entry.LastKnownRemoteModifiedUtc > ChangeTolerance;

        var state = (localChanged, remoteChanged) switch
        {
            (false, false) => SyncState.NoChange,
            (false, true) => SyncState.RemoteOnlyChanged,
            (true, false) => SyncState.LocalOnlyChanged,
            (true, true) => SyncState.Conflict
        };

        var lastSyncedUtc = entry.LastKnownLocalWriteUtc > entry.LastKnownRemoteModifiedUtc
            ? entry.LastKnownLocalWriteUtc
            : entry.LastKnownRemoteModifiedUtc;

        return new SyncDetection(localFilePath, state, localWriteUtc, remoteModifiedUtc, lastSyncedUtc);
    }

    /// <summary>동기화 검토 창용: 사용자가 고른 동작 하나만 실행한다(재조회 후 적용, Skip은 호출하지 않는 전제).</summary>
    public async Task<SyncResult> ApplyAsync(string localFilePath, SyncAction action, CancellationToken ct)
    {
        var entry = _manifest.Get(localFilePath)
            ?? throw new InvalidOperationException($"추적 중이 아닌 파일입니다: {localFilePath}");

        var localWriteUtc = File.GetLastWriteTimeUtc(localFilePath);
        var remoteMetadata = await _uploadService.GetMetadataAsync(entry.DriveItemId, ct);
        var remoteModifiedUtc = remoteMetadata.LastModifiedDateTime ?? entry.LastKnownRemoteModifiedUtc;

        return action switch
        {
            SyncAction.PullRemoteToLocal => await DownloadOnlineOverLocalAsync(localFilePath, entry, remoteMetadata, remoteModifiedUtc, ct),
            SyncAction.PushLocalToRemote => await UploadLocalOverOnlineAsync(localFilePath, entry, localWriteUtc, ct),
            SyncAction.CreateConflictCopy => await CreateConflictCopyAsync(localFilePath, entry, remoteMetadata, localWriteUtc, remoteModifiedUtc, ct),
            _ => throw new ArgumentOutOfRangeException(nameof(action), action, "Skip은 실행 대상이 아닙니다.")
        };
    }

    private async Task<SyncResult> UploadLocalOverOnlineAsync(string localFilePath, ManifestEntry entry, DateTimeOffset localWriteUtc, CancellationToken ct)
    {
        _activityReporter.ReportStarted(localFilePath);
        try
        {
            var updated = await _uploadService.UpdateContentAsync(entry.DriveItemId, localFilePath, ct);
            entry.LastKnownLocalWriteUtc = localWriteUtc;
            entry.LastKnownRemoteModifiedUtc = updated.LastModifiedDateTime ?? DateTimeOffset.UtcNow;
            entry.WebUrl = updated.WebUrl ?? entry.WebUrl;
            _manifest.Set(localFilePath, entry);
            _manifest.Save();
            _logger.Info($"로컬 내용을 온라인으로 업로드: {localFilePath}");
            _activityReporter.ReportCompleted(localFilePath, true);
            return new SyncResult(SyncState.LocalOnlyChanged, entry.WebUrl, null);
        }
        catch
        {
            _activityReporter.ReportCompleted(localFilePath, false);
            throw;
        }
    }

    private async Task<SyncResult> DownloadOnlineOverLocalAsync(string localFilePath, ManifestEntry entry, DriveItem remoteMetadata, DateTimeOffset remoteModifiedUtc, CancellationToken ct)
    {
        _activityReporter.ReportStarted(localFilePath);
        try
        {
            await _uploadService.DownloadContentAsync(entry.DriveItemId, localFilePath, ct);
            entry.LastKnownLocalWriteUtc = File.GetLastWriteTimeUtc(localFilePath);
            entry.LastKnownRemoteModifiedUtc = remoteModifiedUtc;
            entry.WebUrl = remoteMetadata.WebUrl ?? entry.WebUrl;
            _manifest.Set(localFilePath, entry);
            _manifest.Save();
            _logger.Info($"온라인 내용을 로컬로 반영: {localFilePath}");
            _activityReporter.ReportCompleted(localFilePath, true);
            return new SyncResult(SyncState.RemoteOnlyChanged, entry.WebUrl, null);
        }
        catch
        {
            _activityReporter.ReportCompleted(localFilePath, false);
            throw;
        }
    }

    private async Task<SyncResult> CreateConflictCopyAsync(
        string localFilePath, ManifestEntry entry, DriveItem remoteMetadata,
        DateTimeOffset localWriteUtc, DateTimeOffset remoteModifiedUtc, CancellationToken ct)
    {
        _activityReporter.ReportStarted(localFilePath);
        try
        {
            var conflictPath = BuildConflictCopyPath(localFilePath);
            await _uploadService.DownloadContentAsync(entry.DriveItemId, conflictPath, ct);
            entry.LastKnownLocalWriteUtc = localWriteUtc;
            entry.LastKnownRemoteModifiedUtc = remoteModifiedUtc;
            entry.WebUrl = remoteMetadata.WebUrl ?? entry.WebUrl;
            _manifest.Set(localFilePath, entry);
            _manifest.Save();
            _logger.Info($"충돌 사본 생성: {conflictPath}");
            _activityReporter.ReportCompleted(localFilePath, true);
            return new SyncResult(SyncState.Conflict, entry.WebUrl, conflictPath);
        }
        catch
        {
            _activityReporter.ReportCompleted(localFilePath, false);
            throw;
        }
    }

    private static string BuildConflictCopyPath(string localFilePath)
    {
        var dir = Path.GetDirectoryName(localFilePath) ?? ".";
        var name = Path.GetFileNameWithoutExtension(localFilePath);
        var ext = Path.GetExtension(localFilePath);
        var timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        return Path.Combine(dir, $"{name} (Web 충돌사본, {timestamp}){ext}");
    }
}
