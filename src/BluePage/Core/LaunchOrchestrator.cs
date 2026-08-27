using System.Diagnostics;
using Microsoft365OfficeWebLauncher.Logging;
using Microsoft365OfficeWebLauncher.OneDrive;

namespace Microsoft365OfficeWebLauncher.Core;

/// <summary>확장자 판별 → 동기화(업로드/다운로드) → 기본 브라우저로 webUrl 열기를 조율한다.</summary>
public sealed class LaunchOrchestrator
{
    private readonly DocumentTypeCatalog _catalog;
    private readonly SyncCoordinator _syncCoordinator;
    private readonly UploadManifest _manifest;
    private readonly FileLogger _logger;

    public LaunchOrchestrator(
        DocumentTypeCatalog catalog,
        SyncCoordinator syncCoordinator,
        UploadManifest manifest,
        FileLogger logger)
    {
        _catalog = catalog;
        _syncCoordinator = syncCoordinator;
        _manifest = manifest;
        _logger = logger;
    }

    public async Task<int> OpenAsync(string filePath, CancellationToken ct)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath))
        {
            _logger.Error($"파일을 찾을 수 없습니다: {fullPath}");
            ShowError($"파일을 찾을 수 없습니다:\n{fullPath}");
            return 1;
        }

        if (!_catalog.TryResolve(fullPath, out var appDefinition))
        {
            _logger.Error($"지원하지 않는 확장자입니다: {Path.GetExtension(fullPath)}");
            ShowError($"지원하지 않는 파일 형식입니다: {Path.GetExtension(fullPath)}\n" +
                      "appsettings.json의 documentTypes에 추가하면 지원할 수 있습니다.");
            return 1;
        }

        _logger.Info($"열기 시작: {fullPath} ({appDefinition.OfficeApp})");

        SyncResult result;
        try
        {
            result = await _syncCoordinator.PrepareForOpenAsync(fullPath, ct);
        }
        catch (Exception ex)
        {
            _logger.Error($"동기화 실패: {fullPath}", ex);

            // 이전에 이미 업로드된 적 있는 파일이면, 반영은 실패했더라도 알고 있는 온라인 사본을 그대로 열어준다
            // (동기화 실패 = 문서를 아예 못 여는 것보다는 낫다).
            var existingEntry = _manifest.Get(fullPath);
            if (existingEntry is not null && !string.IsNullOrEmpty(existingEntry.WebUrl))
            {
                var reason = GraphErrorHelper.IsResourceLocked(ex)
                    ? "이 문서가 지금 Office Web에서 편집 중이라 로컬 변경사항을 온라인에 반영하지 못했습니다."
                    : "온라인 동기화에 실패했습니다(자세한 내용은 로그 참고).";
                ShowInfo($"{reason}\n일단 온라인의 최신 버전을 그대로 엽니다.");

                OpenInBrowser(existingEntry.WebUrl);
                _logger.Info($"동기화 실패 후 기존 온라인 사본을 열었습니다: {existingEntry.WebUrl}");
                return 0;
            }

            ShowError($"동기화에 실패해 문서를 열 수 없습니다.\n\n{ex.Message}");
            return 1;
        }

        if (result.State == SyncState.Conflict)
        {
            ShowInfo(
                "로컬 파일과 Office Web의 온라인 사본이 모두 변경되어 자동으로 병합할 수 없습니다.\n" +
                $"온라인 사본을 아래 경로에 별도로 저장했습니다. 두 파일을 확인 후 직접 병합해 주세요.\n\n{result.ConflictCopyPath}");
        }

        OpenInBrowser(result.WebUrl);
        _logger.Info($"브라우저로 열기 완료: {result.WebUrl} (상태: {result.State})");
        return 0;
    }

    public async Task<int> SyncOneAsync(string filePath, CancellationToken ct)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath))
        {
            _logger.Error($"동기화 대상 파일을 찾을 수 없습니다: {fullPath}");
            ShowError($"파일을 찾을 수 없습니다:\n{fullPath}");
            return 1;
        }

        var result = await _syncCoordinator.PrepareForOpenAsync(fullPath, ct);
        _logger.Info($"동기화 완료: {fullPath} (상태: {result.State})");

        if (result.State == SyncState.Conflict)
        {
            ShowInfo($"충돌이 감지되어 온라인 사본을 별도 저장했습니다:\n{result.ConflictCopyPath}");
        }

        return 0;
    }

    public async Task<int> SyncAllAsync(CancellationToken ct)
    {
        var paths = _manifest.AllEntries.Keys.ToList();
        _logger.Info($"전체 동기화 시작: {paths.Count}개 파일");

        foreach (var path in paths)
        {
            if (!File.Exists(path))
            {
                _logger.Warn($"로컬에서 사라진 파일은 건너뜁니다: {path}");
                continue;
            }

            try
            {
                var result = await _syncCoordinator.PrepareForOpenAsync(path, ct);
                _logger.Info($"동기화됨: {path} (상태: {result.State})");
            }
            catch (Exception ex) when (GraphErrorHelper.IsResourceLocked(ex))
            {
                _logger.Warn($"온라인에서 편집 중이라 지금은 반영하지 못했습니다(다음 주기에 다시 시도): {path}");
            }
            catch (Exception ex)
            {
                _logger.Error($"동기화 실패: {path}", ex);
            }
        }

        return 0;
    }

    /// <summary>동기화 검토 창용: 쓰기 없이 상태만 조회한다.</summary>
    public Task<SyncDetection> DetectSyncStatusAsync(string filePath, CancellationToken ct) =>
        _syncCoordinator.DetectAsync(filePath, ct);

    /// <summary>동기화 검토 창용: 사용자가 고른 동작 하나만 실행한다.</summary>
    public Task<SyncResult> ApplySyncActionAsync(string filePath, SyncAction action, CancellationToken ct) =>
        _syncCoordinator.ApplyAsync(filePath, action, ct);

    /// <summary>GUI 프로세스에서만 실제 토스트 알림을 연결한다(헤드리스 오픈/CLI는 호출하지 않음).</summary>
    public void AttachActivityReporter(ISyncActivityReporter reporter) => _syncCoordinator.AttachActivityReporter(reporter);

    private static void OpenInBrowser(string url)
    {
        Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
    }

    private static void ShowError(string message) =>
        System.Windows.Forms.MessageBox.Show(message, AppBrand.Name,
            System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

    private static void ShowInfo(string message) =>
        System.Windows.Forms.MessageBox.Show(message, AppBrand.Name,
            System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
}
