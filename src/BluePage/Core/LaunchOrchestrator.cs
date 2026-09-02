using System.Diagnostics;
using Microsoft365OfficeWebLauncher.Cloud;
using Microsoft365OfficeWebLauncher.Config;
using Microsoft365OfficeWebLauncher.Logging;
using Microsoft365OfficeWebLauncher.OneDrive;
using Microsoft365OfficeWebLauncher.UI;

namespace Microsoft365OfficeWebLauncher.Core;

/// <summary>확장자 판별 → 동기화(업로드/다운로드) → 기본 브라우저로 webUrl 열기를 조율한다.</summary>
public sealed class LaunchOrchestrator
{
    private readonly DocumentTypeCatalog _catalog;
    private readonly SyncCoordinator _syncCoordinator;
    private readonly UploadManifest _manifest;
    private readonly FileLogger _logger;
    private readonly DeferredSyncRegistry _deferredSyncRegistry;
    private readonly AppConfig _config;

    public LaunchOrchestrator(
        DocumentTypeCatalog catalog,
        SyncCoordinator syncCoordinator,
        UploadManifest manifest,
        FileLogger logger,
        AppConfig config)
    {
        _catalog = catalog;
        _syncCoordinator = syncCoordinator;
        _manifest = manifest;
        _logger = logger;
        _config = config;
        _deferredSyncRegistry = new DeferredSyncRegistry();
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
        _deferredSyncRegistry.Resume(fullPath);

        if (string.Equals(Path.GetExtension(fullPath), ".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            ShowWarning("XLSM 파일을 웹에서 열 수 있지만 매크로는 실행되지 않습니다.");
        }

        var provider = SelectProvider(fullPath, appDefinition);
        if (provider is null)
        {
            _logger.Info($"사용자가 웹 Office 선택을 취소했습니다: {fullPath}");
            return 0;
        }

        SyncResult result;
        try
        {
            result = await _syncCoordinator.PrepareForOpenAsync(fullPath, ct, provider);
        }
        catch (Exception ex) when (GraphErrorHelper.IsResourceLocked(ex))
        {
            _logger.Warn($"Office Web 편집 잠금 감지, 문서 닫힘 대기 시작: {fullPath}");

            using var waitDialog = new WebDocumentUnlockDialog(
                fullPath,
                checkCt => _syncCoordinator.GetRemoteLockStateAsync(fullPath, checkCt),
                retryCt => _syncCoordinator.PrepareForOpenAsync(fullPath, retryCt, provider));
            var dialogResult = waitDialog.ShowDialog();

            if (dialogResult == DialogResult.Cancel ||
                (waitDialog.SyncResult is null && waitDialog.SyncError is null))
            {
                if (waitDialog.SynchronizationStopped)
                {
                    _deferredSyncRegistry.Defer(fullPath);
                }
                _logger.Info($"사용자가 웹 문서 닫힘 대기를 취소했습니다: {fullPath}");
                return 0;
            }

            if (waitDialog.SyncError is not null)
            {
                _logger.Error($"웹 문서 잠금 해제 후 동기화 실패: {fullPath}", waitDialog.SyncError);
                ShowError($"문서가 닫힌 후 동기화하는 중 오류가 발생했습니다.\n\n{waitDialog.SyncError.Message}");
                return 1;
            }

            result = waitDialog.SyncResult!;
            _logger.Info($"웹 문서 잠금 해제 후 동기화 완료: {fullPath} (상태: {result.State})");
        }
        catch (Exception ex)
        {
            _logger.Error($"동기화 실패: {fullPath}", ex);

            // 이전에 이미 업로드된 적 있는 파일이면, 반영은 실패했더라도 알고 있는 온라인 사본을 그대로 열어준다
            // (동기화 실패 = 문서를 아예 못 여는 것보다는 낫다).
            var existingEntry = _manifest.Get(fullPath);
            if (existingEntry is not null && !string.IsNullOrEmpty(existingEntry.WebUrl))
            {
                ShowInfo("온라인 동기화에 실패했습니다(자세한 내용은 로그 참고).\n" +
                         "일단 온라인의 최신 버전을 그대로 엽니다.");

                OpenInBrowser(existingEntry.WebUrl);
                _logger.Info($"동기화 실패 후 기존 온라인 사본을 열었습니다: {existingEntry.WebUrl}");
                return 0;
            }

            ShowError($"동기화에 실패해 문서를 열 수 없습니다.\n\n{ex.Message}");
            return 1;
        }

        if (result.State == SyncState.Skipped)
        {
            _logger.Info($"사용자 선택으로 문서를 열지 않고 동기화를 종료합니다: {fullPath}");
            return 0;
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
        _deferredSyncRegistry.Resume(fullPath);
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
            if (_deferredSyncRegistry.IsDeferred(path))
            {
                _logger.Debug($"사용자가 동기화하지 않기로 한 파일은 백그라운드에서 건너뜁니다: {path}");
                continue;
            }

            if (!File.Exists(path))
            {
                _deferredSyncRegistry.Resume(path);
                _logger.Warn($"로컬에서 사라진 파일은 건너뜁니다: {path}");
                continue;
            }

            try
            {
                var result = await _syncCoordinator.PrepareForOpenAsync(path, ct);
                _logger.Info($"동기화됨: {path} (상태: {result.State})");
                if (result.State == SyncState.Skipped)
                {
                    _deferredSyncRegistry.Defer(path);
                }
            }
            catch (Exception ex) when (GraphErrorHelper.IsResourceLocked(ex))
            {
                _logger.Warn($"백그라운드 동기화 중 Office Web 편집 잠금 감지, 문서 닫힘 대기 시작: {path}");

                using var waitDialog = new WebDocumentUnlockDialog(
                    path,
                    checkCt => _syncCoordinator.GetRemoteLockStateAsync(path, checkCt),
                    retryCt => _syncCoordinator.PrepareForOpenAsync(path, retryCt));
                var dialogResult = waitDialog.ShowDialog();

                if (dialogResult == DialogResult.Cancel ||
                    (waitDialog.SyncResult is null && waitDialog.SyncError is null))
                {
                    if (waitDialog.SynchronizationStopped)
                    {
                        _deferredSyncRegistry.Defer(path);
                    }
                    _logger.Info($"사용자가 백그라운드 동기화의 웹 문서 닫힘 대기를 취소했습니다: {path}");
                    continue;
                }

                if (waitDialog.SyncError is not null)
                {
                    _logger.Error($"웹 문서 잠금 해제 후 백그라운드 동기화 실패: {path}", waitDialog.SyncError);
                    ShowError($"문서가 닫힌 후 동기화하는 중 오류가 발생했습니다.\n\n{waitDialog.SyncError.Message}");
                    continue;
                }

                var retryResult = waitDialog.SyncResult!;
                _logger.Info($"웹 문서 잠금 해제 후 백그라운드 동기화 완료: {path} (상태: {retryResult.State})");

                if (retryResult.State == SyncState.Skipped)
                {
                    _deferredSyncRegistry.Defer(path);
                    continue;
                }

                if (retryResult.State == SyncState.Conflict)
                {
                    ShowInfo(
                        "로컬 파일과 Office Web의 온라인 사본이 모두 변경되어 자동으로 병합할 수 없습니다.\n" +
                        $"온라인 사본을 아래 경로에 별도로 저장했습니다. 두 파일을 확인 후 직접 병합해 주세요.\n\n{retryResult.ConflictCopyPath}");
                }

                OpenInBrowser(retryResult.WebUrl);
                _logger.Info($"잠금 해제 후 웹 문서를 다시 열었습니다: {retryResult.WebUrl}");
                ShowInfo($"동기화를 완료하고 웹 문서를 다시 열었습니다:\n{Path.GetFileName(path)}");
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
    public Task<SyncResult> ApplySyncActionAsync(string filePath, SyncAction action, CancellationToken ct)
    {
        _deferredSyncRegistry.Resume(Path.GetFullPath(filePath));
        return _syncCoordinator.ApplyAsync(filePath, action, ct);
    }

    /// <summary>GUI 프로세스에서만 실제 토스트 알림을 연결한다(헤드리스 오픈/CLI는 호출하지 않음).</summary>
    public void AttachActivityReporter(ISyncActivityReporter reporter) => _syncCoordinator.AttachActivityReporter(reporter);

    private CloudProvider? SelectProvider(string filePath, OfficeAppDefinition definition)
    {
        if (definition.SupportedProviders.Count == 1)
        {
            return definition.SupportedProviders.Single();
        }

        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        var preference = _config.DocumentProviderPreferences.TryGetValue(extension, out var perExtension) &&
                         !string.Equals(perExtension, "Default", StringComparison.OrdinalIgnoreCase)
            ? perExtension
            : _config.PreferredCloudProvider;

        if (CloudProviderNames.TryParse(preference, out var configured) && definition.Supports(configured))
        {
            return configured;
        }

        using var dialog = new CloudProviderDialog(filePath, definition.SupportedProviders);
        if (dialog.ShowDialog() != DialogResult.OK || dialog.SelectedProvider is null)
        {
            return null;
        }

        if (dialog.RememberAsDefault)
        {
            _config.PreferredCloudProvider = dialog.SelectedProvider.Value.ToString();
            ConfigLoader.Save(_config);
        }
        return dialog.SelectedProvider;
    }

    private static void OpenInBrowser(string url)
    {
        Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
    }

    private static void ShowError(string message) =>
        AppMessageDialog.Show(message, AppBrand.Name, AppMessageKind.Error);

    private static void ShowInfo(string message) =>
        AppMessageDialog.Show(message, AppBrand.Name, AppMessageKind.Information);

    private static void ShowWarning(string message) =>
        AppMessageDialog.Show(message, AppBrand.Name, AppMessageKind.Warning);
}
